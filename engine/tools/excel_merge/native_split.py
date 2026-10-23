# -*- coding: utf-8 -*-
"""Native XLSX filtering/splitting, with explicit decisions for ambiguous changes.

Coordinates in the public payload are absolute and one based. ``axis=columns``
exchanges the roles of rows and columns; no transpose is performed. Records outside
the selected data interval are retained. Blank records are never cleaned up.

Additional, deliberately explicit decisions:
* retain_sheets: other sheets the recipient is allowed to receive;
* retain_all_sheets: keep every original sheet in each output;
* shared_label_ranges: merged text labels allowed to be clipped;
* resolutions: {issue_id: "group_scope"} for references to a reduced range.
* blank_group: "separate" (default) or explicitly "previous"; changes ownership
  only, and never fills the source's blank cells;
* output_fields: absolute field coordinates to retain within the selected field
  interval. Fields outside that interval are preserved;
* include_example_rows: rows whose 序号/编号 column starts with "例" are template
  examples and are carried into every split output instead of being grouped;
* filters also support gt/lt/ge/le (numeric comparison) and oneof
  (comma separated values, any match).
Images and charts are preserved: anchors shift with the record mapping and chart
data references are rewritten; an object whose references become empty or whose
anchor is deleted is dropped with a warning. WPS in-cell pictures (cellimages /
DISPIMG) are re-injected as original package parts after each save. Pivot tables
and rich-data objects still block. Every output passes the shared XLSX package
validation before publication.
Formulas are kept as formulas; Excel recalculates them on opening the output.
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import OrderedDict
from copy import copy, deepcopy
from datetime import date, datetime
import hashlib
import json
import os
import posixpath
from pathlib import Path
import re
import tempfile
import warnings
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.worksheet.dimensions import DimensionHolder
from openpyxl.workbook.properties import CalcProperties

from .merge import Cancelled
from .compatibility import inspect_extensions
from .package_check import validate_xlsx

_CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$")
_COL = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_ROW = re.compile(r"^(\$?)([1-9]\d*)$")
_MONTH = re.compile(r"(?:(\d{4})\s*[-/.年]\s*)?(\d{1,2})(?:\s*月)?")
_RANGE_SEP = re.compile(r"\s*(?:至|到|~|～|—|–)\s*")
_INVALID_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_NAME = re.compile(r"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", re.I)
_DYNAMIC_FUNCS = re.compile(r"(?:^|[^A-Z0-9_.])(INDIRECT|OFFSET|CELL|ADDRESS)\s*\(", re.I)
_POSITION_FUNCS = re.compile(r"(?:^|[^A-Z0-9_.])(ROW|COLUMN|ROWS|COLUMNS)\s*\(", re.I)


class SplitBlocked(ValueError):
    """Export has unresolved preview issues; ``preview`` is safe to show in UI."""

    def __init__(self, preview):
        self.preview = preview
        super().__init__("拆分尚有待确认问题：" + "；".join(
            x["message"] for x in preview["issues"] if x.get("blocking", True))[:500])


def _json_value(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _stable(*parts):
    return hashlib.sha256(json.dumps(parts, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()[:16]


class _Issues:
    def __init__(self, payload):
        self.payload = payload
        self.items = OrderedDict()
        self.warnings = []

    def add(self, code, message, location="", group="", choices=None, blocking=True, **details):
        key = _stable(code, location, group)
        item = dict(id=key, code=code, message=message, location=location,
                    group=group, blocking=blocking, **details)
        if choices:
            item["choices"] = choices(key)
        self.items[key] = item
        return key

    def warn(self, text):
        if text not in self.warnings:
            self.warnings.append(text)

    def scope(self, location, group, reference):
        # 拆分后引用只能覆盖本结果保留的记录——原口径的记录已不在此文件中,
        # "按本组保留记录计算"是拆分语义下唯一合理的解释。自动执行并提示,
        # 不再作为需要人工决策的问题。
        self.warn(f"{location}：引用 {reference} 已按本结果保留的记录重新计算。")
        return True


class _Map:
    def __init__(self, selection, records):
        self.axis = selection["axis"]
        self.start, self.end = selection["data_start"], selection["data_end"]
        kept = set(records)
        self.removed = [i for i in range(self.start, self.end + 1) if i not in kept]

    def point(self, p):
        i = bisect_left(self.removed, p)
        if i < len(self.removed) and self.removed[i] == p:
            return None
        return p - i

    def boundary(self, p):
        """Position of a boundary, even if its original first cell was removed."""
        return p - bisect_left(self.removed, p)

    def segments(self, lo, hi):
        cursor = lo
        for p in self.removed[bisect_left(self.removed, lo):bisect_right(self.removed, hi)]:
            if cursor < p:
                yield cursor, p - 1
            cursor = p + 1
        if cursor <= hi:
            yield cursor, hi

    def partial(self, lo, hi):
        return bisect_right(self.removed, hi) > bisect_left(self.removed, lo)

    def coordinate(self, row, col):
        return (self.point(row), col) if self.axis == "rows" else (row, self.point(col))

    def rects(self, ref, split=False):
        rng = CellRange(ref)
        lo, hi = ((rng.min_row, rng.max_row) if self.axis == "rows"
                  else (rng.min_col, rng.max_col))
        spans = list(self.segments(lo, hi))
        if not spans:
            return []
        if not split:
            spans = [(spans[0][0], spans[-1][1])]
        result = []
        for start, end in spans:
            out = copy(rng)
            if self.axis == "rows":
                out.min_row, out.max_row = self.point(start), self.point(end)
                old_anchor = f"{get_column_letter(rng.min_col)}{start}"
            else:
                out.min_col, out.max_col = self.point(start), self.point(end)
                old_anchor = f"{get_column_letter(start)}{rng.min_row}"
            result.append((str(out), old_anchor))
        return result


def _shift_marker(marker, mapping):
    """锚点标记(0-based)随记录映射平移; 所在记录被删返回 None。"""
    if mapping.axis == "rows":
        new = mapping.point(marker.row + 1)
        if new is None:
            return None
        marker.row = new - 1
    else:
        new = mapping.point(marker.col + 1)
        if new is None:
            return None
        marker.col = new - 1
    return marker


def _shift_anchor(anchor, mapping):
    """from 标记被删 → False(对象丢弃); to 标记被删 → 夹到 from。"""
    if getattr(anchor, "_from", None) is not None and _shift_marker(anchor._from, mapping) is None:
        return False
    to = getattr(anchor, "to", None)
    if to is not None:
        if mapping.axis == "rows":
            new = mapping.point(to.row + 1)
            to.row = (new - 1) if new is not None else anchor._from.row
        else:
            new = mapping.point(to.col + 1)
            to.col = (new - 1) if new is not None else anchor._from.col
    return True


def _map_chart_refs(chart, owner_title, context):
    """改写指向被拆分工作表的图表数据引用; 引用区间删空 → False(丢弃)。"""
    mapping = context["mapping"]
    refs = []
    for series in chart.series or []:
        for holder in (getattr(series, "val", None), getattr(series, "cat", None),
                       getattr(series, "tx", None)):
            if holder is None:
                continue
            for attr in ("numRef", "strRef"):
                ref = getattr(holder, attr, None)
                if ref is not None and ref.f:
                    refs.append(ref)
    title = getattr(getattr(chart, "title", None), "tx", None)
    title_ref = getattr(title, "strRef", None) if title is not None else None
    if title_ref is not None and title_ref.f:
        refs.append(title_ref)
    for ref in refs:
        text = str(ref.f)
        sheet_part, _, rng = text.rpartition("!")
        if sheet_part:
            if sheet_part.strip("'\"") != context["target"]:
                continue                     # 引用其他页: 该页未被拆分, 原样保留
        elif owner_title != context["target"]:
            continue
        try:
            mapped, partial, empty = _map_reference(rng, mapping)
        except ValueError:
            continue                         # 不认识的引用格式: 保持原样
        if empty or mapped is None:
            return False
        if partial:
            context["issues"].warn(
                f"{owner_title} 的图表引用 {rng} 拆分后覆盖的记录集合已变化。")
        ref.f = f"{sheet_part}!{mapped}" if sheet_part else mapped
    return True


def _map_drawings(ws, mapping, context):
    """拆分页上的图片/图表: 锚点平移+引用改写; 无法保留的丢弃并提示。"""
    kept_charts = []
    for chart in list(ws._charts):
        ok = _shift_anchor(chart.anchor, mapping) and _map_chart_refs(chart, ws.title, context)
        if ok:
            kept_charts.append(chart)
        else:
            context["issues"].warn(f"{ws.title} 的一个图表（锚点或数据引用被删空）未随结果保留。")
    ws._charts = kept_charts
    kept_images = []
    for image in list(ws._images):
        if _shift_anchor(image.anchor, mapping):
            kept_images.append(image)
        else:
            context["issues"].warn(f"{ws.title} 的一张图片（锚定在被删记录上）未随结果保留。")
    ws._images = kept_images


def _selection(payload, ws):
    s = dict(payload.get("selection") or {})
    axis = s.get("axis", "rows")
    if axis not in {"rows", "columns"}:
        raise ValueError("selection.axis 必须为 rows 或 columns")
    end_axis = ws.max_row if axis == "rows" else ws.max_column
    end_field = ws.max_column if axis == "rows" else ws.max_row
    defaults = dict(header_start=1, header_end=1, data_start=2,
                    data_end=end_axis, field_start=1, field_end=end_field)
    for k, default in defaults.items():
        value = s.get(k, default)
        if isinstance(value, bool) or not isinstance(value, int) or value < 1:
            raise ValueError(f"selection.{k} 必须为正整数源坐标")
        s[k] = value
    s["axis"] = axis
    if not (s["header_start"] <= s["header_end"] < s["data_start"] <= s["data_end"]):
        raise ValueError("表头必须在数据区之前，且起止坐标不能倒置")
    if s["field_start"] > s["field_end"]:
        raise ValueError("字段起止坐标不能倒置")
    max_axis, max_field = (1048576, 16384) if axis == "rows" else (16384, 1048576)
    if s["data_end"] > max_axis or s["field_end"] > max_field:
        raise ValueError("所选坐标超出 XLSX 范围")
    return s


def _value(ws, row, col):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return ws.cell(rng.min_row, rng.min_col).value
    return cell.value


def _field(ws, s, record, index):
    if not isinstance(index, int) or not s["field_start"] <= index <= s["field_end"]:
        raise ValueError(f"字段坐标 {index} 不在已选择字段范围内")
    return _value(ws, record, index) if s["axis"] == "rows" else _value(ws, index, record)


def _one_month(value):
    if isinstance(value, (date, datetime)):
        return value.year, value.month
    if isinstance(value, int) and not isinstance(value, bool) and 1 <= value <= 12:
        return None, value
    text = str(value).strip()
    m = _MONTH.fullmatch(text)
    if not m:
        # Full ISO/slashed/Chinese dates are accepted without inventing a year.
        m = re.fullmatch(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T].*)?", text)
        if not m:
            raise ValueError(f"无法判断月份：{text}")
        date(int(m[1]), int(m[2]), int(m[3]))
    year, month = int(m[1]) if m[1] else None, int(m[2])
    if not 1 <= month <= 12:
        raise ValueError(f"月份超出 1–12：{text}")
    return year, month


def _number(value):
    """数值比较筛选的可比数字; 无法解析返回 None。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except ValueError:
            return None
    return None


def _month_span(start, end=None):
    if end is None and isinstance(start, str):
        parts = _RANGE_SEP.split(start.strip())
        if len(parts) == 1:
            # A bare 3-5月 is unambiguous; 2026-03 is a single month.
            m = re.fullmatch(r"(\d{1,2})\s*-\s*(\d{1,2})月?", start.strip())
            parts = list(m.groups()) if m else parts
        if len(parts) == 2:
            start, end = parts
    a = _one_month(start)
    b = _one_month(end) if end is not None else a
    if a[0] is None and b[0] is not None:
        a = b[0], a[1]
    if b[0] is None and a[0] is not None:
        b = a[0], b[1]
    inverted = (b < a) if a[0] is not None else (b[1] < a[1])
    if inverted:
        raise ValueError("月份区间倒置或跨年但未给出年份")
    return a, b


def _month_matches(value, target, end=None):
    a, b = _month_span(value, end)
    year, month = _one_month(target)
    if a[0] is None:
        return a[1] <= month <= b[1]
    if year is None:
        raise ValueError("源数据含年份，请选择带年份的目标月份")
    return a <= (year, month) <= b


def _group_records(ws, s, payload, issues):
    groups, selected = OrderedDict(), []
    mode = payload.get("group_mode", "combination")
    if mode not in {"combination", "recipients"}:
        raise ValueError("group_mode 必须为 combination 或 recipients")
    fields = payload.get("group_fields", [])
    output_fields = payload.get("output_fields")
    if output_fields is not None and (not isinstance(output_fields, list) or not output_fields or any(
            isinstance(x, bool) or not isinstance(x, int) or not s["field_start"] <= x <= s["field_end"]
            for x in output_fields)):
        raise ValueError("output_fields 必须是所选字段范围内的非空绝对坐标列表")
    sample_fields = sorted(set(output_fields)) if output_fields is not None else list(
        range(s["field_start"], s["field_end"] + 1))
    columns = []
    for index in sample_fields:
        parts = []
        for header in range(s["header_start"], s["header_end"] + 1):
            label = _value(ws, header, index) if s["axis"] == "rows" else _value(ws, index, header)
            if label is not None and str(label).strip() and (not parts or parts[-1] != str(label)):
                parts.append(str(label))
        columns.append(dict(index=index, label="/".join(parts) or f"字段 {index}"))
    blank_mode = payload.get("blank_group", "separate")
    if blank_mode not in {"separate", "previous"}:
        raise ValueError("blank_group 必须为 separate 或 previous")
    # 模板示例行: 序号/编号列以"例"开头的记录不参与分组, 每张结果都携带
    example_rows, example_set = [], set()
    if payload.get("include_example_rows"):
        serial_field = next((c["index"] for c in columns
                             if str(c["label"]).strip() in ("序号", "编号")), None)
        if serial_field is not None:
            for record in range(s["data_start"], s["data_end"] + 1):
                v = _field(ws, s, record, serial_field)
                if isinstance(v, str) and v.strip().startswith("例"):
                    example_rows.append(record)
            example_set = set(example_rows)
            if example_rows:
                issues.warn(f"示例行（“例”开头序号，共 {len(example_rows)} 行）将加入每张拆分结果。")
    previous, blank_records = {}, []
    for record in range(s["data_start"], s["data_end"] + 1):
        if record in example_set:
            continue
        vals = []
        record_blank = False
        for number, spec in enumerate(fields):
            kind = spec.get("kind")
            if kind == "field":
                value = _field(ws, s, record, spec["index"])
                if value in (None, ""):
                    record_blank = True
                    if blank_mode == "previous":
                        value = previous.get(number)
                else:
                    previous[number] = value
            elif kind == "cell":
                cell = ws[spec["ref"]]
                value = _value(ws, cell.row, cell.column)
            elif kind == "sheet":
                value = ws.title
            else:
                raise ValueError(f"不支持分组依据：{kind}")
            vals.append(_json_value(value))
        matched = True
        for spec in payload.get("filters", []):
            value = _field(ws, s, record, spec["field"])
            if isinstance(value, str) and value.startswith("="):
                issues.add("formula_filter", "筛选依据是公式，请改选有确定值的字段。", str(record))
                matched = False
                break
            op, target = spec.get("op", "equals"), spec.get("value")
            try:
                if op == "equals":
                    ok = str(value).strip() == str(target).strip() if value is not None else target is None
                elif op == "contains":
                    ok = value is not None and str(target) in str(value)
                elif op == "oneof":
                    parts = [p.strip() for p in re.split(r"[，,、]", str(target)) if p.strip()]
                    ok = value is not None and any(str(value).strip() == p for p in parts)
                elif op in {"gt", "lt", "ge", "le"}:
                    num, ref = _number(value), _number(target)
                    if num is None or ref is None:
                        raise ValueError(f"{op} 筛选要求字段值与填写值都是数字（当前值 {value!r}）")
                    ok = {"gt": num > ref, "lt": num < ref,
                          "ge": num >= ref, "le": num <= ref}[op]
                elif op in {"month", "interval"}:
                    if value is None or value == "":
                        ok = False
                    else:
                        end = _field(ws, s, record, spec["end_field"]) if op == "interval" else None
                        if op == "interval" and end in (None, ""):
                            raise ValueError("起止月筛选缺少结束月份")
                        ok = _month_matches(value, target, end)
                else:
                    raise ValueError(f"不支持筛选操作：{op}")
            except (ValueError, TypeError) as exc:
                issues.add("filter_value", f"记录 {record}：{exc}", str(record), field=spec["field"])
                ok = False
            if not ok:
                matched = False
                break
        if not matched:
            continue
        selected.append(record)
        if record_blank:
            blank_records.append(record)
        for value in vals:
            if isinstance(value, str) and value.startswith("="):
                issues.add("formula_group", "分组依据是公式，请改选有确定值的字段。", str(record))
        # Recipients means one copy for each selected field's complete value.
        # Splitting a cell's prose on punctuation would invent business semantics.
        keys = [(v,) for v in vals] if mode == "recipients" else [tuple(vals)]
        if not fields:
            keys = [("筛选结果",)]
        seen = set()
        for key in keys:
            gid = _stable(key)
            if gid in seen:
                continue
            seen.add(gid)
            if gid not in groups:
                label = " · ".join("（空白）" if v in (None, "") else str(v) for v in key)
                groups[gid] = dict(id=gid, label=label, count=0, records=[], sample=[], columns=columns)
            group = groups[gid]
            group["records"].append(record)
            group["count"] += 1
            if len(group["sample"]) < 5:
                group["sample"].append([_json_value(_field(ws, s, record, i))
                                        for i in sample_fields])
    if blank_records and blank_mode == "separate":
        issues.add("blank_groups", "分组字段存在空白，默认单独成组；若原表以空白表示沿用上一单位，可选择沿用。",
                   f"{ws.title}!分组空白", blocking=False, records=blank_records[:20],
                   choices=lambda k: [
                       dict(value="previous", label="空白沿用上一非空分组值（保留原格空白）",
                            patch={"blank_group": "previous"}),
                       dict(value="separate", label="空白单独成组", patch={"blank_group": "separate"})])
    elif blank_records:
        issues.warn("已按用户选择沿用上一非空分组值确定归属；原空白单元格保持空白。")
    if example_rows:
        example_sample = [[_json_value(_field(ws, s, r, i)) for i in sample_fields]
                          for r in example_rows[:2]]
        for group in groups.values():
            group["records"] = sorted(set(group["records"]) | example_set)
            group["count"] = len(group["records"])
            group["sample"] = (example_sample + group["sample"])[:5]
    return list(groups.values()), selected


def _tokens(text):
    value = str(text)
    return Tokenizer(value if value.startswith("=") else "=" + value).items


def _qualified(token, owner):
    if "!" not in token:
        return owner, token, ""
    prefix, ref = token.rsplit("!", 1)
    sheet = prefix[1:-1].replace("''", "'") if prefix.startswith("'") and prefix.endswith("'") else prefix
    return sheet, ref, prefix + "!"


def _names(wb, ws):
    result = dict(wb.defined_names)
    result.update(dict(ws.defined_names))
    return result


def _dependencies(text, owner, names, seen=None):
    result = []
    seen = set() if seen is None else seen
    try:
        tokens = _tokens(text)
    except Exception:
        return result
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        sheet, ref, prefix = _qualified(token.value, owner)
        if prefix:
            result.append((sheet, ref))
        elif ref in names and ref not in seen:
            seen.add(ref)
            result.extend(_dependencies(names[ref].attr_text, owner, names, seen))
    return result


def _formula_sources(ws):
    for cell in list(ws._cells.values()):
        if cell.data_type == "f":
            yield cell.value, cell.coordinate, "formula"
        if getattr(cell, "hyperlink", None):
            if cell.hyperlink.location:
                yield cell.hyperlink.location, cell.coordinate + " 超链接", "hyperlink"
            if cell.hyperlink.target and cell.hyperlink.target.startswith("#"):
                yield cell.hyperlink.target[1:], cell.coordinate + " 超链接", "hyperlink"
    for dv in ws.data_validations.dataValidation:
        for formula in (dv.formula1, dv.formula2):
            if formula:
                yield formula, f"数据验证 {dv.sqref}", "list" if dv.type == "list" else "rule"
    for area, rules in ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            for formula in rule.formula or []:
                yield formula, f"条件格式 {area.sqref}", "rule"
            for attr in ("colorScale", "dataBar", "iconSet"):
                for threshold in getattr(getattr(rule, attr, None), "cfvo", []) or []:
                    if threshold.type == "formula":
                        yield threshold.val, f"条件格式 {area.sqref}", "rule"


def _safe_config(ws, refs):
    """A hidden sheet is auto-carried only if all its values are in explicit lists."""
    if ws.sheet_state == "visible" or ws._charts or ws._images or ws.tables or ws.merged_cells:
        return False
    ranges = []
    try:
        for ref in refs:
            rng = CellRange(ref)
            if rng.min_col != rng.max_col and rng.min_row != rng.max_row:
                return False
            ranges.append(rng)
    except ValueError:
        return False
    cells = [c for c in ws._cells.values() if c.value is not None]
    return bool(cells) and all(c.data_type != "f" and any(
        r.min_row <= c.row <= r.max_row and r.min_col <= c.column <= r.max_col
        for r in ranges) for c in cells)


def _retained(wb, selected, payload, issues):
    explicit = payload.get("retain_sheets", [])
    if not isinstance(explicit, list) or any(x not in wb.sheetnames for x in explicit):
        raise ValueError("retain_sheets 必须是原文件中存在的工作表名称列表")
    if payload.get("retain_all_sheets"):
        issues.warn("已按要求保留原文件的全部工作表（未参与筛选）。")
        return list(wb.sheetnames)
    retained, queue = {selected.title, *explicit}, [selected.title, *explicit]
    while queue:
        owner = queue.pop(0)
        ws = wb[owner]
        dep_refs = {}
        for formula, location, kind in _formula_sources(ws):
            for target, ref in _dependencies(formula, owner, _names(wb, ws)):
                dep_refs.setdefault(target, []).append((ref, kind, location))
        for target, entries in dep_refs.items():
            if target in retained:
                continue
            if target in wb.sheetnames and all(x[1] == "list" for x in entries) and _safe_config(wb[target], [x[0] for x in entries]):
                retained.add(target)
                queue.append(target)
                issues.warn(f"保留被下拉列表引用的隐藏配置页“{target}”；该页全部有效内容均属于被引用列表。")
            elif target in wb.sheetnames:
                issues.add("dependency_sheet", f"“{owner}”引用“{target}”，确认该页可以随结果发出后才能保留。",
                           target, choices=lambda k, target=target: [
                               dict(value="retain", label=f"允许结果包含“{target}”",
                                    patch={"retain_sheets": sorted(set(explicit) | {target})})],
                           sheet=target)
                pending = getattr(issues, "dependency_pending", None)
                if pending is None:
                    pending = issues.dependency_pending = set()
                pending.add(target)
            else:
                issues.add("external_reference", f"存在外部或无法识别的工作表引用：{target}", f"{owner}!{entries[0][2]}")
    dropped = [name for name in wb.sheetnames if name not in retained]
    if dropped:
        issues.warn("以下未选择且未授权的工作表不会随结果发出：" + "、".join(dropped))
    if explicit:
        issues.warn("用户已授权保留以下完整工作表（不参与筛选）：" + "、".join(explicit))
    return [name for name in wb.sheetnames if name in retained]


def _map_reference(ref, mapping, partial_ok=False):
    """Return (mapped reference, partial, empty); preserve $ flags."""
    parts = ref.split(":")
    if len(parts) > 2:
        raise ValueError("不支持三维或多段引用")
    kind = "cell" if _CELL.fullmatch(parts[0]) else "col" if _COL.fullmatch(parts[0]) else "row" if _ROW.fullmatch(parts[0]) else None
    regex = {"cell": _CELL, "col": _COL, "row": _ROW}.get(kind)
    if regex is None or any(not regex.fullmatch(p) for p in parts):
        raise ValueError("不支持此引用格式")
    if len(parts) == 1 and kind != "cell":
        raise ValueError("不是单元格引用")
    matches = [regex.fullmatch(p) for p in parts]
    relevant = kind == "cell" or (kind == "row") == (mapping.axis == "rows")
    if not relevant:
        # SUM(B:B) on a row-filtered sheet still changes its population, even
        # though the spelling of the whole-column reference does not change.
        return ref, bool(mapping.removed), False
    def pos(m):
        if kind == "cell":
            return int(m[4]) if mapping.axis == "rows" else column_index_from_string(m[2])
        return int(m[2]) if kind == "row" else column_index_from_string(m[2])
    a, b = pos(matches[0]), pos(matches[-1])
    if b < a:
        raise ValueError("反向区域引用")
    spans = list(mapping.segments(a, b))
    partial = mapping.partial(a, b)
    if not spans:
        return None, partial, True
    values = [mapping.point(spans[0][0]), mapping.point(spans[-1][1])]
    def build(m, p):
        if kind == "cell":
            return (f"{m[1]}{m[2]}{m[3]}{p}" if mapping.axis == "rows"
                    else f"{m[1]}{get_column_letter(p)}{m[3]}{m[4]}")
        return m[1] + (str(p) if kind == "row" else get_column_letter(p))
    return ":".join(build(m, values[i]) for i, m in enumerate(matches)), partial, False


def _rewrite(formula, owner, target, mapping, wb, retained, issues, location, group,
             metadata=False, seen=None):
    if formula is None:
        return formula
    if not isinstance(formula, str):
        issues.add("array_formula", "数组/数据表公式暂不能安全移动，请调整选择范围。", location, group)
        return formula
    if _DYNAMIC_FUNCS.search(formula) and mapping.removed:
        issues.add("dynamic_reference", "公式通过地址或行列位置动态取值，暂不能安全判断拆分后的含义。", location, group)
        return formula
    if _POSITION_FUNCS.search(formula) and mapping.removed:
        issues.scope(location, group, formula)
    if re.search(r'\bHYPERLINK\s*\(\s*"#', formula, re.I) and mapping.removed:
        issues.add("hyperlink_formula", "公式中的文本超链接地址需要单独调整，暂不能安全移动。", location, group)
    has_equals = formula.startswith("=")
    try:
        tokens = _tokens(formula)
    except Exception:
        issues.add("formula_parse", "无法可靠解析公式，已停止导出。", location, group)
        return formula
    out, names = [], _names(wb, wb[owner])
    seen = set() if seen is None else set(seen)
    for token in tokens:
        value = token.value
        if token.type == "OPERAND" and token.subtype == "ERROR" and value == "#REF!":
            issues.add("broken_reference", "原公式含 #REF!，请先确认引用。", location, group)
        if token.type == "OPERAND" and token.subtype == "RANGE":
            sheet, ref, prefix = _qualified(value, owner)
            if not prefix and ref in names:
                if ref not in seen:
                    seen.add(ref)
                    _rewrite(names[ref].attr_text, owner, target, mapping, wb, retained,
                             issues, location, group, metadata=metadata, seen=seen)
                out.append(value)
                continue
            if sheet not in retained:
                if sheet not in getattr(issues, "dependency_pending", set()):
                    issues.add("missing_sheet", f"公式依赖未保留的工作表“{sheet}”。", location, group)
            elif sheet == target:
                try:
                    mapped, partial, empty = _map_reference(ref, mapping)
                except ValueError:
                    issues.add("unsupported_reference", f"暂不能安全移动引用：{value}", location, group)
                else:
                    if empty:
                        issues.add("missing_reference", f"引用 {value} 的全部单元格将被移除，不能替换成空值。", location, group)
                    elif partial and not metadata and ":" not in ref:
                        issues.add("missing_reference", f"引用 {value} 的单元格将被移除。", location, group)
                    elif partial and not metadata:
                        if issues.scope(location, group, value):
                            value = prefix + mapped
                    else:
                        value = prefix + mapped
        out.append(value)
    return ("=" if has_equals else "") + "".join(out)


def _shift_rule(formula, source_anchor, segment_anchor, **context):
    if formula is None:
        return None
    text = str(formula)
    has_equals = text.startswith("=")
    try:
        shifted = Translator(text if has_equals else "=" + text, origin=source_anchor).translate_formula(segment_anchor)
    except Exception:
        context["issues"].add("rule_formula", "无法可靠平移数据验证/条件格式公式。", context["location"], context["group"])
        return formula
    result = _rewrite(shifted, **context)
    return result if has_equals else result[1:]


def _rule_formula(formula, base, anchor, mapped, context):
    """Check both ends, not just the rule's top-left relative reference.

    A continuous rule segment can still reference a removed record *outside*
    that segment. The coordinate map is monotone, so a change in its displacement
    appears at the endpoint; reject instead of assigning a misleading rule.
    """
    first = _shift_rule(formula, base, anchor, **context)
    if formula is None:
        return first
    rng = CellRange(mapped)
    original = CellRange(anchor)
    last = f"{get_column_letter(original.min_col + rng.max_col - rng.min_col)}{original.min_row + rng.max_row - rng.min_row}"
    if last == anchor:
        return first
    actual_last = _shift_rule(formula, base, last, **context)
    try:
        first_text = str(first)
        expected = Translator(first_text if first_text.startswith("=") else "=" + first_text,
                              origin=f"{get_column_letter(rng.min_col)}{rng.min_row}").translate_formula(
                                  f"{get_column_letter(rng.max_col)}{rng.max_row}")
        actual = str(actual_last)
        if expected != (actual if actual.startswith("=") else "=" + actual):
            context["issues"].add(
                "rule_relative_reference",
                "此验证/条件格式在区域内部的相对引用跨越被移除记录，不能按单一平移保留。",
                context["location"], context["group"])
    except Exception:
        context["issues"].add("rule_formula", "无法校验规则区域末端的相对引用。",
                             context["location"], context["group"])
    return first


def _population_rule(rule, area, mapping, context):
    """Keep a statistical CF rule on one population, not one rule per record run."""
    ranges = sorted(area.sqref.ranges, key=lambda r: (r.min_row, r.min_col))
    mapped = [out for rng in ranges for out, _ in mapping.rects(str(rng))]
    if not mapped:
        return None, None
    new = deepcopy(rule)
    for attr in ("colorScale", "dataBar", "iconSet"):
        for threshold in getattr(getattr(new, attr, None), "cfvo", []) or []:
            if threshold.type != "formula":
                continue
            # Relative CF thresholds have Excel-specific population semantics.
            # Constants and absolute/name references can be preserved exactly.
            try:
                refs = [t.value.rsplit("!", 1)[-1] for t in _tokens(threshold.val)
                        if t.type == "OPERAND" and t.subtype == "RANGE"]
            except Exception:
                refs = []
                context["issues"].add("rule_formula", "无法解析统计条件格式阈值。",
                                      context["location"], context["group"])
            for ref in refs:
                for part in ref.split(":"):
                    cell = _CELL.fullmatch(part)
                    if cell and not (cell[3] if mapping.axis == "rows" else cell[1]):
                        context["issues"].add("relative_statistical_rule",
                                              "统计条件格式阈值含相对引用，暂不能安全改变统计区域。",
                                              context["location"], context["group"])
            threshold.val = _rewrite(str(threshold.val), **context)
    return " ".join(mapped), new


def _transform_rules(ws, mapping, context):
    old_dvs = list(ws.data_validations.dataValidation)
    ws.data_validations.dataValidation = []
    for number, dv in enumerate(old_dvs):
        ranges = sorted(dv.sqref.ranges, key=lambda r: (r.min_row, r.min_col))
        if not ranges:
            continue
        base = f"{get_column_letter(ranges[0].min_col)}{ranges[0].min_row}"
        for rng in ranges:
            for mapped, anchor in mapping.rects(str(rng), split=True):
                new = deepcopy(dv)
                new.sqref = mapped
                ctx = dict(context, location=f"{ws.title}!验证{number + 1}:{rng}")
                new.formula1 = _rule_formula(dv.formula1, base, anchor, mapped, ctx)
                new.formula2 = _rule_formula(dv.formula2, base, anchor, mapped, ctx)
                ws.add_data_validation(new)
    old_cf = list(ws.conditional_formatting._cf_rules.items())
    ws.conditional_formatting = ConditionalFormattingList()
    for number, (area, rules) in enumerate(old_cf):
        ranges = sorted(area.sqref.ranges, key=lambda r: (r.min_row, r.min_col))
        if not ranges:
            continue
        base = f"{get_column_letter(ranges[0].min_col)}{ranges[0].min_row}"
        for rule in rules:
            if rule.type in {"colorScale", "dataBar", "iconSet", "top10", "aboveAverage",
                             "duplicateValues", "uniqueValues"}:
                ctx = dict(context, location=f"{ws.title}!条件格式{number + 1}:{area.sqref}")
                mapped, new = _population_rule(rule, area, mapping, ctx)
                if mapped:
                    ws.conditional_formatting.add(mapped, new)
                continue
            for rng in ranges:
                for mapped, anchor in mapping.rects(str(rng), split=True):
                    new = deepcopy(rule)
                    ctx = dict(context, location=f"{ws.title}!条件格式{number + 1}:{rng}")
                    new.formula = [_rule_formula(f, base, anchor, mapped, ctx) for f in (rule.formula or [])]
                    ws.conditional_formatting.add(mapped, new)


def _map_dimensions(ws, mapping):
    if mapping.axis == "rows":
        entries = list(ws.row_dimensions.items())
        ws.row_dimensions = DimensionHolder(worksheet=ws, default_factory=ws._add_row)
        for index, dim in entries:
            new_index = mapping.point(index)
            if new_index is not None:
                new = copy(dim)
                new.index = new_index
                new.parent = ws
                ws.row_dimensions[new_index] = new
    else:
        entries = list(ws.column_dimensions.items())
        ws.column_dimensions = DimensionHolder(worksheet=ws, default_factory=ws._add_column)
        for key, dim in entries:
            start = dim.min or column_index_from_string(key)
            end = dim.max or start
            for a, b in mapping.segments(start, end):
                new = copy(dim)
                new.min, new.max = mapping.point(a), mapping.point(b)
                new.index = get_column_letter(new.min)
                new.parent = ws
                ws.column_dimensions[new.index] = new


def _metadata(ws, mapping, context):
    def mapped_ref(ref):
        result = mapping.rects(str(ref))
        return result[0][0] if result else None
    for view in ws.views.sheetView:
        if view.pane:
            pane = view.pane
            if pane.topLeftCell:
                c = ws[pane.topLeftCell]
                row = mapping.boundary(c.row) if mapping.axis == "rows" else c.row
                col = mapping.boundary(c.column) if mapping.axis == "columns" else c.column
                pane.topLeftCell = f"{get_column_letter(col)}{row}"
            # A split window uses physical units; only a frozen pane uses record
            # counts. Calling freeze_panes=... would turn an ordinary split into
            # a frozen window and discard the original pane state.
            if pane.state in {"frozen", "frozenSplit"}:
                if mapping.axis == "rows" and pane.ySplit:
                    pane.ySplit = max(0, mapping.boundary(int(pane.ySplit) + 1) - 1)
                if mapping.axis == "columns" and pane.xSplit:
                    pane.xSplit = max(0, mapping.boundary(int(pane.xSplit) + 1) - 1)
        if view.topLeftCell:
            c = ws[view.topLeftCell]
            row = mapping.boundary(c.row) if mapping.axis == "rows" else c.row
            col = mapping.boundary(c.column) if mapping.axis == "columns" else c.column
            view.topLeftCell = f"{get_column_letter(col)}{row}"
        for sel in view.selection:
            if sel.sqref:
                refs = [mapped_ref(str(r)) for r in MultiCellRange(sel.sqref).ranges]
                sel.sqref = " ".join(r for r in refs if r) or "A1"
            if sel.activeCell:
                sel.activeCell = mapped_ref(sel.activeCell) or sel.sqref.split()[0].split(":")[0]
    if ws.print_area:
        refs = [ref for _, ref in _dependencies(str(ws.print_area), ws.title, {})]
        ws.print_area = [r for ref in refs if (r := mapped_ref(ref))]
    if ws.print_title_rows and mapping.axis == "rows":
        mapped, _, empty = _map_reference(ws.print_title_rows, mapping, True)
        ws.print_title_rows = None if empty else mapped
    if ws.print_title_cols and mapping.axis == "columns":
        mapped, _, empty = _map_reference(ws.print_title_cols, mapping, True)
        ws.print_title_cols = None if empty else mapped
    if ws.auto_filter.ref:
        original = CellRange(ws.auto_filter.ref)
        ws.auto_filter.ref = mapped_ref(ws.auto_filter.ref)
        if mapping.axis == "columns" and ws.auto_filter.ref:
            start = CellRange(ws.auto_filter.ref).min_col
            filters = []
            for f in ws.auto_filter.filterColumn:
                col = mapping.point(original.min_col + f.colId)
                if col is not None:
                    f.colId = col - start
                    filters.append(f)
            ws.auto_filter.filterColumn = filters
        if ws.auto_filter.sortState:
            state = ws.auto_filter.sortState
            mapped = mapped_ref(state.ref)
            if not mapped:
                ws.auto_filter.sortState = None
            else:
                state.ref = mapped
                kept = []
                for condition in state.sortCondition:
                    ref = mapped_ref(condition.ref)
                    if ref:
                        condition.ref = ref
                        kept.append(condition)
                state.sortCondition = kept
    breaks = ws.row_breaks if mapping.axis == "rows" else ws.col_breaks
    for br in breaks.brk:
        br.id = max(0, mapping.boundary(br.id + 1) - 1)
    for table in list(ws.tables.values()):
        rng = CellRange(table.ref)
        lo, hi = (rng.min_row, rng.max_row) if mapping.axis == "rows" else (rng.min_col, rng.max_col)
        if mapping.axis == "columns" and mapping.partial(lo, hi):
            context["issues"].add("table_columns", "拆分会删去 Excel 表格的字段，暂不能保留其结构化引用。", f"{ws.title}!{table.ref}", context["group"])
            continue
        mapped = mapped_ref(table.ref)
        if mapped is None or (mapping.axis == "rows" and mapping.point(rng.min_row) is None):
            context["issues"].add("table_header", "拆分将移除 Excel 表格的表头，不能安全保留表格对象。", f"{ws.title}!{table.ref}", context["group"])
            continue
        table.ref = mapped
        if table.autoFilter:
            table.autoFilter.ref = mapped_ref(table.autoFilter.ref)
        if table.sortState:
            context["issues"].add("table_sort", "Excel 表格含排序状态，暂不能安全移动其排序区域。", f"{ws.title}!{table.ref}", context["group"])
        if any(c.calculatedColumnFormula or c.totalsRowFormula for c in table.tableColumns):
            context["issues"].add("table_formula", "Excel 表格含结构化计算公式，暂不能安全调整。", f"{ws.title}!{table.ref}", context["group"])


def _cached_value_reader(path):
    """惰性打开 data_only 版: 公式引用被删空时用"原计算值"回退。

    真实 Excel/WPS 保存的文件带缓存值; openpyxl 生成的文件没有(返回 None,
    回退不生效, 保持原有阻断)。
    """
    store = {}

    def read(sheet, coord):
        if "wb" not in store:
            store["wb"] = load_workbook(path, data_only=True)
        try:
            return store["wb"][sheet][coord].value
        except Exception:
            return None

    def close():
        book = store.pop("wb", None)
        if book is not None:
            try:
                book.close()
            except Exception:
                pass

    return read, close


def _transform(wb, target, s, group, retained, payload, issues, metadata_labels=(), record_filter=None):
    ws = wb[target]
    mapping = _Map(s, group["records"])
    context = dict(owner=target, target=target, mapping=mapping, wb=wb, retained=retained,
                   issues=issues, group=group["id"])
    cached_read, cached_close = _cached_value_reader(payload["path"])
    shared = {str(CellRange(r)) for r in payload.get("shared_label_ranges", [])}
    merges = []
    for rng in list(ws.merged_cells.ranges):
        mapped = mapping.rects(str(rng))
        if not mapped:
            continue
        lo, hi = (rng.min_row, rng.max_row) if mapping.axis == "rows" else (rng.min_col, rng.max_col)
        cell = ws.cell(rng.min_row, rng.min_col)
        if mapping.partial(lo, hi):
            is_text = isinstance(cell.value, str) and not cell.value.startswith("=")
            header_label = str(rng) in metadata_labels and cell.data_type != "f"
            if not header_label and is_text:
                # 文本合并区跨记录 = 各记录共用的标签(拆分语义下唯一合理解释),
                # 自动按共享标签保留在各结果中
                issues.warn(f"合并区 {rng} 跨越保留与移除记录，已按共用标签保留在各结果中。")
            elif not header_label:
                issues.add("merged_record", f"合并区 {rng} 跨越保留与移除记录且内容不是文本标签，请确认这些数据属于哪条记录。",
                           f"{target}!{rng}", group["id"], range=str(rng))
        merges.append((mapped[0][0], copy(cell)))
    # Rewrite before replacing cells, so all formulas resolve the original names.
    for other in retained:
        for cell in list(wb[other]._cells.values()):
            if cell.data_type != "f":
                continue
            if other == target and None in mapping.coordinate(cell.row, cell.column):
                continue
            if other == target and record_filter:
                record_axis, kept_records = record_filter
                pos = cell.row if record_axis["axis"] == "rows" else cell.column
                if record_axis["data_start"] <= pos <= record_axis["data_end"] and pos not in kept_records:
                    continue
            location = f"{other}!{cell.coordinate}"
            before = set(issues.items)
            new_value = _rewrite(cell.value, **dict(context, owner=other, location=location))
            fatal = [k for k in issues.items if k not in before and issues.items[k].get("code")
                     in {"missing_reference", "unsupported_reference"}]
            if fatal:
                cached = cached_read(other, cell.coordinate)
                if cached not in (None, ""):
                    for k in fatal:
                        item = issues.items[k]
                        item.update(blocking=False, resolved=True)
                        item["message"] += " 已自动转为原计算值。"
                    issues.warn(f"{location} 的公式引用不在此结果中，已转为原计算值（拆分前显示的值）。")
                    new_value = cached
            cell.value = new_value
    _transform_rules(ws, mapping, context)
    _metadata(ws, mapping, context)
    _map_dimensions(ws, mapping)
    _map_drawings(ws, mapping, context)
    if mapping.axis == "rows":
        # 序号列连续重排(每组从1): 数值序号与公式序号(=ROW()-N 在行压缩后会
        # 漂移甚至为负)一律重排为连续数字; 文本序号(如"例1")保留原值不计数
        serial_field = None
        for i in range(s["field_start"], s["field_end"] + 1):
            if str(_value(ws, s["header_start"], i) or "").strip() in ("序号", "编号"):
                serial_field = i
                break
        if serial_field is not None:
            counter, renumbered = 0, 0
            for record in sorted(group["records"]):
                cell = ws._cells.get((record, serial_field))
                v = cell.value if cell is not None else None
                if v is None or v == "":
                    continue
                if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit())                         or (isinstance(v, str) and v.strip().startswith("=")):
                    counter += 1
                    if cell is not None and v != counter:
                        cell.value = counter
                    renumbered += 1
            if counter:
                issues.warn("序号列已按本组记录从 1 连续重排（文本序号保留原值）。")
    for owner in retained:
        # Rules on retained pages can also refer back to the filtered sheet.
        if owner != target:
            for chart in wb[owner]._charts or []:
                _map_chart_refs(chart, owner, context)
            for dv in wb[owner].data_validations.dataValidation:
                for attr in ("formula1", "formula2"):
                    setattr(dv, attr, _rewrite(getattr(dv, attr), **dict(
                        context, owner=owner, location=f"{owner}!验证 {dv.sqref}")))
            for area, rules in wb[owner].conditional_formatting._cf_rules.items():
                for rule in rules:
                    rule.formula = [_rewrite(f, **dict(context, owner=owner,
                                    location=f"{owner}!条件格式 {area.sqref}")) for f in (rule.formula or [])]
                    for attr in ("colorScale", "dataBar", "iconSet"):
                        for threshold in getattr(getattr(rule, attr, None), "cfvo", []) or []:
                            if threshold.type == "formula":
                                threshold.val = _rewrite(str(threshold.val), **dict(
                                    context, owner=owner, location=f"{owner}!条件格式 {area.sqref}"))
            for cell in wb[owner]._cells.values():
                if getattr(cell, "hyperlink", None):
                    _rewrite_link(cell, payload, dict(context, owner=owner,
                                  location=f"{owner}!{cell.coordinate} 超链接"))
    for owner, collection in [(target, wb.defined_names)] + [(x, wb[x].defined_names) for x in retained]:
        for name, definition in list(collection.items()):
            dependencies = _dependencies(definition.attr_text, owner, {})
            if any(sheet not in retained for sheet, _ in dependencies):
                del collection[name]
                issues.warn(f"命名区域“{name}”仅供未输出工作表使用，未随结果保留。")
                continue
            definition.attr_text = _rewrite(definition.attr_text, **dict(
                context, owner=owner, location=f"命名区域 {name}", metadata=True))
    original = list(ws._cells.values())
    ws._cells = {}
    ws.merged_cells = MultiCellRange()
    for cell in original:
        row, col = mapping.coordinate(cell.row, cell.column)
        if row is None or col is None:
            continue
        new = copy(cell)
        new.row, new.column = row, col
        if getattr(new, "hyperlink", None):
            _rewrite_link(new, payload, dict(context, location=f"{target}!{cell.coordinate} 超链接"))
        ws._cells[(row, col)] = new
    for ref, old_anchor in merges:
        rng = CellRange(ref)
        anchor = ws.cell(rng.min_row, rng.min_col)
        if isinstance(anchor, MergedCell):
            old_anchor.row, old_anchor.column = rng.min_row, rng.min_col
            ws._cells[(rng.min_row, rng.min_col)] = old_anchor
        ws.merge_cells(ref)
    for name in list(wb.sheetnames):
        if name not in retained:
            del wb[name]
    if not any(w.sheet_state == "visible" for w in wb):
        issues.add("hidden_output", "输出中没有可见工作表，请选择可见数据页。", target, group["id"])
    cached_close()
    wb.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
    ws._current_row = max((row for row, _ in ws._cells), default=0)
    return wb


def _rewrite_link(cell, payload, context):
    cell._hyperlink = copy(cell.hyperlink)
    cell._hyperlink.ref = cell.coordinate
    if cell._hyperlink.location:
        cell._hyperlink.location = _rewrite(cell._hyperlink.location, **context)
    if cell._hyperlink.target and cell._hyperlink.target.startswith("#"):
        cell._hyperlink.target = "#" + _rewrite(cell._hyperlink.target[1:], **context)
    elif cell._hyperlink.target and not re.match(r"^[A-Za-z][A-Za-z0-9+.-]*:", cell._hyperlink.target):
        link_path = Path(cell._hyperlink.target)
        if not link_path.is_absolute():
            cell._hyperlink.target = str((Path(payload["path"]).resolve().parent / link_path).resolve())


def _render_group(wb, sheet, selection, group, retained, payload, issues):
    fields = payload.get("output_fields")
    second_payload = payload
    if fields is not None:
        if not isinstance(fields, list) or not fields or any(
                isinstance(x, bool) or not isinstance(x, int) or not
                selection["field_start"] <= x <= selection["field_end"] for x in fields):
            raise ValueError("output_fields 必须是所选字段范围内的非空绝对坐标列表")
        field_selection = dict(
            axis="columns" if selection["axis"] == "rows" else "rows",
            data_start=selection["field_start"], data_end=selection["field_end"])
        field_group = dict(group, records=sorted(set(fields)))
        header_labels = []
        for rng in wb[sheet].merged_cells.ranges:
            end = rng.max_row if selection["axis"] == "rows" else rng.max_col
            if end < selection["data_start"]:
                header_labels.append(str(rng))
        mapping = _Map(field_selection, fields)
        second_payload = dict(payload, shared_label_ranges=[
            mapped for ref in payload.get("shared_label_ranges", [])
            for mapped, _ in mapping.rects(ref)])
        _transform(wb, sheet, field_selection, field_group, retained, payload, issues,
                   metadata_labels=header_labels, record_filter=(selection, set(group["records"])))
        issues.warn("输出仅保留所选字段范围中的指定字段；范围外的前置或附注字段继续保留。")
    return _transform(wb, sheet, selection, group, retained, second_payload, issues)


def _object_issues(wb, retained, issues, path):
    for name in retained:
        ws = wb[name]
        objects = []
        if ws._pivots:
            objects.append("数据透视表")
        if objects:
            issues.add("unsupported_object", f"“{name}”含{'、'.join(objects)}，当前不能保证拆分后引用与位置不变。", name)
        # 图片/图表改为保留: 锚点随记录平移、图表引用同步改写; 删空的丢弃并提示
        if ws._charts or ws._images:
            issues.warn(f"“{name}”的图片/图表将随结果保留：位置随拆分平移，图表数据引用同步改写；引用被删空的图表会丢弃。")
        if any(getattr(c, "comment", None) for c in ws._cells.values()):
            issues.warn(f"“{name}”的普通批注文字与作者保留；批注框外观由 Excel 重新生成。")
    # openpyxl warns about, or silently omits, some XML extensions. Detect them
    # before saving rather than claiming a faithful result after loss.
    with ZipFile(path) as archive:
        unsupported = [n for n in archive.namelist() if n.startswith((
            "xl/externalLinks/", "xl/activeX/", "xl/embeddings/", "xl/slicer",
            "xl/threadedComments/", "xl/persons/", "xl/ctrlProps/", "xl/richData/"))
            or n in {"xl/connections.xml", "xl/vbaProject.bin", "xl/metadata.xml"}]
        if unsupported:
            issues.add("unsupported_package", "文件含当前无法完整保留的外部链接、控件或扩展对象。", unsupported[0])
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        paths = {x.attrib["Id"]: x.attrib["Target"] for x in rels}
        for sheet in book.findall("m:sheets/m:sheet", ns):
            if sheet.attrib["name"] not in retained:
                continue
            target = paths[sheet.attrib["{" + ns["r"] + "}id"]]
            target = target.lstrip("/") if target.startswith("/") else "xl/" + target
            xml = ET.fromstring(archive.read(target))
            tags = {x.tag.rsplit("}", 1)[-1] for x in xml.iter()}
            if tags & {"AlternateContent", "oleObjects", "controls", "legacyDrawingHF"}:
                issues.add("unsupported_sheet_xml", "工作表含当前无法无损保留的扩展格式或对象。", sheet.attrib["name"])


def _load(path, issues):
    # Extension loss is an explicit compatibility notice. Unrelated read
    # warnings and actual parsing failures still require attention.
    extension_notices = []
    with ZipFile(path) as archive:
        for name in archive.namelist():
            lower = name.lower()
            if lower.startswith("xl/printersettings/"):
                issues.warn("打印机专属二进制设置不保留；常规纸张、方向和页边距继续保留。")
            if lower.startswith("xl/") and lower.endswith(".xml"):
                raw = archive.read(name)
                try:
                    extension_notices.extend(inspect_extensions(name, raw))
                except ValueError as exc:
                    code = "unsupported_sheet_xml" if lower.startswith("xl/worksheets/") else "unsupported_package"
                    issues.add(code, str(exc), name)
    for notice in extension_notices:
        issues.warn(notice)
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always", UserWarning)
        wb = load_workbook(path, data_only=False, keep_links=True, rich_text=True)
    for warning in caught:
        message = str(warning.message)
        if (issubclass(warning.category, UserWarning)
                and message.endswith(" extension is not supported and will be removed")):
            issues.warn("读取兼容提示：" + message)
        else:
            issues.add("read_warning", "读取时发现尚未确认的兼容问题：" + message, str(path))
    return wb


def _package_extras(path):
    """WPS 单元格内嵌图片(xl/cellimages*)部件: openpyxl 不建模, 保存会丢失。

    提取原文件中的 cellimages 部件、其关系文件与引用的媒体文件, 以及
    内容清单/工作簿关系中对应的片段, 供导出后原样回注。__ 开头的键为
    清单片段而非包部件。
    """
    extras = {}
    with ZipFile(path) as archive:
        names = set(archive.namelist())
        core = [n for n in sorted(names) if n.startswith("xl/cellimages")]
        if not core:
            return extras
        dependents = set()
        for part in core:
            extras[part] = archive.read(part)
            stem = part.rsplit("/", 1)
            rels = stem[0] + "/_rels/" + stem[1] + ".rels"
            if rels in names:
                extras[rels] = archive.read(rels)
                owner_dir = posixpath.dirname(part)
                for rel in ET.fromstring(extras[rels]):
                    if rel.get("TargetMode") == "External":
                        continue
                    target = rel.get("Target", "")
                    resolved = posixpath.normpath(
                        target.lstrip("/") if target.startswith("/")
                        else posixpath.join(owner_dir, target))
                    if resolved in names:
                        dependents.add(resolved)
        for item in sorted(dependents):
            if item not in extras:
                extras[item] = archive.read(item)
        extras["__content_types__"] = archive.read("[Content_Types].xml")
        extras["__workbook_rels__"] = archive.read("xl/_rels/workbook.xml.rels")
    return extras


def _merge_package_xml(destination, source, match):
    """把 source 清单中与注入部件相关的节点并入 destination(同一根标签)。"""
    dest_root = ET.fromstring(destination)
    src_root = ET.fromstring(source)
    existing = {match(node) for node in dest_root}
    added = False
    for node in src_root:
        key = match(node)
        if key and key not in existing:
            dest_root.append(node)
            existing.add(key)
            added = True
    if not added:
        return destination
    ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/content-types"
                          if dest_root.tag.endswith("Types") else
                          "http://schemas.openxmlformats.org/package/2006/relationships")
    return ET.tostring(dest_root, encoding="UTF-8", xml_declaration=True)


def _inject_extras(out, extras):
    """保存后的输出包回注 cellimages 部件并修补清单/工作簿关系。"""
    if not extras:
        return
    out = Path(out)
    parts = {k: v for k, v in extras.items() if not k.startswith("__")}
    content = extras.get("__content_types__")
    workbook_rels = extras.get("__workbook_rels__")
    part_names = {"/" + name for name in parts}
    media_exts = {name.rsplit(".", 1)[-1].lower() for name in parts if "/media/" in name}
    staged = out.with_name(out.name + ".injecting")
    with ZipFile(out) as zin, ZipFile(staged, "w", ZIP_DEFLATED) as zout:
        existing = set(zin.namelist())
        for name in zin.namelist():
            data = zin.read(name)
            if name == "[Content_Types].xml" and content is not None:
                data = _merge_package_xml(
                    data, content,
                    lambda node: (node.get("PartName"), "override")
                    if node.tag.endswith("Override") and node.get("PartName") in part_names
                    else (node.get("Extension"), "default")
                    if node.tag.endswith("Default") and node.get("Extension", "").lower() in media_exts
                    else None)
            elif name == "xl/_rels/workbook.xml.rels" and workbook_rels is not None:
                data = _merge_package_xml(
                    data, workbook_rels,
                    lambda node: node.get("Target")
                    if node.tag.endswith("Relationship") and "cellimages" in node.get("Target", "")
                    else None)
            zout.writestr(name, data)
        for name, data in parts.items():
            if name not in existing:
                zout.writestr(name, data)
    os.replace(staged, out)


def _prepare(payload, cancel=None):
    path = Path(payload["path"]).expanduser().resolve()
    if path.suffix.lower() != ".xlsx":
        raise ValueError("原生拆分只支持 .xlsx；请先在 Excel 中另存为 .xlsx")
    if not path.is_file():
        raise ValueError("源文件不存在")
    if cancel and cancel():
        raise Cancelled()
    issues = _Issues(payload)
    wb = _load(path, issues)
    try:
        sheet = payload.get("sheet") or wb.active.title
        if sheet not in wb.sheetnames:
            raise ValueError(f"工作表不存在：{sheet}")
        ws = wb[sheet]
        selection = _selection(payload, ws)
        groups, records = _group_records(ws, selection, payload, issues)
        retained = _retained(wb, ws, payload, issues)
        _object_issues(wb, retained, issues, path)
        dispimg = any(isinstance(c.value, str) and "DISPIMG(" in c.value
                      for name in retained for c in wb[name]._cells.values())
        if dispimg:
            if _package_extras(path):
                issues.warn("检测到单元格内嵌图片（WPS DISPIMG）：图片部件将原样保留在每张结果中。")
            else:
                issues.warn("检测到 DISPIMG 图片函数但未找到对应的图片部件；结果中的函数可能无法显示图片。")
        outside = []
        for cell in ws._cells.values():
            record, field = (cell.row, cell.column) if selection["axis"] == "rows" else (cell.column, cell.row)
            if cell.value is not None and selection["data_start"] <= record <= selection["data_end"] and not (
                    selection["field_start"] <= field <= selection["field_end"]):
                outside.append(cell.coordinate)
        if outside:
            # 字段范围外的内容与所选字段同属一条记录, 随整行/列保留是唯一
            # 不丢数据的处理方式——自动执行并提示
            issues.warn("所选字段之外的有效内容（" + "、".join(outside[:20])
                        + ("等" if len(outside) > 20 else "") + "）将随整行/列一起保留在各结果中。")
        issues.warn("仅删去所选数据区内未入组的整行/列；所选工作表的前置区及区外内容保留。")
        if any(c.data_type == "f" for name in retained for c in wb[name]._cells.values()):
            issues.warn("保留公式文本；结果由 Excel 打开后重新计算，预览不把缓存值作为计算结果。")
        if selection["data_end"] < (ws.max_row if selection["axis"] == "rows" else ws.max_column):
            issues.warn("所选数据区之后的单元格会保留，并随行列压缩移动；它们未参与筛选。")
        for group in groups:
            if cancel and cancel():
                raise Cancelled()
            test_book = _load(path, issues)
            try:
                _render_group(test_book, sheet, selection, group, retained, payload, issues)
            finally:
                test_book.close()
        if not groups:
            issues.add("no_records", "没有符合条件的记录，请调整筛选条件。")
        preview = dict(groups=groups, warnings=issues.warnings,
                       issues=list(issues.items.values()),
                       can_export=not any(x.get("blocking", True) for x in issues.items.values()),
                       retained_sheets=retained,
                       omitted_sheets=[x for x in wb.sheetnames if x not in retained],
                       selection=selection,
                       totals=dict(source_records=selection["data_end"] - selection["data_start"] + 1,
                                   selected_records=len(records), groups=len(groups),
                                   output_records=sum(x["count"] for x in groups)))
        return path, sheet, selection, retained, preview
    finally:
        wb.close()


def split_preview(payload):
    """Read the source and return grouping, actionable issues and preservation notes."""
    return _prepare(payload)[-1]


def _safe_name(label):
    name = _INVALID_NAME.sub("_", str(label)).strip(" .")[:80].strip(" .") or "筛选结果"
    if _RESERVED_NAME.match(name):
        name = "_" + name
    return name


def split_export(payload, cancel=None, progress=None):
    """Export new files, never overwrite, and remove this call's outputs on failure.

    progress(done, total, label, record_count) matches the existing merge callback.
    ``Cancelled`` is the existing engine exception. ``SplitBlocked.preview`` carries
    the current preview so stale UI approvals never bypass a fresh safety check.
    """
    path, sheet, selection, retained, preview = _prepare(payload, cancel)
    if not preview["can_export"]:
        raise SplitBlocked(preview)
    extras = _package_extras(path)
    out_dir = Path(payload.get("out_dir") or path.parent / "拆分结果").expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    created, temporary = [], []
    files = []
    try:
        for i, group in enumerate(preview["groups"], 1):
            if cancel and cancel():
                raise Cancelled()
            name = _safe_name(f"{path.stem}_{group['label']}")
            candidate = out_dir / f"{name}.xlsx"
            suffix = 1
            # O_EXCL is also the final collision check: do not rely on exists().
            while True:
                try:
                    handle = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
                    os.close(handle)
                    created.append(candidate)
                    break
                except FileExistsError:
                    suffix += 1
                    candidate = out_dir / f"{name}_{suffix}.xlsx"
            fd, temporary_name = tempfile.mkstemp(prefix=".native-split-", suffix=".xlsx", dir=out_dir)
            os.close(fd)
            temp = Path(temporary_name)
            temporary.append(temp)
            issues = _Issues(payload)
            wb = _load(path, issues)
            try:
                _render_group(wb, sheet, selection, group, retained, payload, issues)
                blocking = [x for x in issues.items.values() if x.get("blocking", True)]
                if blocking:
                    changed = dict(preview, issues=list(issues.items.values()), can_export=False)
                    raise SplitBlocked(changed)
                wb.save(temp)
            finally:
                wb.close()
            _inject_extras(temp, dict(extras))
            try:
                validate_xlsx(temp)
            except ValueError as error:
                raise ValueError(f"生成文件结构检查未通过：{error}")
            if cancel and cancel():
                raise Cancelled()
            os.replace(temp, candidate)  # only replaces the placeholder owned above
            temporary.remove(temp)
            files.append(dict(path=str(candidate), group_id=group["id"],
                              label=group["label"], count=group["count"]))
            if progress:
                progress(i, len(preview["groups"]), group["label"], group["count"])
        if cancel and cancel():
            raise Cancelled()
        return dict(preview, output=str(out_dir), files=files)
    except BaseException:
        for item in temporary + created:
            item.unlink(missing_ok=True)
        raise
