"""Explicit, identity-based reconciliation of XLSX feedback into a current workbook.

All selection indices are absolute, one-based and inclusive.  For ``rows``,
header/data ranges refer to rows and fields to columns; ``columns`` transposes
these meanings.  Preview and export share the same read-only decision phase.
"""
from collections import defaultdict
from copy import copy, deepcopy
from dataclasses import dataclass
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import tempfile
import warnings
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.worksheet.print_settings import PrintArea

from .merge import Cancelled
from .compatibility import inspect_extensions


def _json(value):
    if isinstance(value, CellRichText):
        return str(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(k): _json(v) for k, v in value.items()}
    if isinstance(value, (tuple, list)):
        return [_json(v) for v in value]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _digest(value):
    text = json.dumps(_json(value), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:20]


def _key(values):
    # Do not silently identify numeric 1 with text "1", or True with 1.
    return tuple((type(v).__name__, _json(v)) for v in values)


def _blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _cancelled(cancel):
    if cancel and cancel():
        raise Cancelled("对照回填已取消")


def _check_package(path):
    if path.suffix.lower() != ".xlsx":
        raise ValueError("对照回填只支持 .xlsx；请先在 Excel/WPS 中另存为 .xlsx")
    if not path.is_file():
        raise ValueError(f"文件不存在：{path}")
    # openpyxl cannot round-trip these objects reliably. Reject before loading,
    # including disguised macro files and signed packages, rather than dropping them.
    prefixes = ("xl/drawings/", "xl/charts/", "xl/media/", "xl/embeddings/",
                "xl/activex/", "xl/ctrlprops/", "xl/pivottables/", "xl/pivotcache/",
                "xl/slicers/", "xl/slicercaches/", "xl/threadedcomments/",
                "xl/persons/", "xl/richdata/", "customxml/", "customui/",
                "_xmlsignatures/", "xl/externallinks/",
                "xl/macrosheets/", "xl/dialogsheets/", "xl/model/", "xl/chartsheets/",
                "xl/querytables/")
    notices = []
    with ZipFile(path) as archive:
        for name in archive.namelist():
            lower = name.lower()
            if lower.startswith("xl/printersettings/"):
                notices.append("打印机专属二进制设置不保留；常规纸张、方向和页边距继续保留。")
            if (lower.startswith(prefixes) or "vbaproject" in lower
                    or lower in ("xl/connections.xml", "xl/metadata.xml", "xl/xmlmaps.xml")):
                raise ValueError(f"含当前回填器无法完整保留的对象：{name}；已阻止导出")
            if lower.startswith("xl/") and lower.endswith(".xml"):
                raw = archive.read(name)
                notices.extend(inspect_extensions(name, raw))
    return list(dict.fromkeys(notices))


@dataclass
class Selection:
    axis: str
    header_start: int
    header_end: int
    data_start: int
    data_end: int
    field_start: int
    field_end: int

    @classmethod
    def parse(cls, value, ws):
        if not isinstance(value, dict):
            raise ValueError("请明确选择表头、数据和字段范围")
        if value.get("axis") not in ("rows", "columns"):
            raise ValueError("selection.axis 必须为 rows 或 columns")
        keys = ("header_start", "header_end", "data_start", "data_end", "field_start", "field_end")
        if any(type(value.get(k)) is not int or value[k] < 1 for k in keys):
            raise ValueError("选择范围必须使用从 1 开始的整数索引")
        result = cls(value["axis"], *(value[k] for k in keys))
        if not (result.header_start <= result.header_end < result.data_start <= result.data_end
                and result.field_start <= result.field_end):
            raise ValueError("范围无效：表头应在数据之前，各范围起点不得大于终点")
        length, width = ((ws.max_row, ws.max_column) if result.axis == "rows"
                         else (ws.max_column, ws.max_row))
        if result.data_end > length or result.field_end > width:
            raise ValueError("所选数据或字段范围超过工作表实际范围")
        return result

    @property
    def fields(self):
        return range(self.field_start, self.field_end + 1)

    def coord(self, record, field):
        return (record, field) if self.axis == "rows" else (field, record)

    def cell(self, ws, record, field):
        return ws.cell(*self.coord(record, field))


class Source:
    def __init__(self, item, fallback, books, notices):
        if not isinstance(item, dict) or not item.get("path"):
            raise ValueError("缺少工作簿路径")
        self.path = Path(item["path"]).expanduser().resolve()
        compatibility = _check_package(self.path)
        for notice in compatibility:
            message = f"{self.path.name}：{notice}"
            if message not in notices:
                notices.append(message)
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always", UserWarning)
            self.book = load_workbook(self.path, data_only=False, rich_text=True)
        books.append(self.book)
        for warning in caught:
            message = str(warning.message)
            if (issubclass(warning.category, UserWarning)
                    and message.endswith(" extension is not supported and will be removed")):
                notice = f"{self.path.name} 读取兼容提示：{message}"
                if notice not in notices:
                    notices.append(notice)
            else:
                raise ValueError(f"{self.path.name} 读取时发现尚未确认的兼容问题：{message}")
        name = item.get("sheet") or self.book.active.title
        if name not in self.book.sheetnames:
            raise ValueError(f"{self.path.name} 中不存在工作表 {name}")
        self.ws = self.book[name]
        selection = item.get("selection")
        if selection is None and isinstance(fallback, dict):
            selection = dict(fallback)
            # An omitted feedback/baseline selection retains the layout but uses
            # that file's own extent: returned files commonly contain fewer tasks.
            selection["data_end"] = self.ws.max_row if selection.get("axis") == "rows" else self.ws.max_column
            selection["field_end"] = self.ws.max_column if selection.get("axis") == "rows" else self.ws.max_row
        self.selection = Selection.parse(selection, self.ws)
        self.field_map = item.get("field_map", {})
        if not isinstance(self.field_map, dict):
            raise ValueError("field_map 必须为目标字段索引到反馈字段索引的对象")
        self.headers = {}
        for field in self.selection.fields:
            parts = []
            for h in range(self.selection.header_start, self.selection.header_end + 1):
                cell = self.selection.cell(self.ws, h, field)
                for merged in self.ws.merged_cells.ranges:
                    if cell.coordinate in merged:
                        cell = self.ws.cell(merged.min_row, merged.min_col)
                        break
                if not _blank(cell.value):
                    parts.append(str(cell.value).strip())
            self.headers[field] = tuple(parts)
        content_hash = hashlib.sha256(self.path.read_bytes()).hexdigest()
        self.token = _digest([content_hash, name, vars(self.selection), self.field_map])
        self.label = f"{self.path.name} / {name}"

    def records(self):
        s = self.selection
        return [(r, {f: self.logical_value(r, f) for f in s.fields})
                for r in range(s.data_start, s.data_end + 1)
                if any(not _blank(s.cell(self.ws, r, f).value) for f in s.fields)]

    def logical_value(self, record, field):
        cell = self.selection.cell(self.ws, record, field)
        if isinstance(cell, MergedCell):
            for merged in self.ws.merged_cells.ranges:
                if cell.coordinate in merged:
                    return self.ws.cell(merged.min_row, merged.min_col).value
        return cell.value

    def descriptor(self):
        return {"path": str(self.path), "sheet": self.ws.title, "id": "source:" + self.token}


class Plan:
    def __init__(self, payload):
        self.payload = payload
        self.books = []
        self.issues, self.changes, self.missing = [], [], []
        self.warnings = []
        self.actions, self.additions, self.red = [], [], []
        self.restore_colors = []
        self.current = None
        self.resolutions = payload.get("resolutions") or {}
        if not isinstance(self.resolutions, dict):
            self.resolutions = {}
        self.feedback_count = self.duplicate_imports = 0
        self.record_count = 0

    def issue(self, kind, identity, message, choices=(), **extra):
        iid = kind + ":" + _digest(identity)
        value = self.resolutions.get(iid)
        options = [{"value": v, "label": label} for v, label in choices]
        resolved = bool(options) and value in {c["value"] for c in options}
        item = {"id": iid, "kind": kind, "message": message, "choices": options,
                "resolved": resolved, "resolution": value if resolved else None, **_json(extra)}
        self.issues.append(item)
        return value if resolved else None

    def close(self):
        for book in self.books:
            book.close()

    def result(self):
        unresolved = sum(not i["resolved"] for i in self.issues)
        fields = [] if not self.current else [
            {"index": f, "title": " / ".join(self.current.headers[f]) or f"字段 {f}",
             "path": list(self.current.headers[f])}
            for f in self.current.selection.fields]
        return _json({
            "changes": self.changes, "issues": self.issues, "missing": self.missing,
            "warnings": self.warnings,
            "can_export": not unresolved, "fields": fields,
            "totals": {"records": self.record_count, "feedback_records": self.feedback_count,
                       "changes": len(self.changes), "missing": len(self.missing),
                       "additions": len(self.additions), "unresolved": unresolved,
                       "duplicate_imports": self.duplicate_imports},
        })


def _fields(plan, name, required=False):
    values = plan.payload.get(name, [])
    valid = (isinstance(values, list) and all(type(v) is int for v in values)
             and len(values) == len(set(values))
             and all(v in plan.current.selection.fields for v in values))
    if not valid or (required and not values):
        plan.issue("field_selection", name, f"请明确选择有效且不重复的 {name} 字段")
        return []
    return values


def _mapping(plan, source, required):
    target = plan.current
    result = {}
    target_counts = defaultdict(int)
    source_by_title = defaultdict(list)
    for title in target.headers.values():
        target_counts[title] += 1
    for field, title in source.headers.items():
        source_by_title[title].append(field)
    for field in required:
        specified = str(field) in source.field_map or field in source.field_map
        explicit = source.field_map.get(str(field), source.field_map.get(field))
        if specified:
            try:
                mapped = int(explicit) if explicit is not None and not isinstance(explicit, bool) else None
            except (TypeError, ValueError):
                mapped = None
            if mapped in source.selection.fields:
                result[field] = mapped
                continue
        else:
            title = target.headers[field]
            matches = source_by_title[title]
            if title and target_counts[title] == 1 and len(matches) == 1:
                result[field] = matches[0]
                continue
        plan.issue("field_mapping", [source.token, field],
                   f"{source.label} 无法唯一对应字段 {field}（{' / '.join(target.headers[field])}）；"
                   "请通过 field_map 指定目标字段与反馈字段的索引",
                   source=source.descriptor(), target_field=field,
                   available_fields=[{"index": f, "path": list(t)} for f, t in source.headers.items()])
    if len(set(result.values())) != len(result):
        plan.issue("field_mapping", [source.token, "duplicate"],
                   f"{source.label} 将多个目标字段映射到了同一反馈字段")
    return result


def _indexed(plan, source, key_fields, mapping, role):
    groups = defaultdict(list)
    if any(f not in mapping for f in key_fields):
        return {}
    for position, values in source.records():
        key_values = [values[mapping[f]] for f in key_fields]
        if any(_blank(v) for v in key_values):
            plan.issue("empty_key", [role, source.token, position],
                       f"{source.label} 的第 {position} 条记录主键为空", source=source.descriptor(),
                       position=position)
            continue
        if any(source.selection.cell(source.ws, position, mapping[f]).data_type == "f" for f in key_fields):
            plan.issue("formula_key", [role, source.token, position],
                       f"{source.label} 的主键含公式，请选择稳定的记录标识字段")
            continue
        groups[_key(key_values)].append((position, values, key_values))
    result = {}
    for key, records in groups.items():
        if len(records) != 1:
            plan.issue("duplicate_key", [role, source.token, key],
                       f"{source.label} 存在重复主键，不能自动选择第一条或最后一条",
                       key=records[0][2], positions=[v[0] for v in records],
                       source=source.descriptor())
        else:
            result[key] = records[0]
    return result


def _translated(value, source_cell, target_cell):
    if isinstance(value, str) and value.startswith("="):
        origin = source_cell if isinstance(source_cell, str) else source_cell.coordinate
        destination = target_cell if isinstance(target_cell, str) else target_cell.coordinate
        try:
            return Translator(value, origin=origin).translate_formula(destination)
        except (TranslatorError, ValueError) as exc:
            raise ValueError(f"公式无法安全平移：{origin} → {destination}") from exc
    # Rich text is retained by openpyxl; comparison uses its textual content.
    return value


def _coordinate(selection, record, field):
    row, col = selection.coord(record, field)
    return f"{get_column_letter(col)}{row}"


def _equivalent(a, b):
    if isinstance(a, CellRichText):
        a = str(a)
    if isinstance(b, CellRichText):
        b = str(b)
    return type(a) is type(b) and a == b or (
        isinstance(a, (int, float)) and not isinstance(a, bool)
        and isinstance(b, (int, float)) and not isinstance(b, bool) and a == b)


def _check_formula_mapping(value, source, target, mapping):
    if not (isinstance(value, str) and value.startswith("=")):
        return
    offsets = {field - mapped for field, mapped in mapping.items()}
    if source.selection.axis != target.selection.axis or len(offsets) != 1:
        raise ValueError(f"{source.label} 含公式且字段发生非等距换序/转置；"
                         "无法证明引用对象保持一致，请在来源中对齐字段后再回填")


def _append_reason(source):
    ws, s = source.ws, source.selection
    end = ws.max_row if s.axis == "rows" else ws.max_column
    if s.data_end != end:
        return "数据区后还有页脚或其他区域，无法保证安全追加"
    if ws.tables:
        return "工作表含 Excel 表对象，暂不能安全扩展表范围"
    if source.book.defined_names:
        return "工作簿含自定义名称，暂不能保证追加后引用范围完整"
    for merged in ws.merged_cells.ranges:
        if (s.axis == "rows" and merged.min_row <= s.data_end <= merged.max_row
                or s.axis == "columns" and merged.min_col <= s.data_end <= merged.max_col):
            return "追加模板记录含合并单元格，无法确定复制方式"
    for sheet in source.book:
        for row in sheet:
            for cell in row:
                if cell.data_type == "f":
                    r, f = ((cell.row, cell.column) if s.axis == "rows" else (cell.column, cell.row))
                    if sheet != ws or not (s.data_start <= r <= s.data_end and f in s.fields):
                        return "数据区外存在公式，暂不能保证新增记录被其正确引用"
    return None


def _analyse(payload, cancel=None, progress=None):
    plan = Plan(payload)
    try:
        _cancelled(cancel)
        current = plan.current = Source(payload, None, plan.books, plan.warnings)
        if current.ws.protection.sheet:
            plan.issue("protected_sheet", current.token, "当前工作表受保护；请先在 Excel/WPS 中确认并解除保护")
        keys = _fields(plan, "key_fields", True)
        updates = _fields(plan, "update_fields", True)
        compares = _fields(plan, "compare_fields")
        if set(keys) & set(updates):
            plan.issue("field_selection", "keys_updated", "主键字段不能同时作为回填字段")
        if not keys or not updates:
            return plan
        records = _indexed(plan, current, keys, {f: f for f in current.selection.fields}, "current")
        plan.record_count = len(records)
        baseline = current
        base_map = {f: f for f in current.selection.fields}
        base_records = records
        if payload.get("baseline"):
            baseline = Source(payload["baseline"], payload.get("selection"), plan.books, plan.warnings)
            base_map = _mapping(plan, baseline, sorted(set(keys + updates + compares)))
            base_records = _indexed(plan, baseline, keys, base_map, "baseline")
        items = payload.get("feedback", [])
        if not isinstance(items, list) or not items:
            plan.issue("feedback_required", "feedback", "请至少选择一个反馈工作簿")
            return plan
        proposals = defaultdict(list)
        seen_imports = set()
        for index, item in enumerate(items):
            _cancelled(cancel)
            feedback = Source(item, payload.get("selection"), plan.books, plan.warnings)
            if feedback.token in seen_imports:
                plan.duplicate_imports += 1
                continue
            seen_imports.add(feedback.token)
            mapping = _mapping(plan, feedback, sorted(set(keys + updates)))
            found = _indexed(plan, feedback, keys, mapping, "feedback")
            if len(mapping) != len(set(keys + updates)):
                continue
            for key, (position, values, key_values) in found.items():
                plan.feedback_count += 1
                record_map = mapping
                if key not in records:
                    # Confirming a new task appends the whole task, including its
                    # identity/context fields, not just the selected update fields.
                    record_map = _mapping(plan, feedback, list(current.selection.fields))
                    if len(record_map) != len(current.selection.fields):
                        continue
                proposals[key].append((feedback, position, values, key_values, record_map))
            if progress:
                progress(index + 1, len(items), feedback.label, plan.feedback_count)
        plan.missing = [{"key": values, "position": position, "status": "not_received",
                         "action": "keep_current", "message": "本次未反馈，保留当前内容"}
                        for key, (position, _, values) in records.items() if key not in proposals]
        final = {key: (pos, dict(values), key_values, current.descriptor())
                 for key, (pos, values, key_values) in records.items()}
        appended = current.selection.data_end
        for key, candidates in proposals.items():
            _cancelled(cancel)
            exists = key in records
            position = records[key][0] if exists else appended + 1
            converted = []
            for source, row, values, key_values, mapping in candidates:
                new_values = {}
                for field in (keys + updates if exists else list(current.selection.fields)):
                    target_cell = _coordinate(current.selection, position, field)
                    source_cell = source.selection.cell(source.ws, row, mapping[field])
                    if source_cell.data_type == "f" and not isinstance(source_cell.value, str):
                        raise ValueError("所选回填字段含数组/数据表公式，暂不能证明独立回填和引用平移安全")
                    _check_formula_mapping(values[mapping[field]], source, current, mapping)
                    new_values[field] = _translated(values[mapping[field]], source_cell, target_cell)
                converted.append((source, new_values, key_values))
            unique = {}
            for source, values, key_values in converted:
                unique.setdefault(_digest(values), (source, values, key_values))
            chosen = next(iter(unique.values()))
            if len(unique) > 1:
                choices = [("keep", "保留当前内容" if exists else "跳过此新增记录")]
                choices.extend(("source:" + src.token, f"采用 {src.label}") for src, _, _ in unique.values())
                answer = plan.issue("conflict", key, "同一记录收到不同反馈，请选择采用的反馈或保留当前内容",
                                    choices, key=chosen[2],
                                    candidates=[{"source": src.descriptor(), "values": vals}
                                                for src, vals, _ in unique.values()])
                if answer is None or answer == "keep":
                    continue
                chosen = next(value for value in unique.values() if "source:" + value[0].token == answer)
            source, values, key_values = chosen
            if not exists:
                reason = _append_reason(current)
                choices = [("append", "追加为新记录"), ("skip", "跳过此新增记录")] if not reason else [
                    ("skip", "跳过此新增记录")]
                answer = plan.issue("addition", key, "发现当前表中没有的记录" + (f"；{reason}" if reason else ""),
                                    choices, key=key_values, source=source.descriptor(), append_blocked=reason)
                if answer != "append":
                    continue
                appended += 1
                position = appended
                plan.additions.append(position)
                final[key] = (position, {}, key_values, source.descriptor())
            old_position, existing, _, _ = final[key]
            existing.update(values)
            final[key] = (old_position, existing, key_values, source.descriptor())
            for field in (list(current.selection.fields) if not exists else updates):
                cell = current.selection.cell(current.ws, position, field) if exists else None
                if exists and _equivalent(current.logical_value(position, field), values[field]):
                    continue
                if isinstance(cell, MergedCell):
                    plan.issue("merged_target", [key, field],
                               f"目标 {cell.coordinate} 是合并单元格的一部分，无法独立回填", key=key_values)
                else:
                    plan.actions.append((position, field, values[field]))
        # Compare final content against the explicit baseline; never infer change
        # from red font, or from the physical row number inside a moved formula.
        for key, (position, values, key_values, source) in final.items():
            baseline_record = base_records.get(key)
            fields = sorted(set(updates + compares))
            if key not in records:
                plan.changes.append({"id": "new:" + _digest(key), "key": key_values, "field": None,
                                     "before": None, "after": values, "kind": "new",
                                     "source": source, "position": position, "marked": bool(compares)})
            for field in fields:
                after = values.get(field)
                before = None
                if baseline_record and field in base_map:
                    br, bv, _ = baseline_record
                    _check_formula_mapping(bv[base_map[field]], baseline, current, base_map)
                    before = _translated(bv[base_map[field]],
                                         baseline.selection.cell(baseline.ws, br, base_map[field]),
                                         _coordinate(current.selection, position, field))
                if _equivalent(before, after):
                    if (payload.get("baseline") and baseline_record and field in compares
                            and field in base_map and key in records):
                        # Red is a presentation of this baseline comparison, not a
                        # permanent accumulated state. A reverted value adopts the
                        # baseline's colour while keeping the current font's other
                        # attributes. An originally red baseline remains red.
                        current_cell = current.selection.cell(current.ws, position, field)
                        color = current_cell.font.color
                        if color and color.type == "rgb" and color.rgb[-6:].upper() == "FF0000":
                            baseline_cell = baseline.selection.cell(baseline.ws, baseline_record[0],
                                                                    base_map[field])
                            plan.restore_colors.append((position, field, copy(baseline_cell.font.color)))
                    continue
                marked = field in compares
                kind = "cancel" if isinstance(after, str) and after.strip() in ("取消", "已取消", "作废") else "update"
                plan.changes.append({"id": "change:" + _digest([key, field]), "key": key_values,
                                     "field": field, "field_title": " / ".join(current.headers[field]),
                                     "before": before, "after": after, "kind": kind,
                                     "source": source, "position": position, "marked": marked,
                                     "baseline_missing": baseline_record is None})
                if marked:
                    plan.red.append((position, field))
        if payload.get("out"):
            out = Path(payload["out"]).expanduser().resolve()
            sources = {current.path, baseline.path}
            sources.update(Path(i["path"]).expanduser().resolve() for i in items)
            if out.suffix.lower() != ".xlsx":
                plan.issue("output", "extension", "输出文件必须使用 .xlsx 扩展名")
            if out in sources:
                plan.issue("output", "source", "输出路径不能覆盖当前表、基线或任何反馈文件")
            if out.exists() or out.with_suffix(".changes.json").exists():
                plan.issue("output", "exists", "输出文件或变更清单已经存在，请选择新的文件名")
        return plan
    except Cancelled:
        plan.close()
        raise
    except (ValueError, OSError, BadZipFile, InvalidFileException, KeyError, TypeError) as exc:
        plan.issue("invalid_input", str(exc), str(exc))
        return plan


def reconcile_preview(payload):
    """Return JSON-safe changes/issues without changing any workbook."""
    plan = _analyse(payload)
    try:
        return plan.result()
    finally:
        plan.close()


def _extend_range(rng, axis, template, destination):
    result = CellRange(str(rng))
    if axis == "rows" and result.max_row == template:
        result.max_row = destination
    elif axis == "columns" and result.max_col == template:
        result.max_col = destination
    return result


def _append_template(source, positions):
    if not positions:
        return
    ws, s = source.ws, source.selection
    template = s.data_end
    for dest in positions:
        if s.axis == "rows":
            dimension = copy(ws.row_dimensions[template])
            dimension.index = dest
            ws.row_dimensions[dest] = dimension
            pairs = [(ws.cell(template, c), ws.cell(dest, c)) for c in range(1, ws.max_column + 1)]
        else:
            dimension = copy(ws.column_dimensions[get_column_letter(template)])
            dimension.index = get_column_letter(dest)
            ws.column_dimensions[dimension.index] = dimension
            pairs = [(ws.cell(r, template), ws.cell(r, dest)) for r in range(1, ws.max_row + 1)]
        for original, cell in pairs:
            cell._style = copy(original._style)
            if original.data_type == "f":
                cell.value = _translated(original.value, original, cell)
    last = positions[-1]
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange([_extend_range(r, s.axis, template, last) for r in dv.sqref.ranges])
    conditional = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        ranges = MultiCellRange([_extend_range(r, s.axis, template, last) for r in cf.sqref.ranges])
        for rule in ws.conditional_formatting[cf]:
            conditional.add(str(ranges), deepcopy(rule))
    ws.conditional_formatting = conditional
    if ws.print_area:
        ws.print_area = [_extend_range(r, s.axis, template, last).coord
                         for r in PrintArea.from_string(ws.print_area).ranges]


def reconcile_export(payload, cancel=None, progress=None):
    """Export to a new XLSX and adjacent .changes.json; unresolved issues block."""
    plan = _analyse(payload, cancel, progress)
    temporary = []
    created = []
    try:
        if not payload.get("out"):
            plan.issue("output", "missing", "请选择输出文件路径")
        result = plan.result()
        result["exported"] = False
        if not result["can_export"]:
            return result
        _cancelled(cancel)
        out = Path(payload["out"]).expanduser().resolve()
        log = out.with_suffix(".changes.json")
        if not out.parent.is_dir():
            raise ValueError("输出目录不存在")
        _append_template(plan.current, plan.additions)
        for position, field, value in plan.actions:
            _cancelled(cancel)
            plan.current.selection.cell(plan.current.ws, position, field).value = value
        for position, field, color in plan.restore_colors:
            cell = plan.current.selection.cell(plan.current.ws, position, field)
            font = copy(cell.font)
            font.color = color
            cell.font = font
        for position, field in plan.red:
            cell = plan.current.selection.cell(plan.current.ws, position, field)
            font = copy(cell.font)
            font.color = "FFFF0000"
            cell.font = font
        for suffix in (".xlsx", ".json"):
            fd, name = tempfile.mkstemp(prefix=".reconcile-", suffix=suffix, dir=out.parent)
            os.close(fd)
            temporary.append(Path(name))
        plan.current.book.save(temporary[0])
        result.update({"out": str(out), "changes_path": str(log), "exported": True})
        temporary[1].write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        _cancelled(cancel)
        # Hard links create destinations exclusively, so a concurrent file creation
        # cannot be overwritten between preview and final publication.
        for temp, destination in zip(temporary, (out, log)):
            os.link(temp, destination)
            created.append(destination)
        return result
    except (ValueError, OSError) as exc:
        for path in created:
            path.unlink(missing_ok=True)
        plan.issue("output", str(exc), f"导出失败：{exc}")
        result = plan.result()
        result["exported"] = False
        return result
    finally:
        for path in temporary:
            path.unlink(missing_ok=True)
        plan.close()
