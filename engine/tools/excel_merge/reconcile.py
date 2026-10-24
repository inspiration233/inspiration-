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
from openpyxl.styles import PatternFill

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
    hard_prefixes = ("xl/activex/", "xl/ctrlprops/", "xl/pivottables/", "xl/pivotcache/",
                     "xl/slicers/", "xl/slicercaches/", "xl/threadedcomments/",
                     "xl/persons/", "xl/richdata/", "customxml/", "customui/",
                     "_xmlsignatures/", "xl/externallinks/",
                     "xl/macrosheets/", "xl/dialogsheets/", "xl/model/", "xl/chartsheets/",
                     "xl/querytables/", "xl/embeddings/")
    soft_prefixes = ("xl/drawings/", "xl/charts/", "xl/media/")
    notices = []
    with ZipFile(path) as archive:
        for name in archive.namelist():
            lower = name.lower()
            if lower.startswith("xl/printersettings/"):
                notices.append("打印机专属二进制设置不保留；常规纸张、方向和页边距继续保留。")
            if lower.startswith(soft_prefixes):
                # 图形/图表/媒体: 回填不改布局,这些对象在原位不动;可能不完美保留但不会丢数据
                notices.append(f"{name} 将按原样随结果保留；外观如有变化请核对。")
                continue
            if (lower.startswith(hard_prefixes) or "vbaproject" in lower
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
        self._cached_book = None
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

    def close(self):
        if self._cached_book:
            try:
                self._cached_book.close()
            except Exception:
                pass
        try:
            self.book.close()
        except Exception:
            pass

    def cached_value(self, record, field):
        """公式格的计算值(真实 Excel/WPS 保存的文件才有缓存); 供公式主键使用。"""
        if self._cached_book is None:
            try:
                self._cached_book = load_workbook(self.path, data_only=True)
            except Exception:
                self._cached_book = False
        if not self._cached_book:
            return None
        try:
            return self.selection.cell(self._cached_book[self.ws.title], record, field).value
        except Exception:
            return None

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
        self.additions_skipped = 0

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
                       "additions": len(self.additions), "additions_skipped": self.additions_skipped,
                       "unresolved": unresolved, "duplicate_imports": self.duplicate_imports,
                       "updates": sum(1 for c in self.changes if c.get("kind") in ("update", "cancel")),
                       "fills": sum(1 for c in self.changes if c.get("kind") == "fill"),
                       "clears": sum(1 for c in self.changes if c.get("kind") == "clear")},
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
    by_position = []
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
            if title and field in matches:
                # 多列同名(如多列都叫 XXX)时按列位置对应: 只在同名候选中选择
                # 位置一致的那一列(同模板文件列布局一致), 歧义无第二种合理解释
                result[field] = field
                by_position.append(field)
                continue
        plan.issue("field_mapping", [source.token, field],
                   f"{source.label} 无法唯一对应字段 {field}（{' / '.join(target.headers[field])}）；"
                   "请在界面的字段对应表中为它选择来源列",
                   source=source.descriptor(), target_field=field,
                   available_fields=[{"index": f, "path": list(t)} for f, t in source.headers.items()])
    if by_position:
        plan.warnings.append(
            f"{source.label} 有 {len(by_position)} 个同名字段（如 {' / '.join(target.headers[by_position[0]])}）"
            "无法按标题区分，已按相同列位置对应")
    if len(set(result.values())) != len(result):
        plan.issue("field_mapping", [source.token, "duplicate"],
                   f"{source.label} 将多个目标字段映射到了同一反馈字段")
    return result


def _indexed(plan, source, key_fields, mapping, role, known_keys=None):
    groups = defaultdict(list)
    merged_groups = []
    if any(f not in mapping for f in key_fields):
        return {}
    empty_count = 0
    formula_key_count = 0
    for position, values in source.records():
        key_values = [values[mapping[f]] for f in key_fields]
        if any(_blank(v) for v in key_values):
            # 主键为空无法对应任何记录: 跳过该条(等价保留现状), 汇总提示
            empty_count += 1
            continue
        if any(source.selection.cell(source.ws, position, mapping[f]).data_type == "f" for f in key_fields):
            cached = [source.cached_value(position, mapping[f]) for f in key_fields]
            if all(v not in (None, "") for v in cached):
                # 公式主键(如 =A2&B2 拼接编号)按计算值匹配
                formula_key_count += 1
                key_values = cached
            else:
                plan.issue("formula_key", [role, source.token, position],
                           f"{source.label} 的主键含公式且文件中没有可用的计算值，请选择稳定的记录标识字段")
                continue
        groups[_key(key_values)].append((position, values, key_values))
    result = {}
    for key, records in groups.items():
        if len(records) == 1:
            result[key] = records[0]
            continue
        mapped_cols = set(mapping.values())
        relevant = lambda rec: {c: rec[1].get(c) for c in mapped_cols}
        if all(relevant(rec) == relevant(records[0]) for rec in records[1:]):
            # 主键相同且本次操作关心的字段(主键+回填/比较字段)内容一致:
            # 差异只在无关列, 取哪条都一样 → 自动合并
            plan.warnings.append(
                f"{source.label} 有 {len(records)} 条主键相同的记录，本次回填关注的字段内容一致，已合并处理。")
            result[key] = records[0]
        elif role == "feedback":
            # 同一反馈文件内主键重复且关注字段不同: 视同多份反馈冲突
            if known_keys is not None and key in known_keys:
                plan.warnings.append(
                    f"{source.label} 对记录 {' / '.join(str(v) for v in records[0][2])} 给出了 "
                    f"{len(records)} 条不同反馈（第 {'、'.join(str(v[0]) for v in records)} 条），已保留当前内容。")
            else:
                # 当前表按此主键没有这条记录: 说"保留当前"会误导, 实际是整组忽略
                plan.warnings.append(
                    f"{source.label} 有 {len(records)} 条主键相同的反馈（{' / '.join(str(v) for v in records[0][2])}，"
                    f"第 {'、'.join(str(v[0]) for v in records)} 条）内容互相矛盾，"
                    "无法确定采用哪条，已全部忽略。")
        else:
            merged_key = any(
                isinstance(source.selection.cell(source.ws, position, mapping[f]), MergedCell)
                for position, _, _ in records for f in key_fields)
            if merged_key:
                # 纵向合并组: 多行共用合并锚点值, 不是表做错了, 是主键选了组名列;
                # 同一原因聚合为一条提示
                merged_groups.append((records[0][2], len(records)))
            else:
                plan.issue("duplicate_key", [role, source.token, key],
                           f"{source.label} 存在重复主键且内容不同，无法确定哪条是这条记录；"
                           "自动跳过会导致反馈被误认为新记录追加，请先修正原表的主键",
                           key=records[0][2], positions=[v[0] for v in records],
                           source=source.descriptor())
    if merged_groups:
        values = "、".join(" / ".join(str(v) for v in key_values) for key_values, _ in merged_groups[:4])
        if len(merged_groups) > 4:
            values += "…"
        merged_rows = sum(rows for _, rows in merged_groups)
        # 建议改用的主键列: 数据全非空且每行唯一, 优先标题含序号/编号
        columns = {f: [] for f in source.selection.fields}
        for recs in groups.values():
            for _, values_row, _ in recs:
                for f in columns:
                    columns[f].append(values_row.get(f))
        unique_fields = [f for f, vals in columns.items()
                         if vals and not any(_blank(v) for v in vals)
                         and len({_key([v]) for v in vals}) == len(vals)]
        preferred = next((f for f in unique_fields
                          if any(t and ("序号" in t or "编号" in t)
                                 for t in source.headers.get(f, ()))), None)
        suggested = preferred or (unique_fields[0] if unique_fields else None)
        issue_extra = {"keys": [k for k, _ in merged_groups],
                       "source": source.descriptor()}
        if suggested is not None:
            issue_extra["suggested_key"] = {"field": suggested,
                                            "title": " / ".join(source.headers.get(suggested, ()))}
        plan.issue("duplicate_key", [role, source.token, "merged"],
                   f"{source.label} 的主键列含合并单元格：{len(merged_groups)} 组"
                   f"（{values}）每组多行共用一个合并值、各行内容又不同；"
                   f"受影响的 {merged_rows} 行数据当前都无法按此主键对应，"
                   "请改选每行唯一的列（如序号/编号）作主键",
                   **issue_extra)
    if empty_count:
        plan.warnings.append(
            f"{source.label} 有 {empty_count} 条主键为空的记录已跳过（无法与任何记录对应，保留现状）。")
    if formula_key_count:
        plan.warnings.append(
            f"{source.label} 有 {formula_key_count} 条主键为公式，已按其计算值匹配。")
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


def _example_marked(value):
    return isinstance(value, str) and value.strip().startswith("例")


def _group_records(plan, source, key_field, group_field, mapping, role):
    """分组身份索引。底稿序号通常全局连续、拆分产物序号按组内从 1 开始,
    数值不一致是分发-回填流程的常态, 因此不按序号数值匹配:
    "例"开头序号按值全局对应(模板示例); 其余记录按(组值, 组内第几条真实记录)对应。"""
    group_counter = defaultdict(int)
    result = {}
    duplicates = defaultdict(list)
    empty_count = 0
    for position, values in source.records():
        key_value = values[mapping[key_field]]
        group_value = values[mapping[group_field]]
        if _blank(key_value) or _blank(group_value):
            empty_count += 1
            continue
        if _example_marked(key_value):
            identity = ("ex",) + _key([key_value])
            display = [key_value]
        else:
            group_key = _key([group_value])
            group_counter[group_key] += 1
            identity = ("grp", group_key, group_counter[group_key])
            display = [group_value, group_counter[group_key]]
        if identity in result:
            duplicates[identity].append((position, values, display))
        else:
            result[identity] = (position, values, display)
    for identity, extra in duplicates.items():
        all_records = [result[identity]] + extra
        mapped_cols = set(mapping.values())
        relevant = lambda rec: {c: rec[1].get(c) for c in mapped_cols}
        if all(relevant(rec) == relevant(all_records[0]) for rec in all_records[1:]):
            plan.warnings.append(
                f"{source.label} 有 {len(all_records)} 条身份相同的记录"
                f"（{' / '.join(str(v) for v in all_records[0][2])}），本次回填关注的字段内容一致，已合并处理。")
        else:
            plan.issue("duplicate_key", [role, source.token, list(identity)],
                       f"{source.label} 存在身份重复且内容不同的记录"
                       f"（{' / '.join(str(v) for v in all_records[0][2])}，"
                       f"第 {'、'.join(str(r[0]) for r in all_records)} 条）；请先修正或更换分组/序号列",
                       key=all_records[0][2], positions=[r[0] for r in all_records],
                       source=source.descriptor())
    if empty_count:
        plan.warnings.append(
            f"{source.label} 有 {empty_count} 条记录的序号或分组为空，已跳过（无法对应，保留现状）。")
    return result


def _key_column_merged(source, key_field):
    """主键列在数据区是否含沿记录方向的合并(合并块=一条记录的子行结构)。"""
    s = source.selection
    for merged in source.ws.merged_cells.ranges:
        if s.axis == "rows":
            if (merged.min_row < merged.max_row and merged.min_col <= key_field <= merged.max_col
                    and merged.max_row >= s.data_start and merged.min_row <= s.data_end):
                return True
        elif (merged.min_col < merged.max_col and merged.min_row <= key_field <= merged.max_row
                and merged.max_col >= s.data_start and merged.min_col <= s.data_end):
            return True
    return False


def _serial_ordinal(value):
    """拆分产物的组内序号是 1..N 的数字; 文本例N等返回 None。"""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int):
        return value
    text = str(value).strip()
    return int(text) if text.isdigit() else None


def _align_group_serials(found, base_group_ords, key_field, mapping):
    """按主键对位: 反馈某组的序号值若互不相同且都落在底稿组内序号范围内
    (单位删行而未重排), 用"序号值=组内第k条"重对位——精确定位被删记录,
    并避免其后的记录整体滑位。序号不可解读的组保持原按位对位。
    返回 (调整后的found, 已按主键对位的组集合)。"""
    groups = defaultdict(list)
    passthrough = []
    for ident, entry in found.items():
        if isinstance(ident, tuple) and ident[0] == "grp":
            groups[ident[1]].append((ident, entry))
        else:
            passthrough.append((ident, entry))
    adjusted, aligned = {}, set()
    for gkey, members in groups.items():
        base_ks = base_group_ords.get(gkey) or set()
        targets = {}
        usable = bool(base_ks)
        for ident, entry in members:
            k = _serial_ordinal(entry[1].get(mapping.get(key_field, key_field)))
            if k is None or k not in base_ks or k in targets:
                usable = False
                break
            targets[k] = entry
        if usable:
            aligned.add(gkey)
            for k, entry in targets.items():
                position, values, display = entry
                adjusted[("grp", gkey, k)] = (position, values, [display[0], k])
        else:
            for ident, entry in members:
                adjusted[ident] = entry
    for ident, entry in passthrough:
        adjusted[ident] = entry
    return adjusted, aligned


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
        group_field = payload.get("group_field")
        group_mode = False
        auto_block = False
        if group_field is not None:
            if not (isinstance(group_field, int) and not isinstance(group_field, bool)
                    and group_field in current.selection.fields):
                plan.issue("field_selection", "group_field", "分组列必须是底表当前字段范围内的一个列")
                return plan
            if len(keys) != 1:
                plan.issue("field_selection", "group_key", "分组匹配模式下，主键只能选择一个序号列")
                return plan
            if group_field in keys or group_field in updates:
                plan.issue("field_selection", "group_conflict", "分组列不能同时作为主键或回填字段")
                return plan
            group_mode = True
            plan.warnings.append(
                "分组匹配：反馈中每组的第 k 条记录对应底稿同名组的第 k 条（序号数值不要求一致）；"
                "若反馈在组内中间插行或删行，结果需人工核对。")
        elif len(keys) == 1 and _key_column_merged(current, keys[0]):
            # 名称类主键常纵向合并: 合并块本身是一条记录, 块内行是子行。
            # 按块匹配是使用该列作主键的唯一合理读法, 自动启用并留痕。
            auto_block = True
            group_mode = True
            group_field = keys[0]
            plan.warnings.append(
                f"主键列{' / '.join(current.headers[keys[0]])}含合并单元格：已按“每个合并块 = 一条记录”匹配"
                "（块内第 k 行对应反馈同名块的第 k 行）。若该列只是分组标签，"
                "请改用“分组（单位）”选这一列，并另选每行唯一的序号列作主键。")
        if group_mode:
            records = _group_records(plan, current, keys[0], group_field,
                                     {f: f for f in current.selection.fields}, "current")
        else:
            records = _indexed(plan, current, keys, {f: f for f in current.selection.fields}, "current")
        plan.record_count = len(records)
        baseline = current
        base_map = {f: f for f in current.selection.fields}
        base_records = records
        if payload.get("baseline"):
            baseline = Source(payload["baseline"], payload.get("selection"), plan.books, plan.warnings)
            if group_mode:
                base_map = _mapping(plan, baseline, sorted(set(keys + updates + compares + [group_field])))
                base_records = _group_records(plan, baseline, keys[0], group_field, base_map, "baseline")
            else:
                base_map = _mapping(plan, baseline, sorted(set(keys + updates + compares)))
                base_records = _indexed(plan, baseline, keys, base_map, "baseline")
        items = payload.get("feedback", [])
        if not isinstance(items, list) or not items:
            plan.issue("feedback_required", "feedback", "请至少选择一个反馈工作簿")
            return plan
        proposals = defaultdict(list)
        seen_imports = set()
        # 删除检查(分组模式): 反馈某组只交回部分记录 → 缺失条疑似被删除
        base_group_ords = defaultdict(set)
        if group_mode:
            for ident in records:
                if isinstance(ident, tuple) and ident[0] == "grp":
                    base_group_ords[ident[1]].add(ident[2])
        suspected_deleted = set()
        for index, item in enumerate(items):
            _cancelled(cancel)
            feedback = Source(item, payload.get("selection"), plan.books, plan.warnings)
            if feedback.token in seen_imports:
                plan.duplicate_imports += 1
                continue
            seen_imports.add(feedback.token)
            required = sorted(set(keys + updates + ([group_field] if group_mode else [])))
            mapping = _mapping(plan, feedback, required)
            full_map = None
            if group_mode:
                found = _group_records(plan, feedback, keys[0], group_field, mapping, "feedback")
                found, serial_aligned = _align_group_serials(found, base_group_ords, keys[0], mapping)
                base_groups = {k[1] for k in records if isinstance(k, tuple) and k[0] == "grp"}
                labels = {}
                for identity, (_, _, display) in found.items():
                    if isinstance(identity, tuple) and identity[0] == "grp" and identity[1] not in base_groups:
                        labels.setdefault(identity[1], display[0])
                for _, label in sorted(labels.items(), key=lambda pair: str(pair[1])):
                    plan.warnings.append(
                        f"{feedback.label} 含底表中没有的分组：{label}，该组记录将作为新记录处理。")
            else:
                found = _indexed(plan, feedback, keys, mapping, "feedback", known_keys=set(records))
            if group_mode:
                fb_groups = defaultdict(set)
                for ident in found:
                    if isinstance(ident, tuple) and ident[0] == "grp":
                        fb_groups[ident[1]].add(ident[2])
                for gkey, fb_ks in fb_groups.items():
                    base_ks = base_group_ords.get(gkey, set())
                    missing_ks = base_ks - fb_ks
                    if not missing_ks:
                        continue
                    label = found[("grp", gkey, min(fb_ks))][2][0]
                    for k in missing_ks:
                        suspected_deleted.add(("grp", gkey, k))
                    if gkey in serial_aligned:
                        # 按主键对位成功: 序号缺口精确点名被删的记录, 其余对位不受影响
                        for k in sorted(missing_ks):
                            base_serial = records[("grp", gkey, k)][1].get(keys[0])
                            plan.warnings.append(
                                f"{feedback.label} 的组「{label}」缺少序号 {base_serial}"
                                f"（底稿组内第 {k} 条），可能已被删除；其余记录已按序号对位。")
                    else:
                        # 组内按"第 k 条"对位: 中间删行会让后续记录整体滑位,
                        # 无法断定缺的是哪一条, 因此报数量差并提示核对整组
                        plan.warnings.append(
                            f"{feedback.label} 的组「{label}」交回 {len(fb_ks)} 条，底稿有 {len(base_ks)} 条，"
                            f"少 {len(missing_ks)} 条——可能有记录被删除；该组按第 k 条对位，"
                            "对应关系可能整体错位，请逐条核对该组的改动。")
            if len(mapping) != len(required):
                continue
            for key, (position, values, key_values) in found.items():
                plan.feedback_count += 1
                record_map = mapping
                if key not in records:
                    # Confirming a new task appends the whole task, including its
                    # identity/context fields, not just the selected update fields.
                    if full_map is None:
                        full_map = _mapping(plan, feedback, list(current.selection.fields))
                    record_map = full_map
                    if len(record_map) != len(current.selection.fields):
                        continue
                proposals[key].append((feedback, position, values, key_values, record_map))
            if progress:
                progress(index + 1, len(items), feedback.label, plan.feedback_count)
        plan.missing = [{"key": values, "position": position, "status": "not_received",
                         "action": "keep_current",
                         "message": ("该组其余记录已交回而此条缺失，可能已被删除；已保留当前内容"
                                     if key in suspected_deleted else "本次未反馈，保留当前内容")}
                        for key, (position, _, values) in records.items() if key not in proposals]
        final = {key: (pos, dict(values), key_values, current.descriptor())
                 for key, (pos, values, key_values) in records.items()}
        appended = current.selection.data_end
        append_reason = None
        append_checked = False
        merged_skips = []
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
                # 多份反馈冲突: 默认保留当前内容, 冲突明细进变更清单与提示, 不再阻断
                if exists:
                    current_values = {f: records[key][1].get(f) for f in updates}
                    plan.changes.append({
                        "id": "conflict:" + _digest(key), "key": chosen[2], "field": None,
                        "field_title": "多份反馈冲突（已保留当前内容）",
                        "before": current_values, "after": current_values, "kind": "conflict",
                        "source": chosen[0].descriptor(), "position": records[key][0], "marked": False,
                        "candidates": [{"source": src.descriptor(), "values": vals}
                                       for src, vals, _ in unique.values()]})
                    plan.warnings.append(
                        f"记录 {' / '.join(str(v) for v in chosen[2])} 收到 {len(unique)} 份不同反馈，已保留当前内容（详见变更清单）。")
                    continue
                plan.warnings.append(
                    f"新记录 {' / '.join(str(v) for v in chosen[2])} 收到 {len(unique)} 份不同反馈，已按第一份处理。")
            source, values, key_values = chosen
            if not exists:
                if not append_checked:
                    append_checked = True
                    append_reason = _append_reason(current)
                if append_reason:
                    # 追加不安全(数据区后有页脚/表对象/合并模板等): 跳过是唯一安全
                    # 动作, 自动执行并聚合提示, 不逐条要求人工确认
                    plan.additions_skipped += 1
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
                    # 遮盖区(非左上角)物理上写不进值: 唯一合理动作是跳过该格,
                    # 从待写入值中剔除以保持变更清单与实际写入一致
                    merged_skips.append(cell.coordinate)
                    existing.pop(field, None)
                    final[key] = (old_position, existing, key_values, source.descriptor())
                else:
                    plan.actions.append((position, field, values[field]))
        if plan.additions_skipped:
            plan.warnings.append(
                f"发现 {plan.additions_skipped} 条反馈记录不在当前表中，因{append_reason}，"
                "已自动跳过（未追加）；如需并入请调整当前表后重试")
        if merged_skips:
            head = "、".join(merged_skips[:4]) + ("…" if len(merged_skips) > 4 else "")
            plan.warnings.append(
                f"有 {len(merged_skips)} 处回填目标位于合并单元格的非左上角（{head}），"
                "无法独立写入，已跳过；如需回填这些列请先调整合并单元格")
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
                        # 红色填充是本次对照的呈现, 不是永久累积状态。值改回与基线一致时
                        # 恢复基线的填充色, 字体保持当前样子; 基线本来就是红的则保持红。
                        current_cell = current.selection.cell(current.ws, position, field)
                        fill = current_cell.fill
                        if (fill and fill.fill_type == "solid"
                                and str(fill.fgColor.rgb or "")[-6:].upper() == "FF0000"):
                            baseline_cell = baseline.selection.cell(baseline.ws, baseline_record[0],
                                                                    base_map[field])
                            plan.restore_colors.append((position, field, copy(baseline_cell.fill)))
                    continue
                marked = field in compares
                if isinstance(after, str) and after.strip() in ("取消", "已取消", "作废"):
                    kind = "cancel"
                elif _blank(before) and not _blank(after):
                    kind = "fill"        # 填写: 底稿为空, 反馈填入
                elif not _blank(before) and _blank(after):
                    kind = "clear"       # 清空: 底稿有值, 反馈删空
                else:
                    kind = "update"
                plan.changes.append({"id": "change:" + _digest([key, field]), "key": key_values,
                                     "field": field, "field_title": " / ".join(current.headers[field]),
                                     "before": before, "after": after, "kind": kind,
                                     "source": source, "position": position, "marked": marked,
                                     "baseline_missing": baseline_record is None})
                plan.red.append((position, field))   # 默认: 全部改动格红填充
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


def _auto_selection(path, sheet):
    """引擎侧自动识别某张表的范围(界面未逐张打开反馈/底稿各表时使用)。"""
    try:
        from . import reader
        from .preview import _suggest
        tables, rd = reader.list_tables(path, True)
        src = next((s for s, _ in tables if s is not None and s.sheet_name == sheet), None)
        if src is None:
            return None
        ws = rd.wb[sheet]
        suggestion, _recognition = _suggest(path, sheet, ws.max_row, ws.max_column)
        return suggestion
    except Exception:
        return None


def _sheet_payloads(payload):
    """多表回填: payload["sheets"] 为逐表配置列表, 反馈缺省选区时自动识别。
    返回 (逐表子载荷, 未参与表及原因)。"""
    sheets = payload.get("sheets") or []
    out, skipped = [], []
    for entry in sheets:
        sub = dict(payload)
        sub["sheet"] = entry["sheet"]
        for key in ("selection", "key_fields", "update_fields", "compare_fields",
                    "group_field", "result_form"):
            if entry.get(key) is not None:
                sub[key] = entry[key]
        feedback, ok = [], True
        for item in payload.get("feedback", []):
            one = dict(item)
            one["sheet"] = entry["sheet"]
            known = entry.get("feedback_selections", {}).get(item["path"])
            if known:
                one["selection"] = known
            else:
                auto = _auto_selection(item["path"], entry["sheet"])
                if auto is None:
                    skipped.append((entry["sheet"], f"反馈文件无法识别该表：{item['path']}"))
                    ok = False
                    break
                one["selection"] = auto
            feedback.append(one)
        if ok:
            sub["feedback"] = feedback
            sub.pop("sheets", None)
            out.append((entry["sheet"], sub))
    return out, skipped


def _merge_sheet_reports(reports):
    """合并逐表回填报告: 问题/警告带表名前缀, 变化标注来源表。"""
    merged = {"changes": [], "issues": [], "missing": [], "warnings": [],
              "can_export": True, "fields": [], "totals": {}}
    totals = {}
    for sheet, report in reports:
        tag = f"[{sheet}] "
        merged["changes"] += [dict(c, sheet=sheet) for c in report.get("changes", [])]
        merged["issues"] += [dict(i, **({"message": tag + str(i.get("message", ""))})
                                   if i.get("message") else {}) for i in report.get("issues", [])]
        merged["missing"] += [dict(m, sheet=sheet) for m in report.get("missing", [])]
        merged["warnings"] += [tag + w if not w.startswith("[") else w for w in report.get("warnings", [])]
        merged["can_export"] = merged["can_export"] and bool(report.get("can_export"))
        merged["fields"] += report.get("fields", [])
        for key, value in (report.get("totals") or {}).items():
            if isinstance(value, (int, float)):
                totals[key] = totals.get(key, 0) + value
    merged["totals"] = totals
    return merged


def reconcile_multi_preview(payload, cancel=None, progress=None):
    entries, skipped = _sheet_payloads(payload)
    reports = [(sheet, {"changes": [], "issues": [{"message": reason, "choices": []}],
                        "missing": [], "warnings": [], "can_export": False, "totals": {}})
               for sheet, reason in skipped]
    for index, (sheet, sub) in enumerate(entries):
        _cancelled(cancel)
        report = reconcile_preview(sub)
        reports.append((sheet, report))
        if progress:
            progress(index + 1, len(payload.get("sheets") or []), sheet, 0)
    return _merge_sheet_reports(reports)


def reconcile_multi_export(payload, cancel=None, progress=None):
    """逐表链式导出: 每张表的回填以上一张表的结果为底, 最终合成一个文件。"""
    import tempfile
    out = Path(payload["out"]).expanduser().resolve()
    entries, skipped = _sheet_payloads(payload)
    if skipped:
        raise ValueError("；".join(f"表「{sheet}」{reason}" for sheet, reason in skipped))
    if not entries:
        raise ValueError("没有可回填的工作表配置")
    import shutil
    with tempfile.TemporaryDirectory(prefix=".excel-multi-reconcile-") as temp:
        current_input = None
        reports = []
        for index, (sheet, sub) in enumerate(entries):
            _cancelled(cancel)
            staged = Path(temp) / f"stage-{index}.xlsx"
            if current_input is not None:
                # 上一张表的回填结果作为本表底稿, 反馈仍是原始反馈文件
                sub = dict(sub, path=str(current_input))
            sub_result = reconcile_export(dict(sub, out=str(staged)), cancel, progress)
            if not sub_result.get("exported"):
                return sub_result
            reports.append((sheet, sub_result))
            current_input = staged
        result = _merge_sheet_reports(reports)
        log = out.with_suffix(".changes.json")
        if out.exists() or log.exists():
            raise ValueError("输出文件或变更清单已存在，请选择新名称")
        shutil.move(str(current_input), out)
        audit = {**result, "out": str(out), "changes_path": str(log), "exported": True}
        log.write_text(json.dumps(audit, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return audit


def reconcile_preview(payload):
    """Return JSON-safe changes/issues without changing any workbook."""
    if payload.get("sheets"):
        return reconcile_multi_preview(payload, cancel=None, progress=None)
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


def _pick_marker_colors(ws, selection):
    """自动选标记色: 扫描数据区现有纯色填充, 为"新值/原值"各挑一个
    与现有颜色(及彼此)距离足够大的候选色, 避免与模板撞色导致提示不明显。"""
    def rgb_bytes(color):
        try:
            text = str(color.rgb or "")
            return bytes.fromhex(text[-6:]) if len(text) >= 6 else None
        except ValueError:
            return None

    def distance(a, b):
        return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5

    existing = set()
    if ws and selection:
        for row in ws.iter_rows(min_row=selection.data_start if selection.axis == "rows" else 1,
                                max_row=selection.data_end if selection.axis == "rows" else ws.max_row,
                                min_col=1 if selection.axis == "rows" else selection.data_start,
                                max_col=ws.max_column if selection.axis == "rows" else selection.data_end):
            for cell in row:
                fill = cell.fill
                if fill and fill.fill_type == "solid":
                    rgb = rgb_bytes(fill.fgColor)
                    if rgb:
                        existing.add(rgb)
    new_candidates = ["FFFF0000", "FFE74C3C", "FFC0392B", "FFD63031"]
    old_candidates = ["FFF39C12", "FFE67E22", "FFFFC000", "FFF5A623"]
    def pick(cands, avoid):
        for cand in cands:
            rgb = bytes.fromhex(cand[-6:])
            if all(distance(rgb, x) >= 110 for x in avoid) and all(distance(rgb, x) >= 110 for x in existing):
                return cand
        return cands[0]
    color_new = pick(new_candidates, [])
    color_old = pick(old_candidates, [bytes.fromhex(color_new[-6:])])
    return color_new, color_old


def _insert_compare_rows(plan):
    """保留对比输出: 每条改动记录下方插入完整"原值行"(与该行格式一致),
    新值格红底、原值格橙底; 疑似删除行整行橙底; 新增行整行红底。
    行排列插行 / 列排列插列; 跨越插入点的合并区按累计位移重建。"""
    ws = plan.current.ws
    sel = plan.current.selection
    rows_axis = sel.axis == "rows"
    fields = range(sel.field_start, sel.field_end + 1)
    changed_fields = {}
    before_values = {}
    for change in plan.changes:
        if change.get("field") is not None and change.get("kind") not in ("new", "conflict"):
            changed_fields.setdefault(change["position"], set()).add(change["field"])
            before_values[(change["position"], change["field"])] = change.get("before")
    color_new, color_old = plan.marker_colors

    # 1) 快照改动记录整行的当前样式(新值已写入; 原值逐字段用 before 恢复)
    snapshots = {}
    for pos in changed_fields:
        snapshots[pos] = {f: (sel.cell(ws, pos, f).value, copy(sel.cell(ws, pos, f)._style))
                          for f in fields}

    # 2) 拆开全部合并区, 记录原始坐标
    original_merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in list(ws.merged_cells.ranges)]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    # 3) 自下而上插入原值行/列并回填
    insert_points = sorted(changed_fields, reverse=True)
    for pos in insert_points:
        P = pos + 1
        if rows_axis:
            ws.insert_rows(P)
            ws.row_dimensions[P].height = ws.row_dimensions[pos].height
        else:
            ws.insert_cols(P)
            src_dim, dst_dim = ws.column_dimensions[get_column_letter(pos)], ws.column_dimensions[get_column_letter(P)]
            dst_dim.width = src_dim.width
        for f in fields:
            cell = sel.cell(ws, P, f)
            value, style = snapshots[pos][f]
            cell.value = before_values.get((pos, f), value)   # 改动格恢复原值
            cell._style = copy(style)
            if f in changed_fields[pos]:
                cell.fill = PatternFill(fill_type="solid", fgColor=color_old)

    # 4) 合并区按累计位移重建: 新位置 = 原位置 + (插入点中 <= 原位置 的数量)
    def shifted(x):
        return x + sum(1 for P in (pos + 1 for pos in insert_points) if P <= x)

    for (r1, c1, r2, c2) in original_merges:
        if rows_axis:
            ws.merge_cells(start_row=shifted(r1), start_column=c1, end_row=shifted(r2), end_column=c2)
        else:
            ws.merge_cells(start_row=r1, start_column=shifted(c1), end_row=r2, end_column=shifted(c2))

    # 5) 删除行整行橙底 / 新增行整行红底(位置按同一位移规则换算)
    for entry in plan.missing:
        if "被删除" in entry.get("message", "") and entry.get("position"):
            for f in fields:
                sel.cell(ws, shifted(entry["position"]), f).fill = PatternFill(
                    fill_type="solid", fgColor=color_old)
    for pos in plan.additions:
        for f in fields:
            sel.cell(ws, shifted(pos), f).fill = PatternFill(fill_type="solid", fgColor=color_new)


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
    if payload.get("sheets"):
        return reconcile_multi_export(payload, cancel, progress)
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
        form = payload.get("result_form") or "compare"
        plan.marker_colors = _pick_marker_colors(plan.current.ws, plan.current.selection)
        _append_template(plan.current, plan.additions)
        for position, field, value in plan.actions:
            _cancelled(cancel)
            plan.current.selection.cell(plan.current.ws, position, field).value = value
        for position, field, fill in plan.restore_colors:
            plan.current.selection.cell(plan.current.ws, position, field).fill = fill
        for position, field in plan.red:
            # 改动标红 = 红色填充(字体原样保留), 颜色经自动选色避免与模板撞色
            plan.current.selection.cell(plan.current.ws, position, field).fill = PatternFill(
                fill_type="solid", fgColor=plan.marker_colors[0])
        if form == "compare":
            _insert_compare_rows(plan)
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
