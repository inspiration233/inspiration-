# -*- coding: utf-8 -*-
"""四种合并模式的编排 + 结果报告。需求 §3.1–§3.5。

模式：
  single   A. 多文件 → 一张总表
  workbook B. 多文件 → 一个工作簿（每源一页）
  sheets   C. 文件内多 Sheet → 一张总表（与 A 同实现，语义上单文件）
  group    D. 多文件按 Sheet 名分组汇总
"""
import datetime as dt
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path

from . import reader as rd
from .align import AlignError, align_sources
from .recognize import DEFAULT_SUMMARY_KEYWORDS, MergeRange, is_summary_row, normalize_name
from .writer import OutputBook, XLSX_MAX_ROWS

MODES = ("single", "workbook", "sheets", "group")


class Cancelled(Exception):
    pass


@dataclass
class Options:
    mode: str = "single"
    align: str = "loose"                    # strict | loose
    style: str = "unified"                  # unified | preserve
    template_index: int = 0                 # 表头模板 = 第几个源（0 起）
    remove_summary: bool = True
    summary_keywords: tuple = DEFAULT_SUMMARY_KEYWORDS
    source_column: bool = False
    include_hidden: bool = False
    title_as_header: bool = False           # 以模板标题作总表标题（默认关）
    force_header_row: int = 0               # 0=自动探测


@dataclass
class _FileEntry:
    path: str
    status: str = "ok"                      # ok | failed
    reason: str = ""
    sheets: list = field(default_factory=list)   # 每 sheet 的体检/处理记录


def _new_report(mode, options):
    return {
        "mode": mode,
        "align": options.align,
        "style": options.style,
        "files": [],
        "totals": {"files_ok": 0, "files_failed": 0, "sheets_merged": 0,
                   "rows_written": 0, "summary_rows_removed": 0, "empty_rows_skipped": 0,
                   "empty_sources": 0, "hidden_sheets_skipped": 0,
                   "big_numbers_converted": 0},
        "warnings": [],
        "header_detections": [],            # 每源: display/header_row/height/title/低置信/警告
        "duration_sec": 0.0,
    }


def _finalize(report):
    """报告转为纯 JSON 结构（dataclass → dict）。"""
    report["files"] = [asdict(f) for f in report["files"]]
    return report


def _sheet_record(src):
    st = src.structure
    return {"display": src.display, "sheet": st.sheet_name,
            "header_row": st.header_row, "height": st.height,
            "columns": list(st.keys), "title": st.title,
            "low_confidence": st.low_confidence, "warnings": list(st.warnings)}


# --------------------------------------------------------------- 体检（scan）

def scan(paths, options=None):
    options = options or Options()
    report = _new_report("scan", options)
    readers = []
    try:
        for p in paths:
            entry = _FileEntry(str(p))
            report["files"].append(entry)
            try:
                tables, reader = rd.list_tables(p, options.include_hidden)
                readers.append(reader)
            except rd.FileError as e:
                entry.status, entry.reason = "failed", e.reason
                report["totals"]["files_failed"] += 1
                continue
            report["totals"]["files_ok"] += 1
            for src, skip_reason in tables:
                if src is None:
                    entry.sheets.append({"sheet": skip_reason, "skipped": True})
                    report["totals"]["hidden_sheets_skipped"] += 1
                    continue
                entry.sheets.append(_sheet_record(src))
                report["header_detections"].append(_sheet_record(src))
    finally:
        for r in readers:
            r.close()
    return _finalize(report)


# --------------------------------------------------------------- 合并（merge）

def merge(paths, out_path, options=None, progress=None, cancel=None):
    options = options or Options()
    if options.mode not in MODES:
        raise ValueError(f"未知合并模式 {options.mode}")
    t0 = time.monotonic()
    report = _new_report(options.mode, options)
    totals = report["totals"]
    readers, sources = [], []
    try:
        # 1) 打开所有文件 → 源列表
        for p in paths:
            entry = _FileEntry(str(p))
            report["files"].append(entry)
            try:
                tables, reader = rd.list_tables(p, options.include_hidden)
                readers.append(reader)
            except rd.FileError as e:
                entry.status, entry.reason = "failed", e.reason
                totals["files_failed"] += 1
                continue
            totals["files_ok"] += 1
            real_tables = [t for t, _ in tables if t is not None]
            for src, skip_reason in tables:
                if src is None:
                    entry.sheets.append({"sheet": skip_reason, "skipped": True})
                    totals["hidden_sheets_skipped"] += 1
                    continue
                if len(real_tables) <= 1:
                    src.display = Path(p).stem    # 单 Sheet 文件: 来源只记文件名
                rec = _sheet_record(src)
                entry.sheets.append(rec)
                report["header_detections"].append(rec)
                report["warnings"].extend(
                    f"「{src.display}」{w}" for w in src.structure.warnings)
                if src.structure.empty:
                    entry.sheets.append({"sheet": src.structure.sheet_name,
                                         "skipped": True, "reason": "整表为空"})
                    continue
                sources.append(src)

        usable = [s for s in sources if not _is_empty_source(s)]
        for s in sources:
            if _is_empty_source(s):
                report["warnings"].append(f"「{s.display}」仅有表头无数据，已跳过")
                totals["empty_sources"] += 1

        if not usable:
            report["warnings"].append("没有可合并的数据")
            return report

        book = OutputBook(out_path, options.style)

        # 2) 按模式分表
        if options.mode == "workbook":
            groups = [[s] for s in usable]
        elif options.mode == "group":
            by_name = {}
            for s in usable:
                by_name.setdefault(s.structure.sheet_name, []).append(s)
            groups = list(by_name.values())
        else:  # single / sheets
            groups = [usable]

        # 3) 逐组合并写出
        for group in groups:
            template_index = min(options.template_index, len(group) - 1)
            try:
                _merge_group(book, group, options, template_index,
                             report, progress, cancel)
            except AlignError as e:
                raise
        book.save()
        totals["big_numbers_converted"] = book.big_numbers_converted
    finally:
        for r in readers:
            r.close()
    report["output"] = str(out_path)
    report["duration_sec"] = round(time.monotonic() - t0, 2)
    report = _finalize(report)
    _write_log(out_path, report)
    return report


def _is_empty_source(src):
    """仅表头无数据 → 结构上无法预知，保守由行迭代计数判定（在 _merge_group 记录）。"""
    return False


def _merge_group(book, group, options, template_index, report, progress, cancel):
    totals = report["totals"]
    base_name = (group[0].structure.sheet_name if options.mode in ("workbook", "group")
                 else "总表")

    if options.mode == "workbook":
        out_keys, maps = list(group[0].structure.keys), [[i for i in range(len(group[0].structure.keys))]]
        group, maps = [group[0]], maps
        src_list = group
    else:
        out_keys, maps = align_sources(group, options.align, template_index)
        src_list = group

    template = src_list[template_index] if options.mode != "workbook" else group[0]
    tkeys = template.structure.keys
    tpos = {k: i for i, k in enumerate(tkeys)}

    # 表头（模板原样 + 追加列 + 可选来源列/总表标题）
    header_rows = template.header_cells()
    height = max(1, len(header_rows))
    title_row = None
    if options.title_as_header and template.structure.title:
        title_row = [_Pseudo(template.structure.title)]
    extra_texts = []
    for k in out_keys:
        if k not in tpos:
            extra_texts.append(_extra_header(group, maps, out_keys.index(k), k))
    source_label = "来源" if options.source_column else None
    if title_row:
        header_rows = [title_row] + header_rows

    # 模板合并平移到输出坐标（含标题行整体 +1）
    row_shift = (template.structure.header_row - 1) - (1 if title_row else 0)
    merges = [MergeRange(mg.min_col, mg.min_row - row_shift, mg.max_col,
                         min(mg.max_row - row_shift, height))
              for mg in template.structure.merges]
    if title_row:
        width = max(max(len(r) for r in header_rows),
                    len(tkeys) + len(extra_texts) + (1 if source_label else 0))
        title_merges = [MergeRange(1, 1, width, 1)]
        body_merges = [MergeRange(mg.min_col, mg.min_row + 1, mg.max_col, mg.max_row + 1)
                       for mg in merges]
        merges = title_merges + body_merges

    page = book.add_table(base_name, header_rows, merges, extra_texts, source_label)
    page_capacity = XLSX_MAX_ROWS

    rows_total = 0
    for fi, (src, mapping) in enumerate(zip(src_list, maps)):
        data_rows = 0
        for cells in src.iter_data_rows():
            if cancel and cancel():
                raise Cancelled()
            values = [getattr(c, "value", c) for c in cells]
            if all(v in (None, "") for v in values):
                totals["empty_rows_skipped"] += 1
                continue
            if options.remove_summary and is_summary_row(
                    values, options.summary_keywords):
                totals["summary_rows_removed"] += 1
                continue
            out_cells = [(cells[si] if si is not None and si < len(cells) else None)
                         for si in mapping]
            if source_label:
                out_cells.append(_Pseudo(src.display))   # 来源列恒为文本值
            if page.row_index >= page_capacity - 2:
                page.paginate_restart(header_rows, merges, extra_texts, source_label)
                report["warnings"].append(
                    f"「{base_name}」超出单页行数上限，已自动续写新页")
            page.write_row(out_cells)
            rows_total += 1
            data_rows += 1
        if data_rows == 0:
            report["warnings"].append(f"「{src.display}」仅有表头无数据，已跳过")
            totals["empty_sources"] += 1
        totals["sheets_merged"] += 1 if data_rows else 0
        if progress:
            progress(fi + 1, len(src_list), src.display, rows_total)
    totals["rows_written"] += rows_total
    return page


def _extra_header(group, maps, out_idx, key):
    """追加列的表头文本：取第一个拥有该键的源的原样表头文字。"""
    for src, mapping in zip(group, maps):
        si = mapping[out_idx]
        if si is not None:
            for row in src.header_cells():
                cell = row[si] if si < len(row) else None
                if cell is not None and getattr(cell, "value", None) not in (None, ""):
                    return str(cell.value)
    return key


class _Pseudo:
    __slots__ = ("value", "number_format", "font", "fill", "border", "alignment")

    def __init__(self, value):
        self.value = value
        self.number_format = "General"
        self.font = None
        self.fill = None
        self.border = None
        self.alignment = None


# --------------------------------------------------------------- 日志（脱敏）

def _write_log(out_path, report):
    """运行日志：只记路径/统计/错误，不记任何单元格内容（安全底线）。"""
    p = Path(out_path)
    log = p.with_name(p.stem + "_运行日志.log")
    lines = [
        f"时间: {dt.datetime.now():%Y-%m-%d %H:%M:%S}",
        f"模式: {report['mode']}  对齐: {report['align']}  样式: {report['style']}",
        f"输出: {report.get('output', '')}  耗时: {report['duration_sec']}s",
        f"统计: {report['totals']}",
    ]
    for f in report["files"]:
        f = f if isinstance(f, dict) else asdict(f)
        lines.append(f"文件[{f['status']}]: {f['path']}" + (f"  原因: {f['reason']}" if f.get("reason") else ""))
        for s in f["sheets"]:
            lines.append(f"  - {s}")
    for w in report["warnings"]:
        lines.append(f"警告: {w}")
    try:
        log.write_text("\n".join(lines), encoding="utf-8")
    except OSError:
        pass
