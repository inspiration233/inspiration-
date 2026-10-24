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
from .align import AlignError, align_sources, key_parts
from .recognize import (DEFAULT_SUMMARY_KEYWORDS, MergeRange, is_numberlike,
                       is_summary_row, normalize_name)
from .writer import OutputBook, XLSX_MAX_ROWS

MODES = ("single", "workbook", "sheets", "group")


class Cancelled(Exception):
    pass


@dataclass
class Options:
    mode: str = "single"
    align: str = "loose"                    # strict | loose
    style: str = "fidelity"                 # fidelity(原生保真,默认) | fast(高速流式)
    auto_group: bool = True                 # 异构表自动分组到不同 Sheet
    merge_shape: str = "auto"               # auto | stack(数据汇总) | append(完整表拼接)
    orientation_overrides: dict = field(default_factory=dict)  # {display: vertical/horizontal}
    template_index: int = 0                 # 表头模板 = 第几个源（0 起）
    remove_summary: bool = True
    summary_keywords: tuple = DEFAULT_SUMMARY_KEYWORDS
    strip_serial_rows: bool = True
    min_fill_ratio: float = 0.20  # 行内非空格占比低于此值 → 剔除(默认20%)          # 剔除"仅有序号"的空行
    renumber_serial: bool = True            # 序号列重排(合并后连续编号)
    source_column: bool = False
    include_hidden: bool = False
    title_as_header: bool = False           # 以模板标题作总表标题（默认关）
    force_header_row: int = 0               # 0=自动探测
    selections: list = field(default_factory=list)
    preserve_blank_rows: bool = False       # 新工作台默认开启；兼容旧 CLI
    preserve_template_preamble: bool = False
    metadata_columns: bool = True
    exact_headers: bool = False
    template_path: str = ""
    template_sheet: str = ""


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
                   "big_numbers_converted": 0,
                   "formulas_translated": 0, "formulas_to_value": 0,
                   "validations_kept": 0, "validations_dropped": 0,
                   "cf_kept": 0, "cf_dropped": 0, "row_heights_set": 0,
                   "tables_kept": 0, "tables_downgraded": 0,
                   "group_count": 1},
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
            "orientation": getattr(st, "orientation", "vertical"),
            "orientation_confident": getattr(st, "orientation_confident", True),
            "native_cross": getattr(st, "native_cross", False),
            "header_row": st.header_row, "height": st.height,
            "data_end_row": getattr(st, "data_end_row", 0),
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
            ov = options.orientation_overrides or {}
            for src, skip_reason in tables:
                if src is None:
                    entry.sheets.append({"sheet": skip_reason, "skipped": True})
                    report["totals"]["hidden_sheets_skipped"] += 1
                    continue
                if src.display in ov or options.force_header_row:
                    src.structure = reader.build_structure(
                        src.structure.sheet_name, options.force_header_row,
                        ov.get(src.display))
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
                from .selection import source_spec, apply_selection
                spec = source_spec(options, p, src.sheet_name)
                if spec is False:
                    continue
                override = options.orientation_overrides.get(src.display)
                if override or options.force_header_row:
                    src.structure = reader.build_structure(src.sheet_name, options.force_header_row, override)
                apply_selection(src, spec)
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

        if options.template_path:
            resolved_template = Path(options.template_path).resolve()
            template = next((s for s in sources if Path(s.path).resolve() == resolved_template
                             and s.sheet_name == options.template_sheet), None)
            if template is None:
                if not sources:
                    # 没有任何可用来源: 真正的原因几乎总是"表里没有数据", 不能让人
                    # 去检查模板设置
                    skipped_names = [entry.get("sheet") for group in report.get("groups", [])
                                     for entry in group.get("sheets", []) if entry.get("skipped")]
                    detail = "、".join(str(n) for n in skipped_names[:4]) if skipped_names else "所有工作表"
                    raise ValueError(
                        f"没有可合并的数据：{detail} 整表为空（只有格式没有内容）。"
                        "请打开文件确认里面是否有数据；若数据在别的文件/工作表，请改选那份。")
                raise ValueError("所选版式参考表不在合并来源中，请重新选择")
            options.template_index = sources.index(template)
        from .selection import apply_field_maps
        apply_field_maps(sources, options.template_index)
        usable = [s for s in sources if not _is_empty_source(s)]
        for s in sources:
            if _is_empty_source(s):
                report["warnings"].append(f"「{s.display}」仅有表头无数据，已跳过")
                totals["empty_sources"] += 1

        if not usable:
            report["warnings"].append("没有可合并的数据")
            return report

        options.style = {"unified": "fast", "preserve": "fidelity"}.get(
            options.style, options.style)

        # 1.5) 同名同内容Sheet去重: 两个文件中名字和内容完全相同的Sheet只保留一份
        # (典型的配置/规格表场景: 每个单位文件里附同一份参数表)
        from . import reader as _rd
        def _sheets_identical(s1, s2):
            if s1.structure.keys != s2.structure.keys:
                return False
            try:
                r1 = _rd.open_reader(s1.path)
                r2 = _rd.open_reader(s2.path)
                w1, w2 = r1.wb[s1.sheet_name], r2.wb[s2.sheet_name]
                if w1.max_row != w2.max_row or w1.max_column != w2.max_column:
                    r1.close(); r2.close()
                    return False
                identical = all(
                    w1.cell(r, c).value == w2.cell(r, c).value
                    for r in range(1, w1.max_row + 1)
                    for c in range(1, w1.max_column + 1)
                )
                r1.close(); r2.close()
                return identical
            except Exception:
                return False

        deduped = []
        seen_sheets = []
        for s in usable:
            is_dup = False
            # 原生交叉表不去重: 横向拼接会把各来源记录列都追加到右侧
            if getattr(s.structure, "native_cross", False):
                deduped.append(s)
                continue
            for prev in seen_sheets:
                if (s.structure.sheet_name == prev.structure.sheet_name
                        and _sheets_identical(s, prev)):
                    report["warnings"].append(
                        f"「{s.display}」与「{prev.display}」内容完全相同，已去重保留一份")
                    is_dup = True
                    break
            if not is_dup:
                deduped.append(s)
                seen_sheets.append(s)
        usable = deduped

        # 2) 分组：workbook=每源一页；group=按Sheet名；single/sheets=按结构签名异构分组
        def _sig(s):
            return (frozenset(s.structure.keys), s.structure.height)

        def _seg_set(s):
            signatures = getattr(s.structure, "key_signatures", None)
            return {key_parts(k, signature=signatures is not None)[-1]
                    for k in (signatures if signatures is not None else s.structure.keys)}

        def _by_signature(sources):
            # 末段键重叠度聚类: >50% 视为同构(宽容对齐可处理小差异),
            # 结构迥异(统计表 vs 明细)自动分组
            clusters = []
            for s in sources:
                ss = _seg_set(s)
                best, best_r = None, 0.0
                for gs, members in clusters:
                    r = len(ss & gs) / max(len(ss), len(gs), 1)
                    if r > best_r:
                        best, best_r = (gs, members), r
                if best is not None and best_r >= 0.5:
                    best[1].append(s)
                else:
                    clusters.append((ss, [s]))
            return [m for _, m in clusters]

        if options.mode == "workbook":
            groups = [(s.structure.sheet_name, [s]) for s in usable]
        elif options.mode == "group":
            by_name = {}
            for s in usable:
                by_name.setdefault(s.structure.sheet_name, []).append(s)
            groups = []
            for bucket in by_name.values():
                # 同名Sheet=同一模板: 任一成员识别为交叉表 → 全组横向拼接,
                # 不按签名再拆(记录列名不同是常态, 拆开反而丢拼接)
                if any(getattr(s.structure, "native_cross", False) for s in bucket):
                    subs = [bucket]
                else:
                    subs = _by_signature(bucket) if options.auto_group else [bucket]
                groups.extend((b[0].structure.sheet_name, b) for b in subs)
        else:  # single / sheets
            subs = _by_signature(usable) if options.auto_group else [usable]
            stems = [b[0].path.stem for b in subs]
            groups = []
            for i, b in enumerate(subs):
                if i == 0:
                    name = "总表"
                elif stems.count(b[0].path.stem) > 1:
                    name = b[0].structure.sheet_name      # 同一文件拆出的组用Sheet名
                else:
                    name = b[0].path.stem                 # 跨文件的组用文件名
                groups.append((name, b))
        # 合并形态: 单源组=拼接(完整还原); 多源组按前置区共性推荐(全同→汇总, 有差异→拼接保形)
        def _shape_of(b):
            if options.merge_shape in ("stack", "append"):
                return options.merge_shape
            if len(b) == 1:
                return "append"                    # 单源组: 完整还原(含前置区)
            # 多源: 一律汇总(一个表头+数据堆叠+前置区逐行共性检测);
            # 表头相同还逐张重复属冗余——拼接式仅在用户明确选择时使用
            return "stack"

        groups = [(n, b, _shape_of(b)) for (n, b) in groups]
        totals["group_count"] = len(groups)
        report["groups"] = [{"sheet": n, "sources": [s.display for s in b],
                             "shape": sh} for n, b, sh in groups]
        if len(groups) > 1:
            report["warnings"].append(
                f"检测到 {len(groups)} 种不同表结构，已自动分组到不同 Sheet（详见报告）")

        # 3) 逐组合并写出（保真 / 高速 两条路径）
        if options.style == "fidelity":
            from . import fidelity as fid
            book = fid.FidelityBook(out_path)
            try:
                for name, grp, shape in groups:
                    reference = sources[min(options.template_index, len(sources) - 1)]
                    group_template = grp.index(reference) if reference in grp else 0
                    fid.merge_group_fidelity(book, grp, options, group_template, report,
                                             progress, cancel, base_name=name,
                                             shape=shape)
                book.save()
            except (Cancelled, KeyboardInterrupt):
                try:
                    book.save()
                except Exception:
                    pass
                report["cancelled"] = True
                report["warnings"].append("已取消：输出文件为取消前已完成的部分")
                totals["big_numbers_converted"] += book.big_numbers_converted
                report["output"] = str(out_path)
                report["duration_sec"] = round(time.monotonic() - t0, 2)
                report = _finalize(report)
                _write_log(out_path, report)
                return report
            totals["big_numbers_converted"] += book.big_numbers_converted
        else:
            book = OutputBook(out_path, "fast")
            try:
                for name, group, _shape in groups:
                    template_index = min(options.template_index, len(group) - 1)
                    _merge_group(book, group, options, template_index,
                                 report, progress, cancel, base_name=name)
            except Cancelled:
                try:
                    book.save()                     # 落盘取消前已完成的部分
                except Exception:
                    pass
                report["cancelled"] = True
                report["warnings"].append("已取消：输出文件为取消前已完成的部分")
                totals["big_numbers_converted"] = book.big_numbers_converted
                report["output"] = str(out_path)
                report["duration_sec"] = round(time.monotonic() - t0, 2)
                report = _finalize(report)
                _write_log(out_path, report)
                return report
            except AlignError:
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


def _merge_group(book, group, options, template_index, report, progress, cancel,
                 base_name=None):
    totals = report["totals"]
    base_name = base_name or (group[0].structure.sheet_name
                              if options.mode in ("workbook", "group") else "总表")

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

    # 高速路径：仅工作簿拷贝模式还原数据区结构（保真路径独立实现）
    keep_look = options.mode == "workbook"
    if keep_look:
        inv0 = {si: oi for oi, si in enumerate(maps[0]) if si is not None}
        widths = {inv0[c - 1] + 1: w for c, w in template.structure.col_widths.items()
                  if inv0.get(c - 1) is not None}
        if widths:
            page.set_col_widths(widths)

    # 序号列重排: 输出键为 序号/编号 的列用连续计数覆盖(跨文件连续)
    serial_idx = next((i for i, k in enumerate(out_keys)
                       if k in ("序号", "编号")), None)         if options.renumber_serial else None
    serial_counter = 0

    rows_total = 0
    for fi, (src, mapping) in enumerate(zip(src_list, maps)):
        data_rows = 0
        row_map = {}                       # 源行号 → (分页代际, 输出行号)
        dstart = src.structure.header_row + src.structure.height
        for row_no, cells in src.iter_data_rows():
            if cancel and cancel():
                raise Cancelled()
            values = [getattr(c, "value", c) for c in cells]
            if all(v in (None, "") for v in values) and not options.preserve_blank_rows:
                totals["empty_rows_skipped"] += 1
                continue
            if options.strip_serial_rows:
                non_empty = [(i, v) for i, v in enumerate(values) if v not in (None, "")]
                if (len(non_empty) == 1 and is_numberlike(non_empty[0][1])):
                    totals["serial_rows_removed"] = totals.get("serial_rows_removed", 0) + 1
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
            if serial_idx is not None and serial_idx < len(out_cells):
                serial_counter += 1
                out_cells[serial_idx] = _Renumber(
                    serial_counter, out_cells[serial_idx])
            page.write_row(out_cells)
            rows_total += 1
            data_rows += 1
            row_map[row_no] = (page.page_epoch, page.row_index)
        # 数据区合并还原：把源合并范围翻译到输出坐标（同页且行连续才有效）
        if keep_look and data_rows:
            inv_map = {si: oi for oi, si in enumerate(mapping) if si is not None}
            kept = 0
            for mg in getattr(src.structure, "all_merges", []):
                if mg.min_row < dstart:
                    continue
                e1, r1 = row_map.get(mg.min_row, (None, None))
                e2, r2 = row_map.get(mg.max_row, (None, None))
                if e1 is None or e1 != e2 or r2 - r1 != mg.max_row - mg.min_row:
                    continue
                c1, c2 = inv_map.get(mg.min_col - 1), inv_map.get(mg.max_col - 1)
                if c1 is None or c2 is None or c2 - c1 != mg.max_col - mg.min_col:
                    continue
                page.add_extra_merge(c1 + 1, r1, c2 + 1, r2)
                kept += 1
            if kept:
                totals["data_merges_preserved"] = \
                    totals.get("data_merges_preserved", 0) + kept
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


class _Renumber:
    """序号重排值：新序号 + 原单元格(保留其样式)。"""
    __slots__ = ("value", "_style_src")

    def __init__(self, value, style_src):
        self.value = value
        self._style_src = style_src if hasattr(style_src, "value") else None


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
