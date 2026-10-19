# -*- coding: utf-8 -*-
"""原生保真模式引擎（需求 §3.3–§3.7，V2.0）。

与高速流式路径并行：openpyxl 普通模式整表载入（公式/验证/条件格式/行高/批注
只有普通模式读得到），普通模式写出（样式/验证/条件格式直接落盘，无需 XML 注入）。

保真清单：单元格样式 / 数字格式 / 列宽 / 行高 / 表头结构(按组模板) / 数据区合并 /
公式(引用按映射平移) / 数据验证(含配置页搬运) / 条件格式 / 批注 / 冻结窗格 /
打印设置 / 智能表格(单源保留, 多源降级自动筛选)。
"""
import re
from copy import copy as _copy, deepcopy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from .align import align_sources
from .recognize import is_numberlike, is_summary_row
from .writer import XLSX_MAX_ROWS, big_number_guard, sanitize_sheet_name


def _real(style):
    """StyleProxy → 真实样式对象（proxy 的 copy/deepcopy 会递归或携带源簿集合）。"""
    return getattr(style, "_StyleProxy__target", style)


def _copy_cell_style(dst, src):
    """跨工作簿样式复制：按属性注册到目标簿样式表。"""
    try:
        f = _real(src.font)
        dst.font = _copy(f)
        fill = _real(src.fill)
        if getattr(fill, "patternType", None):
            dst.fill = _copy(fill)
        dst.border = _copy(_real(src.border))
        dst.alignment = _copy(_real(src.alignment))
        fmt = src.number_format
        if fmt and fmt != "General":
            dst.number_format = fmt
    except Exception:
        pass

_REF_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$")
_SHEET_REF_RE = re.compile(r"^(?:'([^']+)'|([^!]+))!(.+)$")


# ================================================================ 公式引用平移

def _map_endpoint(ep, col_map, row_map):
    m = _REF_RE.match(ep)
    if not m:
        return None
    d1, col, d2, row = m.groups()
    oc = col_map.get(column_index_from_string(col.upper()))
    orr = row_map.get(int(row))
    if oc is None or orr is None:
        return None
    return f"{d1}{get_column_letter(oc)}{d2}{orr}"


def translate_formula(formula, col_map, row_map):
    """引用平移：RANGE token 按列映射+行映射重写；跨表引用/命名范围/整列范围 → None(转值)。"""
    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return None
    out = []
    for t in tokens:
        v = t.value
        if t.type == "OPERAND" and t.subtype == "RANGE":
            if "!" in v:
                return None                          # 跨表引用
            parts = v.split(":")
            if len(parts) == 1:
                nv = _map_endpoint(parts[0], col_map, row_map)
            elif len(parts) == 2:
                a = _REF_RE.match(parts[0])
                b = _REF_RE.match(parts[1])
                if not a or not b:                   # 整列/整行范围(G:G / 1:3)
                    return None
                nv1 = _map_endpoint(parts[0], col_map, row_map)
                nv2 = _map_endpoint(parts[1], col_map, row_map)
                if nv1 and nv2:
                    m1, m2 = _REF_RE.match(nv1), _REF_RE.match(nv2)
                    if int(m2.group(4)) - int(m1.group(4)) != int(b.group(4)) - int(a.group(4)):
                        return None                  # 区间中间有跳行 → 失真, 转值
                    nv = nv1 + ":" + nv2
                else:
                    nv = None
            else:
                return None
            if nv is None:
                return None
            out.append(nv)
        elif t.type == "OPERAND" and t.subtype == "NAME":
            if v.upper() in ("TRUE", "FALSE"):
                out.append(v)
            else:
                return None                          # 命名范围等 → 转值
        else:
            out.append(v)
    return "".join(out)


def translate_range_str(ref, col_map, row_map):
    """'B6:B21' 风格区间平移；不可平移返回 None。"""
    parts = ref.split(":")
    a = _map_endpoint(parts[0], col_map, row_map)
    b = _map_endpoint(parts[1], col_map, row_map) if len(parts) > 1 else a
    if a is None or b is None:
        return None
    if len(parts) > 1:
        sa, sb = _REF_RE.match(parts[0]), _REF_RE.match(parts[1])
        na, nb = _REF_RE.match(a), _REF_RE.match(b)
        if int(nb.group(4)) - int(na.group(4)) != int(sb.group(4)) - int(sa.group(4)):
            return None
    return a if len(parts) == 1 else f"{a}:{b}"


# ================================================================ 保真内容源

class FidelityContent:
    """xlsx 整表载入（公式版 + 缓存值版双开）。xls/csv 降级为纯值（报告注明）。"""

    def __init__(self, src):
        self.src = src
        self.kind = Path(src.path).suffix.lower()
        self.full = self.kind == ".xlsx"
        self.wb_f = self.wb_v = None
        if self.full:
            try:
                self.wb_f = load_workbook(src.path, data_only=False)
                self.wb_v = load_workbook(src.path, data_only=True)
            except Exception:
                self.wb_f = None

    @property
    def available(self):
        return (not self.full) or (self.wb_f is not None)

    def sheet(self):
        return self.wb_f[self.src.sheet_name] if self.wb_f else None

    def sheet_v(self):
        return self.wb_v[self.src.sheet_name] if self.wb_v else None

    def close(self):
        for wb in (self.wb_f, self.wb_v):
            try:
                wb.close()
            except Exception:
                pass


# ================================================================ 保真输出簿

class FidelityBook:
    def __init__(self, path):
        self.path = Path(path)
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self._used = set()
        self._cfg_ws = None
        self._cfg_col = 0
        self._cfg_registry = {}
        self.big_numbers_converted = 0

    def page(self, name):
        return self.wb.create_sheet(sanitize_sheet_name(name, self._used))

    def config_sheet(self):
        if self._cfg_ws is None:
            self._cfg_ws = self.wb.create_sheet("_下拉配置")
            self._cfg_ws.sheet_state = "hidden"
        return self._cfg_ws

    def copy_list_source(self, wb_f, sheet_title, coord):
        """把源选项区域复制到隐藏配置页，返回 '_下拉配置'!$X$1:$X$n 引用。"""
        key = (str(wb_f), sheet_title, coord)
        if key in self._cfg_registry:
            return self._cfg_registry[key]
        try:
            vals = [[c.value for c in row]
                    for row in wb_f[sheet_title][coord]]
        except Exception:
            return None
        flat = [v for row in vals for v in row if v not in (None, "")]
        if not flat:
            return None
        ws = self.config_sheet()
        self._cfg_col += 1
        col = get_column_letter(self._cfg_col)
        for i, v in enumerate(flat, start=1):
            ws.cell(row=i, column=self._cfg_col, value=v)
        ref = f"'_下拉配置'!${col}$1:${col}${len(flat)}"
        self._cfg_registry[key] = ref
        return ref

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(self.path)


# ================================================================ 组合并(保真)

def _extract_meta(pre_rows):
    """从前置区行提取 标签：值 对（"制表人：" 后跟值格, 或 "标签：值" 同格）。"""
    out = {}
    for row in pre_rows:
        for i, v in enumerate(row):
            if not isinstance(v, str):
                continue
            m = re.match(r"^([^：:]{1,12})[：:]\s*$", v.strip())
            if m:
                val = next((x for x in row[i + 1:] if x not in (None, "")), None)
                if val is not None:
                    out[m.group(1)] = str(val)
                continue
            m2 = re.match(r"^(.{2,10})[：:](\S.{0,28})$", v.strip())
            if m2 and "：" not in m2.group(2) and ":" not in m2.group(2):
                out[m2.group(1)] = m2.group(2)
    return out


def merge_group_fidelity(book, group, options, template_index, report,
                         progress=None, cancel=None, base_name="总表",
                         shape="stack"):
    totals = report["totals"]
    template = group[template_index]
    if options.mode == "workbook":
        out_keys = list(group[0].structure.keys)
        maps = [[i for i in range(len(out_keys))]]
        src_list = [group[0]]
    else:
        out_keys, maps = align_sources(group, options.align, template_index)
        src_list = group
    st = template.structure
    tkeys = st.keys
    tpos = {k: i for i, k in enumerate(tkeys)}
    contents = [FidelityContent(s) for s in src_list]
    t_content = contents[template_index if options.mode != "workbook" else 0]

    extra_texts = [k for k in out_keys if k not in tpos]

    # 汇总式前置区逐行共性检测: 前缀全同行保留一份; 差异区的 标签:值 → 元数据列
    meta_labels, meta_vals = [], []
    common_pre_n = 0
    if shape == "stack" and len(src_list) > 1:
        pres = [s.structure.pre_rows for s in src_list]
        for i in range(min(len(p) for p in pres)):
            row = pres[0][i]
            is_blank = not any(v not in (None, "") for v in row)
            if is_blank:
                continue                     # 空行跳过, 不断链
            if all(p[i] == row for p in pres):
                common_pre_n = i + 1         # 共同非空行 → 保留
            else:
                break                        # 内容不同 → 之后全部进差异区
        dicts = [_extract_meta(p[common_pre_n:]) for p in pres]
        cand = {}
        for d in dicts:
            for k, v in d.items():
                cand.setdefault(k, []).append(v)
        meta_labels = [k for k, vals in cand.items()
                       if len(vals) >= 2 and len(set(map(str, vals))) > 1]
        meta_vals = dicts
    meta_col0 = len(tkeys) + len(extra_texts)
    extra_texts = extra_texts + meta_labels
    source_label = "来源" if options.source_column else None
    title_row = st.title if (options.title_as_header and st.title) else None
    base_height = st.height or 1
    height = base_height + (1 if title_row else 0)
    total_width = len(tkeys) + len(extra_texts) + (1 if source_label else 0)

    serial_idx = next((i for i, k in enumerate(out_keys) if k in ("序号", "编号")), None) \
        if options.renumber_serial else None
    serial_counter = 0
    rows_total = 0

    # ---------------- 分页页面（表头/标题/合并/冻结/打印/列宽 每页重写） ----------------
    ws = None
    page_no = 0

    def _write_pre(ws, start_row, cont, sst_src, limit=None):
        """写出前置区(值+样式+合并平移), limit=只写前N行(汇总式共性前缀), 返回占用行数。"""
        if sst_src.header_row <= 1:
            return 0
        n_rows = limit if limit is not None else sst_src.header_row - 1
        if not cont.full:
            for r_i, prow in enumerate(sst_src.pre_rows[:n_rows]):
                for c_i, v in enumerate(prow, start=1):
                    if v is not None:
                        ws.cell(row=start_row + r_i, column=c_i, value=v)
            return min(n_rows, len(sst_src.pre_rows))
        rows = [[c for c in r] for r in cont.sheet().iter_rows(
            min_row=1, max_row=sst_src.header_row - 1)][:n_rows]
        for r_idx, r in enumerate(rows, start=1):
            for c_idx, cell in enumerate(r, start=1):
                out = ws.cell(row=start_row + r_idx - 1, column=c_idx)
                v = cell.value
                if v is not None:
                    out.value = v
                if cell.has_style:
                    _copy_cell_style(out, cell)
                if cell.comment:
                    out.comment = Comment(cell.comment.text, cell.comment.author)
        for mg in getattr(sst_src, "all_merges", []):
            if mg.max_row < sst_src.header_row and mg.max_row <= n_rows:
                try:
                    ws.merge_cells(start_column=mg.min_col,
                                   start_row=mg.min_row + start_row - 1,
                                   end_column=mg.max_col,
                                   end_row=mg.max_row + start_row - 1)
                except Exception:
                    pass
        return len(rows)

    def _write_header_at(ws, at_row, src_i):
        """把 src_i 自己的表头块写到 at_row（拼接式逐源表头）。"""
        sst_i = src_i.structure
        hcells = src_i.header_cells()
        for r_idx, row in enumerate(hcells, start=1):
            for c_idx, cell in enumerate(row, start=1):
                out = ws.cell(row=at_row + r_idx - 1, column=c_idx)
                v = getattr(cell, "value", None)
                if v is not None:
                    out.value = v
                if getattr(cell, "has_style", False):
                    _copy_cell_style(out, cell)
                cm = getattr(cell, "comment", None)
                if cm:
                    out.comment = Comment(cm.text, cm.author)
            for j, text in enumerate(extra_texts):
                if r_idx == 1:
                    ws.cell(row=at_row + r_idx - 1, column=len(tkeys) + 1 + j,
                            value=text)
        shift = sst_i.header_row - 1
        for mg in sst_i.merges:
            try:
                ws.merge_cells(start_column=mg.min_col, start_row=mg.min_row - shift + at_row - 1,
                               end_column=mg.max_col,
                               end_row=mg.max_row - shift + at_row - 1)
            except Exception:
                pass

    def _new_page(write_pre_from=None):
        nonlocal ws, page_no
        name = base_name if page_no == 0 else f"{base_name}({page_no + 1})"
        ws = book.page(name)
        page_no += 1
        r0 = 0
        if write_pre_from is not None and st.header_row > 1:
            r0 = _write_pre(ws, 1, write_pre_from, st, limit=pre_limit)
        elif title_row:
            ws.cell(row=1, column=1, value=title_row)
            if total_width > 1:
                ws.merge_cells(start_column=1, start_row=1,
                               end_column=total_width, end_row=1)
            r0 = 1
        for r_idx, row in enumerate(template.header_cells(), start=1):
            rr = r_idx + r0
            for c_idx, cell in enumerate(row, start=1):
                out = ws.cell(row=rr, column=c_idx)
                v = getattr(cell, "value", None)
                if v is not None:
                    out.value = v
                if getattr(cell, "has_style", False):
                    _copy_cell_style(out, cell)
                cm = getattr(cell, "comment", None)
                if cm:
                    out.comment = Comment(cm.text, cm.author)
            for j, text in enumerate(extra_texts):
                if r_idx == 1:
                    ws.cell(row=rr, column=len(tkeys) + 1 + j, value=text)
        if source_label:
            ws.cell(row=1 + r0, column=total_width, value=source_label)
        shift = st.header_row - 1
        for mg in st.merges:
            try:
                ws.merge_cells(start_column=mg.min_col, start_row=mg.min_row - shift + r0,
                               end_column=mg.max_col,
                               end_row=min(mg.max_row - shift + r0, r0 + base_height))
            except Exception:
                pass
        if base_height > 1:
            for j in range(len(extra_texts) + (1 if source_label else 0)):
                c = len(tkeys) + 1 + j
                try:
                    ws.merge_cells(start_column=c, start_row=1 + r0,
                                   end_column=c, end_row=r0 + base_height)
                except Exception:
                    pass
        # 冻结窗格：照模板(行号按块位移)，无则表头+1
        freeze = t_content.sheet().freeze_panes if t_content.full else None
        if freeze and str(freeze) != "A1":
            m = re.match(r"^([A-Z]+)(\d+)$", str(freeze))
            if m:
                row_shift = (1 + r0 + base_height) - (st.header_row + st.height)
                ws.freeze_panes = f"{m.group(1)}{max(1, int(m.group(2)) + row_shift)}"
        else:
            ws.freeze_panes = f"A{height + 1}"
        # 打印设置
        if t_content.full:
            try:
                ts = t_content.sheet()
                for attr in ("orientation", "paperSize", "scale", "fitToWidth",
                             "fitToHeight"):
                    if getattr(ts.page_setup, attr, None) is not None:
                        setattr(ws.page_setup, attr, getattr(ts.page_setup, attr))
                for attr in ("left", "right", "top", "bottom", "header", "footer"):
                    if getattr(ts.page_margins, attr, None) is not None:
                        setattr(ws.page_margins, attr, getattr(ts.page_margins, attr))
            except Exception:
                pass
        for c, w in (st.col_widths or {}).items():
            ws.column_dimensions[get_column_letter(c)].width = w
        return ws

    # 首页前置区: 拼接式=模板全量; 汇总式=逐行共性前缀(全同行保留一份)
    if shape == "append":
        first_pre, pre_limit = t_content, None
    elif common_pre_n > 0:
        first_pre, pre_limit = t_content, common_pre_n
    else:
        first_pre, pre_limit = None, 0
    ws = _new_page(write_pre_from=first_pre)
    block_base = ws.max_row                     # 前置区(如有)+表头 已占的行数
    hdr_start = block_base - (st.height or 1) + 1

    # ---------------- 数据行（两遍：先完整行映射，公式引用才可翻译） ----------------
    for fi, (src, mapping) in enumerate(zip(src_list, maps)):
        cont = contents[src_list.index(src)]
        sst = src.structure
        dstart = sst.header_row + sst.height
        inv_map = {si: oi for oi, si in enumerate(mapping) if si is not None}
        col_map = {sc + 1: oc + 1 for sc, oc in inv_map.items()}

        def _record_cells():
            """产出 [(源行号, cells_f, row_v)]; 横向表自动转置。"""
            if not cont.full:
                return [(row_no, cells, None)
                        for row_no, cells in src.iter_data_rows()]
            sheet_f = cont.sheet()
            sheet_v = cont.sheet_v()
            if sst.orientation == "horizontal":
                grid_f = [list(r) for r in sheet_f.iter_rows()]
                grid_v = [list(r) for r in sheet_v.iter_rows()] if sheet_v else []
                ncols = max((len(r) for r in grid_f), default=0)
                recs = []
                for c in range(dstart - 1, ncols):
                    cells_f = [(grid_f[r][c] if c < len(grid_f[r]) else None)
                               for r in range(len(grid_f))]
                    row_v = ([(grid_v[r][c] if r < len(grid_v) and c < len(grid_v[r])
                               else None) for r in range(len(grid_v))]
                             if grid_v else None)
                    recs.append((c + 1, cells_f, row_v))
                return recs
            rows_v_iter = sheet_v.iter_rows(min_row=dstart) if sheet_v else None
            recs = []
            for row_f in sheet_f.iter_rows(min_row=dstart):
                row_v = next(rows_v_iter, None) if rows_v_iter else None
                if row_f:
                    recs.append((row_f[0].row, list(row_f), row_v))
            return recs

        def _vals(cells_f, row_v):
            out = []
            for i, c in enumerate(cells_f):
                if c is None:
                    out.append(None)
                elif row_v is not None and i < len(row_v) and row_v[i] is not None:
                    out.append(row_v[i].value)
                else:
                    v = getattr(c, "value", c)
                    out.append(None if isinstance(v, str) and v.startswith("=") else v)
            return out

        keep = []
        for row_no, cells_f, row_v in _record_cells():
            if _skip_row(_vals(cells_f, row_v), options, totals):
                continue
            keep.append((row_no, cells_f, row_v))
        if cancel and cancel():
            raise KeyboardInterrupt("cancelled")
        row_map = {}
        if shape == "append" and fi > 0:
            # 拼接式: 空行 + 该源前置区 + 该源表头 成块
            at = ws.max_row + 2
            if sst.header_row > 1:
                at += _write_pre(ws, at, cont, sst)
            _write_header_at(ws, at, src)
            hdr_start = at
        cursor = (ws.max_row if shape == "append" else block_base) + rows_total
        for row_no, _f, _v in keep:
            if ws.max_row >= XLSX_MAX_ROWS - 2:
                _new_page()
                report["warnings"].append(
                    f"「{base_name}」超出单页行数上限，已自动续写新页")
                block_base = ws.max_row
                cursor = block_base
            cursor += 1
            row_map[row_no] = (page_no, cursor)

        flat_map = {r: pos for r, (_p, pos) in row_map.items()}
        data_rows = 0
        for row_no, cells_f, row_v in keep:
            pg, out_row = row_map[row_no]
            if pg != page_no:
                continue
            if serial_idx is not None:
                serial_counter += 1
            for oi, si in enumerate(mapping):
                out = ws.cell(row=out_row, column=oi + 1)
                if si is None:
                    continue
                cf = cells_f[si] if si < len(cells_f) else None
                if cf is None:
                    continue
                if cont.full:
                    v = cf.value
                    if isinstance(v, str) and v.startswith("="):
                        cached = (row_v[si].value
                                  if row_v is not None and si < len(row_v)
                                  and row_v[si] is not None else None)
                        newf = translate_formula(v, col_map, flat_map)
                        if newf is not None:
                            out.value = newf if newf.startswith("=") else "=" + newf
                            totals["formulas_translated"] = \
                                totals.get("formulas_translated", 0) + 1
                        else:
                            out.value = cached
                            totals["formulas_to_value"] = \
                                totals.get("formulas_to_value", 0) + 1
                    else:
                        v, conv = big_number_guard(v)
                        if conv:
                            book.big_numbers_converted += 1
                        out.value = v
                    if getattr(cf, "has_style", False):
                        _copy_cell_style(out, cf)
                    cm = getattr(cf, "comment", None)
                    if cm:
                        out.comment = Comment(cm.text, cm.author)
                else:
                    v = getattr(cf, "value", cf)
                    v, conv = big_number_guard(v)
                    if conv:
                        book.big_numbers_converted += 1
                    out.value = v
            if serial_idx is not None:
                ws.cell(row=out_row, column=serial_idx + 1).value = serial_counter
            if source_label:
                ws.cell(row=out_row, column=total_width).value = src.display
            for j, label in enumerate(meta_labels):
                ws.cell(row=out_row, column=meta_col0 + 1 + j,
                        value=(meta_vals[fi] or {}).get(label))
            totals["rows_written"] = totals.get("rows_written", 0) + 1
            data_rows += 1

        rows_total += data_rows
        if data_rows == 0:
            report["warnings"].append(f"「{src.display}」仅有表头无数据，已跳过")
            totals["empty_sources"] += 1
        totals["sheets_merged"] += 1 if data_rows else 0

        # 数据区合并还原（同页且行连续）
        if data_rows:
            kept = 0
            for mg in getattr(sst, "all_merges", []):
                if mg.min_row < dstart:
                    continue
                a, b = row_map.get(mg.min_row), row_map.get(mg.max_row)
                if a is None or b is None or a[0] != b[0]:
                    continue
                r1, r2 = a[1], b[1]
                if r2 - r1 != mg.max_row - mg.min_row:
                    continue
                c1, c2 = inv_map.get(mg.min_col - 1), inv_map.get(mg.max_col - 1)
                if c1 is None or c2 is None or c2 - c1 != mg.max_col - mg.min_col:
                    continue
                try:
                    ws.merge_cells(start_column=c1 + 1, start_row=r1,
                                   end_column=c2 + 1, end_row=r2)
                    kept += 1
                except Exception:
                    pass
            if kept:
                totals["data_merges_preserved"] = \
                    totals.get("data_merges_preserved", 0) + kept

        # 行高
        if cont.full and data_rows:
            sheet_f = cont.sheet()
            for sr, (_pg, orr) in row_map.items():
                try:
                    h = sheet_f.row_dimensions[sr].height
                except Exception:
                    h = None
                if h:
                    ws.row_dimensions[orr].height = h
                    totals["row_heights_set"] = totals.get("row_heights_set", 0) + 1

        # 数据验证 / 条件格式 / 智能表格
        if cont.full and data_rows:
            _copy_validations(book, ws, cont, col_map, flat_map, dstart, totals)
            _copy_cond_formats(ws, cont, col_map, flat_map, dstart, totals)
            _copy_tables(book, ws, cont, col_map, flat_map, len(src_list), totals,
                         hdr_start=hdr_start, sst=sst)

        if progress:
            progress(fi + 1, len(src_list), src.display, rows_total)

    for c in contents:
        c.close()
    return ws


def _skip_row(vals, options, totals):
    if all(v in (None, "") for v in vals):
        totals["empty_rows_skipped"] += 1
        return True
    if options.remove_summary and is_summary_row(vals, options.summary_keywords):
        totals["summary_rows_removed"] += 1
        return True
    if options.strip_serial_rows:
        ne = [(i, v) for i, v in enumerate(vals) if v not in (None, "")]
        if len(ne) == 1 and ne[0][0] == 0 and is_numberlike(ne[0][1]):
            totals["serial_rows_removed"] = totals.get("serial_rows_removed", 0) + 1
            return True
    return False


def _write_row_fidelity(ws, out_row, cells_f, row_v, mapping, inv_map, row_map,
                        row_no, serial_idx, options, book, totals, report, src,
                        serial_value=None):
    col_map = {sc + 1: oc + 1 for sc, oc in inv_map.items()}
    for oi, si in enumerate(mapping):
        out = ws.cell(row=out_row, column=oi + 1)
        if si is None:
            continue
        cell_f = cells_f[si] if si < len(cells_f) else None
        if cell_f is None:
            continue
        v = cell_f.value
        if isinstance(v, str) and v.startswith("="):
            cached = row_v[si].value if row_v is not None and si < len(row_v) else None
            newf = translate_formula(v, col_map, row_map)
            if newf is not None:
                out.value = newf if newf.startswith("=") else "=" + newf
                totals["formulas_translated"] = totals.get("formulas_translated", 0) + 1
            else:
                out.value = cached
                totals["formulas_to_value"] = totals.get("formulas_to_value", 0) + 1
        else:
            v, conv = big_number_guard(v)
            if conv:
                book.big_numbers_converted += 1
            out.value = v
        if cell_f.has_style:
            _copy_cell_style(out, cell_f)
        if cell_f.comment:
            out.comment = Comment(cell_f.comment.text, cell_f.comment.author)
    if serial_value is not None and serial_idx is not None:
        cur = ws.cell(row=out_row, column=serial_idx + 1)
        cur.value = serial_value
    totals["rows_written"] = totals.get("rows_written", 0) + 1


def _write_plain_row(ws, out_row, out_cells, serial_idx):
    for col, v in out_cells:
        ws.cell(row=out_row, column=col, value=v)
    if serial_idx is not None:
        ws.cell(row=out_row, column=serial_idx + 1, value=out_row - 1)


# --------------------------------------------------------------- 数据验证

def _copy_validations(book, ws, cont, col_map, row_map, dstart, totals):
    sheet = cont.sheet()
    cfg_wb = cont.wb_f
    for dv in sheet.data_validations.dataValidation or []:
        new_sqref = []
        for cr in dv.sqref.ranges:
            c1, c2 = col_map.get(cr.min_col), col_map.get(cr.max_col)
            r1, r2 = row_map.get(cr.min_row), row_map.get(cr.max_row)
            if None in (c1, c2, r1, r2) or (r2 - r1) != (cr.max_row - cr.min_row) \
                    or (c2 - c1) != (cr.max_col - cr.min_col):
                continue
            new_sqref.append(CellRange(min_col=c1, min_row=r1, max_col=c2, max_row=r2))
        if not new_sqref:
            totals["validations_dropped"] = totals.get("validations_dropped", 0) + 1
            continue
        f1 = dv.formula1
        if dv.type == "list" and f1 and not f1.startswith('"'):
            m = _SHEET_REF_RE.match(str(f1))
            if m:
                sheet_title = m.group(1) or m.group(2)
                coord = m.group(3).replace("$", "")
                new_ref = book.copy_list_source(cfg_wb, sheet_title, coord)
            else:
                new_ref = book.copy_list_source(cfg_wb, cont.src.sheet_name,
                                                str(f1).replace("$", ""))
            if new_ref is None:
                totals["validations_dropped"] = totals.get("validations_dropped", 0) + 1
                continue
            f1 = new_ref
        ndv = DataValidation(type=dv.type, formula1=f1, formula2=dv.formula2,
                             allow_blank=dv.allowBlank, showDropDown=dv.showDropDown,
                             operator=dv.operator, error=dv.error,
                             errorTitle=dv.errorTitle, prompt=dv.prompt,
                             promptTitle=dv.promptTitle, showErrorMessage=dv.showErrorMessage,
                             showInputMessage=dv.showInputMessage)
        for cr in new_sqref:
            ndv.add(cr)
        ws.add_data_validation(ndv)
        totals["validations_kept"] = totals.get("validations_kept", 0) + 1


# --------------------------------------------------------------- 条件格式

def _copy_cond_formats(ws, cont, col_map, row_map, dstart, totals):
    sheet = cont.sheet()
    try:
        cf_all = list(sheet.conditional_formatting)
    except Exception:
        return
    for cf in cf_all:
        for rule in cf.rules:
            if rule.formula:
                nfs = []
                ok = True
                for f in rule.formula:
                    nf = translate_formula("=" + f, col_map, row_map)
                    if nf is None:
                        ok = False
                        break
                    nfs.append(nf[1:])
                if not ok:
                    totals["cf_dropped"] = totals.get("cf_dropped", 0) + 1
                    continue
                rule = deepcopy(rule)
                rule.formula = nfs
            else:
                rule = deepcopy(rule)
            for cr in cf.sqref.ranges:
                c1, c2 = col_map.get(cr.min_col), col_map.get(cr.max_col)
                r1, r2 = row_map.get(cr.min_row), row_map.get(cr.max_row)
                if None in (c1, c2, r1, r2) or (r2 - r1) != (cr.max_row - cr.min_row):
                    continue
                ws.conditional_formatting.add(
                    f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}", rule)
                totals["cf_kept"] = totals.get("cf_kept", 0) + 1


# --------------------------------------------------------------- 智能表格

def _copy_tables(book, ws, cont, col_map, row_map, group_size, totals,
                 hdr_start=None, sst=None):
    sheet = cont.sheet()
    dstart = sst.header_row + sst.height if sst else 10 ** 9

    def _tr(r):
        if r < dstart and hdr_start is not None:
            return hdr_start + (r - sst.header_row)     # 表头区行 → 输出表头坐标
        return row_map.get(r)

    for tbl in list(getattr(sheet, "tables", {}).values()):
        name = getattr(tbl, "displayName", "表")
        if group_size > 1:
            totals["tables_downgraded"] = totals.get("tables_downgraded", 0) + 1
            continue
        m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", str(tbl.ref))
        if not m:
            totals["tables_downgraded"] = totals.get("tables_downgraded", 0) + 1
            continue
        c1 = col_map.get(column_index_from_string(m.group(1)))
        c2 = col_map.get(column_index_from_string(m.group(3)))
        r1, r2 = _tr(int(m.group(2))), _tr(int(m.group(4)))
        if None in (c1, c2, r1, r2):
            totals["tables_downgraded"] = totals.get("tables_downgraded", 0) + 1
            continue
        try:
            nt = Table(displayName=f"{name}_{len(book.wb.sheetnames)}",
                       ref=f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}")
            if tbl.tableStyleInfo:
                nt.tableStyleInfo = deepcopy(tbl.tableStyleInfo)
            ws.add_table(nt)
            totals["tables_kept"] = totals.get("tables_kept", 0) + 1
        except Exception:
            totals["tables_downgraded"] = totals.get("tables_downgraded", 0) + 1
