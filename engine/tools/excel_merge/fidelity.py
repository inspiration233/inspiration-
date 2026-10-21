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
from openpyxl.formula.translate import Translator
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
        # WPS兼容: 跳过 patternType=None 的空 PatternFill(会产生无效的<fill/>元素);
        # GradientFill 无 patternType 属性但本身有效, 需保留
        if hasattr(fill, "stop") or getattr(fill, "patternType", None):
            dst.fill = _copy(fill)
        dst.border = _copy(_real(src.border))
        dst.alignment = _copy(_real(src.alignment))
        fmt = src.number_format
        if fmt and fmt != "General":
            dst.number_format = fmt
    except Exception:
        pass


def _copy_row_dimension(ws, index, source):
    """Register dimension styles in the destination, never copy foreign IDs.

    RowDimension's StyleArray contains workbook-local font/fill/format indices.
    Copying it directly can produce an otherwise well-formed, unreadable XLSX.
    """
    target = _copy(source)
    target.index, target.parent = index, ws
    target._style = None
    if source.has_style:
        for attribute in ("font", "fill", "border", "alignment", "protection"):
            setattr(target, attribute, _copy(_real(getattr(source, attribute))))
        target.number_format = source.number_format
        target.quotePrefix = source.quotePrefix
        target.pivotButton = source.pivotButton
    ws.row_dimensions[index] = target

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
    if any(t.type == "FUNC" and t.subtype == "OPEN" and
                     t.value.rstrip("(").upper().removeprefix("_XLFN.") in
                     {"INDIRECT", "OFFSET", "CELL", "ADDRESS"} for t in tokens):
        # Addresses held in strings or calculated at runtime cannot follow a
        # structural move by translating RANGE tokens alone. Even unchanged
        # coordinates may depend on cells or sheets omitted from the output.
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
                    source_columns = range(column_index_from_string(a.group(2)),
                                           column_index_from_string(b.group(2)) + 1)
                    mapped_columns = [col_map.get(c) for c in source_columns]
                    if None in mapped_columns or mapped_columns != list(range(
                            column_index_from_string(m1.group(2)),
                            column_index_from_string(m2.group(2)) + 1)):
                        return None                  # 列换序不能改变范围参与计算的集合
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
            if coord in wb_f.defined_names:
                destinations = list(wb_f.defined_names[coord].destinations)
                if len(destinations) != 1:
                    return None
                sheet_title, coord = destinations[0]
            cells = wb_f[sheet_title][coord]
            if not isinstance(cells, tuple):
                cells = ((cells,),)
            if any(c.data_type == "f" for row in cells for c in row):
                # Copying formula text into a new hidden sheet would change its
                # dependencies. Let the caller report a dropped validation.
                return None
            vals = [[c.value for c in row] for row in cells]
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
    src_list = [group[0]] if options.mode == "workbook" else group
    contents = [FidelityContent(s) for s in src_list]
    unavailable = [s.display for s, cont in zip(src_list, contents) if not cont.available]
    if unavailable:
        for opened in contents:
            opened.close()
        raise ValueError("无法读取原生工作表内容：" + "、".join(unavailable))
    t_content = contents[template_index if options.mode != "workbook" else 0]

    def use_source_blocks(reason):
        nonlocal shape
        shape = "append"
        report["warnings"].append(reason + "；已在同一工作表按来源保留完整表块。")
        for item in report.get("groups", []):
            if item.get("sheet") == base_name:
                item["shape"] = "append"

    if (shape != "append" and len(src_list) > 1
            and getattr(options, "preserve_template_preamble", False)):
        pres = [s.structure.pre_rows for s in src_list]
        # Shared formula text does not imply shared meaning: a source's summary
        # or header formula can refer to that source's own records.
        source_header_formula = any(
            cell.data_type == "f"
            for source, cont in zip(src_list, contents) if cont.full
            for cell in cont.sheet()._cells.values()
            if (cell.column if source.structure.orientation == "horizontal" else cell.row)
            < source.structure.header_row + source.structure.height
            and getattr(source.structure, "field_offset", 0)
            < (cell.row if source.structure.orientation == "horizontal" else cell.column)
            <= getattr(source.structure, "field_end",
                       getattr(source.structure, "field_offset", 0) + len(source.structure.keys)))
        if any(p != pres[0] for p in pres[1:]):
            use_source_blocks("各来源的标题或前置信息不同，不能用第一张表的抬头代表全部记录")
        elif source_header_formula:
            use_source_blocks("前置区或表头含各来源自己的公式，需要分别保留计算依据")

    if options.mode == "workbook":
        out_keys = list(group[0].structure.keys)
        maps = [[i for i in range(len(out_keys))]]
    elif shape == "append":
        width = max(len(s.structure.keys) for s in src_list)
        out_keys = list(template.structure.keys) + [""] * (width - len(template.structure.keys))
        maps = [[i if i < len(s.structure.keys) else None for i in range(width)] for s in src_list]
    else:
        try:
            out_keys, maps = align_sources(group, options.align, template_index,
                                          allow_leaf=not getattr(options, "exact_headers", False))
        except BaseException:
            for opened in contents:
                opened.close()
            raise
        if (getattr(options, "preserve_template_preamble", False)
                and any(value not in (None, "") for s in src_list
                        for row in s.structure.pre_rows for value in row)
                and any(si is not None and si != oi for mapping in maps for oi, si in enumerate(mapping))):
            use_source_blocks("字段顺序不同且含前置信息，需保持各块原列位置以保留其引用关系")
            width = max(len(s.structure.keys) for s in src_list)
            out_keys = list(template.structure.keys) + [""] * (width - len(template.structure.keys))
            maps = [[i if i < len(s.structure.keys) else None for i in range(width)] for s in src_list]
    if shape == "append" and any(getattr(s, "field_pairs", {}) or getattr(s, "unmapped_targets", set()) for s in src_list):
        report["warnings"].append("完整表块按各自原列顺序保留；手工字段对应仅用于单表头汇总，本次未改变各块列位置。")
    st = template.structure
    tkeys = st.keys
    tpos = {k: i for i, k in enumerate(tkeys)}
    for src, cont in zip(src_list, contents):
        if src.structure.orientation != "horizontal" or not cont.full:
            continue
        original = cont.sheet()
        if (any(c.data_type == "f" for c in original._cells.values())
                or original.data_validations.dataValidation or original.conditional_formatting
                or original.tables or original.defined_names or cont.wb_f.defined_names):
            for opened in contents:
                opened.close()
            raise ValueError(
                f"「{src.display}」横向记录包含公式、验证、条件格式或命名/表格依赖，"
                "当前拼接路径不能可靠转置这些引用。请使用各表独立成页以保留原布局。")

    # A deliberately unmapped source field may have the same visible label as
    # a template column. Its data still needs a distinct extra output column.
    extra_texts = [] if shape == "append" else list(out_keys[len(tkeys):])
    if shape == "append" and any(s.structure.col_widths != st.col_widths for s in src_list):
        report["warnings"].append("各表块列宽不同；同一工作表的列宽只能共用，已采用版式参考表列宽，各块单元格样式保留。")

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
        if not getattr(options, "metadata_columns", True):
            meta_labels = []
        meta_vals = dicts
    meta_col0 = len(tkeys) + len(extra_texts)
    extra_texts = extra_texts + meta_labels
    source_label = "来源" if options.source_column else None
    title_row = st.title if (options.title_as_header and st.title) else None
    base_height = st.height or 1
    height = base_height + (1 if title_row else 0)
    total_width = (len(out_keys) if shape == "append" else len(tkeys) + len(extra_texts)) + (1 if source_label else 0)

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
        off = getattr(sst_src, "field_offset", 0)
        if sst_src.orientation == "horizontal":
            rows = [[cont.sheet().cell(r, c) for r in range(off + 1, getattr(sst_src, "field_end", cont.sheet().max_row) + 1)]
                    for c in range(1, sst_src.header_row)][:n_rows]
        else:
            rows = [[c for c in r] for r in cont.sheet().iter_rows(
                min_row=1, max_row=sst_src.header_row - 1, min_col=off + 1,
                max_col=getattr(sst_src, "field_end", cont.sheet().max_column))][:n_rows]
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

    def _new_page(write_pre_from=None, header_source=None):
        nonlocal ws, page_no
        page_source = header_source or template
        page_st = page_source.structure
        page_height = page_st.height or 1
        name = base_name if page_no == 0 else f"{base_name}({page_no + 1})"
        ws = book.page(name)
        page_no += 1
        r0 = 0
        if write_pre_from is not None and page_st.header_row > 1:
            r0 = _write_pre(ws, 1, write_pre_from, page_st, limit=pre_limit)
        elif title_row:
            ws.cell(row=1, column=1, value=title_row)
            if total_width > 1:
                ws.merge_cells(start_column=1, start_row=1,
                               end_column=total_width, end_row=1)
            r0 = 1
        for r_idx, row in enumerate(page_source.header_cells(), start=1):
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
        shift = page_st.header_row - 1
        for mg in page_st.merges:
            try:
                ws.merge_cells(start_column=mg.min_col, start_row=mg.min_row - shift + r0,
                               end_column=mg.max_col,
                               end_row=min(mg.max_row - shift + r0, r0 + page_height))
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
                row_shift = (1 + r0 + page_height) - (st.header_row + st.height)
                ws.freeze_panes = f"{m.group(1)}{max(1, int(m.group(2)) + row_shift)}"
        elif not getattr(options, "preserve_template_preamble", False):
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
                if getattr(options, "preserve_template_preamble", False):
                    for attr in ("page_setup", "page_margins", "print_options", "sheet_properties",
                                 "sheet_format", "sheet_view", "protection", "oddHeader", "oddFooter",
                                 "evenHeader", "evenFooter", "firstHeader", "firstFooter"):
                        if attr == "sheet_view":
                            # WPS兼容: 拷贝views但重置workbookViewId(新簿只有一个窗口)
                            ws.views = deepcopy(ts.views)
                            ws.views.sheetView = ws.views.sheetView[:1]
                            if ws.views.sheetView:
                                ws.views.sheetView[0].workbookViewId = 0
                            # A new output workbook has one window/view. Source
                            # window IDs are not valid references in this book.
                            ws.views.sheetView = ws.views.sheetView[:1]
                            if ws.views.sheetView:
                                ws.views.sheetView[0].workbookViewId = 0
                        else:
                            setattr(ws, attr, deepcopy(getattr(ts, attr)))
                    # Printer-specific binary parts/relationships are not copied.
                    ws.page_setup.id = None
                    if ts.print_title_rows:
                        ws.print_title_rows = ts.print_title_rows
                    if ts.print_title_cols:
                        ws.print_title_cols = ts.print_title_cols
                    for row_num in range(1, st.header_row + st.height):
                        if row_num in ts.row_dimensions:
                            _copy_row_dimension(ws, row_num, ts.row_dimensions[row_num])
            except Exception:
                pass
        for c, w in (st.col_widths or {}).items():
            ws.column_dimensions[get_column_letter(c)].width = w
        return ws

    # 首页前置区: 拼接式=模板全量; 汇总式=逐行共性前缀(全同行保留一份)
    initial_source = src_list[0] if shape == "append" else template
    if shape == "append" or getattr(options, "preserve_template_preamble", False):
        first_pre, pre_limit = (contents[0] if shape == "append" else t_content), None
    elif common_pre_n > 0:
        first_pre, pre_limit = t_content, common_pre_n
    else:
        first_pre, pre_limit = None, 0
    ws = _new_page(write_pre_from=first_pre, header_source=initial_source)
    block_base = ws.max_row                     # 前置区(如有)+表头 已占的行数
    hdr_start = block_base - (initial_source.structure.height or 1) + 1

    # ---------------- 数据行（两遍：先完整行映射，公式引用才可翻译） ----------------
    for fi, (src, mapping) in enumerate(zip(src_list, maps)):
        cont = contents[src_list.index(src)]
        sst = src.structure
        if shape == "append":
            serial_idx = next((i for i, key in enumerate(sst.keys) if key in ("序号", "编号")), None) \
                if options.renumber_serial else None
        dstart = getattr(sst, "data_start", sst.header_row + sst.height)
        inv_map = {si: oi for oi, si in enumerate(mapping) if si is not None}
        field_offset = getattr(sst, "field_offset", 0)
        col_map = {sc + field_offset + 1: oc + 1 for sc, oc in inv_map.items()}

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
                for c in range(dstart - 1, min(ncols, getattr(sst, "data_end", ncols))):
                    cells_f = [(grid_f[r][c] if c < len(grid_f[r]) else None)
                               for r in range(len(grid_f))]
                    row_v = ([(grid_v[r][c] if r < len(grid_v) and c < len(grid_v[r])
                               else None) for r in range(len(grid_v))]
                             if grid_v else None)
                    fend = getattr(sst, "field_end", len(cells_f))
                    recs.append((c + 1, cells_f[field_offset:fend],
                                 row_v[field_offset:fend] if row_v else None))
                return recs
            max_row = getattr(sst, "data_end", sheet_f.max_row)
            max_col = getattr(sst, "field_end", sheet_f.max_column)
            rows_v_iter = sheet_v.iter_rows(min_row=dstart, max_row=max_row,
                                           min_col=field_offset + 1, max_col=max_col) if sheet_v else None
            recs = []
            for row_f in sheet_f.iter_rows(min_row=dstart, max_row=max_row,
                                           min_col=field_offset + 1, max_col=max_col):
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

        # 第一遍: 全部收集行值(用于检测数据区域结束)
        all_rows = list(_record_cells())
        all_vals = [_vals(cf, rv) for _, cf, rv in all_rows]
        data_end = detect_data_end(all_vals)

        keep = []
        for idx, (row_no, cells_f, row_v) in enumerate(all_rows):
            if data_end is not None and idx >= data_end:
                break  # 数据区域已结束, 后续是序号占位尾部
            if _skip_row(_vals(cells_f, row_v), options, totals):
                continue
            keep.append((row_no, cells_f, row_v))
        if cancel and cancel():
            raise KeyboardInterrupt("cancelled")
        row_map = {}
        prefix_start = 1 if first_pre is not None else None
        prefix_count = (sst.header_row - 1 if pre_limit is None else pre_limit) if prefix_start else 0
        if shape == "append" and fi > 0:
            # 拼接式: 空行 + 该源前置区 + 该源表头 成块
            at = ws.max_row + 2
            prefix_start, prefix_count = at, sst.header_row - 1
            if sst.header_row > 1:
                at += _write_pre(ws, at, cont, sst)
            _write_header_at(ws, at, src)
            hdr_start = at
        cursor = ws.max_row if shape == "append" else block_base + rows_total
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
        # Header references belong to the output header, not the data-row map.
        for r in range(sst.header_row, sst.header_row + sst.height):
            flat_map[r] = hdr_start + r - sst.header_row
        for r in range(1, min(prefix_count, sst.header_row - 1) + 1):
            flat_map[r] = prefix_start + r - 1
        # Prefix/header formulas were initially copied with their native styles.
        # Rewrite them only once the complete source block mapping is available.
        # Every appended source owns its own prefix/header, even when a different
        # source was chosen as the visual template.
        if cont.full and (shape == "append" or src is template):
            for source_row in range(1, sst.header_row + sst.height):
                target_row = flat_map.get(source_row)
                if target_row is None:
                    continue
                if source_row in cont.sheet().row_dimensions:
                    _copy_row_dimension(ws, target_row, cont.sheet().row_dimensions[source_row])
                for source_col, target_col in col_map.items():
                    original = cont.sheet().cell(source_row, source_col)
                    if original.data_type != "f":
                        continue
                    converted = translate_formula(original.value, col_map, flat_map)
                    target_cell = ws.cell(target_row, target_col)
                    if converted is None:
                        target_cell.value = cont.sheet_v().cell(source_row, source_col).value if cont.sheet_v() else None
                        totals["formulas_to_value"] = totals.get("formulas_to_value", 0) + 1
                    else:
                        target_cell.value = converted if converted.startswith("=") else "=" + converted
                        totals["formulas_translated"] = totals.get("formulas_translated", 0) + 1
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
                affected_rows = [r for r in row_map if mg.min_row <= r <= mg.max_row]
                affected_cols = [c for c in inv_map if mg.min_col - 1 <= c <= mg.max_col - 1]
                if not affected_rows or not affected_cols:
                    continue
                def dropped_merge():
                    totals["data_merges_dropped"] = totals.get("data_merges_dropped", 0) + 1
                    report.setdefault("issues", []).append({
                        "id": f"data-merge:{src.display}:{mg.min_col}:{mg.min_row}",
                        "message": f"「{src.display}」数据合并区 "
                                   f"{get_column_letter(mg.min_col)}{mg.min_row}:"
                                   f"{get_column_letter(mg.max_col)}{mg.max_row} "
                                   "被选区或字段对应切断，不能保留且不能复制未知归属的记录。"})
                a, b = row_map.get(mg.min_row), row_map.get(mg.max_row)
                if a is None or b is None or a[0] != b[0]:
                    dropped_merge()
                    continue
                r1, r2 = a[1], b[1]
                if r2 - r1 != mg.max_row - mg.min_row:
                    dropped_merge()
                    continue
                c1, c2 = inv_map.get(mg.min_col - 1), inv_map.get(mg.max_col - 1)
                mapped_columns = [inv_map.get(c) for c in range(mg.min_col - 1, mg.max_col)]
                if (c1 is None or c2 is None or
                        mapped_columns != list(range(c1, c2 + 1))):
                    dropped_merge()
                    continue
                try:
                    ws.merge_cells(start_column=c1 + 1, start_row=r1,
                                   end_column=c2 + 1, end_row=r2)
                    kept += 1
                except Exception:
                    dropped_merge()
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
    if all(v in (None, "") for v in vals) and not getattr(options, "preserve_blank_rows", False):
        totals["empty_rows_skipped"] += 1
        return True
    if options.remove_summary and is_summary_row(vals, options.summary_keywords):
        totals["summary_rows_removed"] += 1
        return True
    if options.strip_serial_rows:
        ne = [(i, v) for i, v in enumerate(vals) if v not in (None, "")]
        # 恰好1个非空格 + 值像序号(数字形态,不限列位置) → 仅序号的空行
        if len(ne) == 1 and is_numberlike(ne[0][1]):
            totals["serial_rows_removed"] = totals.get("serial_rows_removed", 0) + 1
            return True
    return False


DATA_END_SPARSE_THRESHOLD = 10  # 连续N行稀疏→数据区域结束

def detect_data_end(rows_vals):
    """检测数据区域结束位置: 找到第一个连续≥N行稀疏(≤2非空格)的起点.

    返回数据结束的行索引(不含), None=全部是数据.
    政务表单的典型场景: r6~r13有实际数据, r14~r104只有序号占位.
    """
    sparse_run = 0
    for i, vals in enumerate(rows_vals):
        ne_count = sum(1 for v in vals if v not in (None, ""))
        if ne_count <= 2:
            sparse_run += 1
            if sparse_run >= DATA_END_SPARSE_THRESHOLD:
                return i - sparse_run + 1  # 数据结束于稀疏行起点
        else:
            sparse_run = 0
    return None


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
        ranges = sorted(dv.sqref.ranges, key=lambda r: (r.min_row, r.min_col))
        segments = [segment for cr in ranges for segment in _mapped_segments(cr, col_map, row_map)]
        if not segments:
            continue
        base = f"{get_column_letter(ranges[0].min_col)}{ranges[0].min_row}"
        f1 = dv.formula1
        list_reference = dv.type == "list" and f1 and not str(f1).startswith('"')
        if list_reference:
            source_formula = str(f1).lstrip("=")
            m = _SHEET_REF_RE.match(source_formula)
            if m:
                sheet_title = m.group(1) or m.group(2)
                coord = m.group(3).replace("$", "")
                new_ref = book.copy_list_source(cfg_wb, sheet_title, coord)
            else:
                new_ref = book.copy_list_source(cfg_wb, cont.src.sheet_name,
                                                source_formula.replace("$", ""))
            if new_ref is None:
                totals["validations_dropped"] = totals.get("validations_dropped", 0) + 1
                continue
            f1 = new_ref
        for source_range, target_range in segments:
            converted = [f1 if list_reference else _mapped_rule_formula(
                f1, base, source_range, target_range, col_map, row_map),
                _mapped_rule_formula(dv.formula2, base, source_range, target_range, col_map, row_map)]
            if ((f1 is not None and converted[0] is None)
                    or (dv.formula2 is not None and converted[1] is None)):
                totals["validations_dropped"] = totals.get("validations_dropped", 0) + 1
                continue
            ndv = deepcopy(dv)
            ndv.sqref = str(target_range)
            ndv.formula1, ndv.formula2 = converted
            ws.add_data_validation(ndv)
            totals["validations_kept"] = totals.get("validations_kept", 0) + 1


# --------------------------------------------------------------- 条件格式

def _mapped_segments(cr, col_map, row_map):
    """Source/target rectangles with constant offsets, including column reorder."""
    def spans(mapping, lo, hi):
        pairs = sorted((s, t) for s, t in mapping.items() if lo <= s <= hi)
        result = []
        for source, target in pairs:
            if result and source == result[-1][1] + 1 and target == result[-1][3] + 1:
                result[-1][1], result[-1][3] = source, target
            else:
                result.append([source, source, target, target])
        return result
    return [(CellRange(min_col=cs[0], max_col=cs[1], min_row=rs[0], max_row=rs[1]),
             CellRange(min_col=cs[2], max_col=cs[3], min_row=rs[2], max_row=rs[3]))
            for cs in spans(col_map, cr.min_col, cr.max_col)
            for rs in spans(row_map, cr.min_row, cr.max_row)]


def _mapped_cell_ranges(cr, col_map, row_map):
    return [mapped for _, mapped in _mapped_segments(cr, col_map, row_map)]


def _mapped_rule_formula(formula, base, source, target, col_map, row_map):
    """Rebase relative rules before structural mapping and verify the far corner."""
    if formula is None:
        return None
    text = "=" + str(formula).lstrip("=")
    try:
        first = Translator(text, origin=base).translate_formula(
            f"{get_column_letter(source.min_col)}{source.min_row}")
        last = Translator(text, origin=base).translate_formula(
            f"{get_column_letter(source.max_col)}{source.max_row}")
        mapped_first = translate_formula(first, col_map, row_map)
        mapped_last = translate_formula(last, col_map, row_map)
        if mapped_first is None or mapped_last is None:
            return None
        expected_last = Translator("=" + mapped_first,
                                  origin=f"{get_column_letter(target.min_col)}{target.min_row}").translate_formula(
            f"{get_column_letter(target.max_col)}{target.max_row}")
        if expected_last != "=" + mapped_last:
            return None
        return mapped_first
    except Exception:
        return None


def _copy_cond_formats(ws, cont, col_map, row_map, dstart, totals):
    sheet = cont.sheet()
    try:
        cf_all = list(sheet.conditional_formatting)
    except Exception:
        totals["cf_dropped"] = totals.get("cf_dropped", 0) + 1
        return
    for cf in cf_all:
        ranges = sorted(cf.sqref.ranges, key=lambda r: (r.min_row, r.min_col))
        segments = [segment for cr in ranges for segment in _mapped_segments(cr, col_map, row_map)]
        if not segments:
            continue
        base = f"{get_column_letter(ranges[0].min_col)}{ranges[0].min_row}"
        for rule in cf.rules:
            if rule.type in {"colorScale", "dataBar", "iconSet", "top10", "aboveAverage",
                             "duplicateValues", "uniqueValues"}:
                new = deepcopy(rule)
                safe = True
                for attr in ("colorScale", "dataBar", "iconSet"):
                    for threshold in getattr(getattr(new, attr, None), "cfvo", []) or []:
                        if threshold.type != "formula":
                            continue
                        text = "=" + str(threshold.val).lstrip("=")
                        try:
                            refs = [t.value for t in Tokenizer(text).items
                                    if t.type == "OPERAND" and t.subtype == "RANGE"]
                            # Relative statistical thresholds have population
                            # semantics; do not pretend they are per-cell rules.
                            if any(not _REF_RE.fullmatch(p) or
                                   not (_REF_RE.fullmatch(p)[1] and _REF_RE.fullmatch(p)[3])
                                   for ref in refs for p in ref.split(":")):
                                safe = False
                            translated = translate_formula(text, col_map, row_map)
                            if translated is None:
                                safe = False
                            else:
                                threshold.val = translated
                        except Exception:
                            safe = False
                if safe:
                    # One rule for the entire retained population, not one
                    # separate color scale/ranking for each surviving segment.
                    ws.conditional_formatting.add(" ".join(str(target) for _, target in segments), new)
                    totals["cf_kept"] = totals.get("cf_kept", 0) + 1
                else:
                    totals["cf_dropped"] = totals.get("cf_dropped", 0) + 1
                continue
            for source, target in segments:
                nfs = [_mapped_rule_formula(f, base, source, target, col_map, row_map)
                       for f in rule.formula or []]
                if any(f is None for f in nfs):
                    totals["cf_dropped"] = totals.get("cf_dropped", 0) + 1
                    continue
                new = deepcopy(rule)
                new.formula = nfs
                ws.conditional_formatting.add(str(target), new)
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
