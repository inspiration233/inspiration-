# -*- coding: utf-8 -*-
"""流式写出层。

关键实现约束（实测结论）：
- openpyxl write_only 模式不支持 merge_cells → 先流式保存，再后处理 zip 把
  mergeCells 注入 sheet XML（插在 </sheetData> 之后，符合 schema 顺序）。
- openpyxl 写出数字按 %.16g 序列化 → 超过 15 位有效数字的数值必须转文本防损。
- 日期值 append 时自动带 yyyy-mm-dd 类格式；统一样式模式仍透传源 number_format。
"""
import re
import zipfile
from copy import copy
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .recognize import MergeRange

XLSX_MAX_ROWS = 1_048_576
BIG_NUMBER = 10 ** 15          # 15 位以内安全（%.16g 实测）

_ILLEGAL_SHEET = re.compile(r"[:\\/?*\[\]]")
_THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
_HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
_HEADER_FILL = PatternFill("solid", fgColor="4472C4")


def sanitize_sheet_name(name, used):
    name = _ILLEGAL_SHEET.sub(" ", str(name)).strip() or "Sheet"
    name = name[:31]
    base, idx = name, 1
    while name in used:
        suffix = f"({idx + 1})"
        name = base[: 31 - len(suffix)] + suffix
        idx += 1
    used.add(name)
    return name


def big_number_guard(value):
    """>15 位数值转文本（保值优先于保类型）。返回 (新值, 是否转换)。"""
    if isinstance(value, bool):
        return value, False
    if isinstance(value, int) and abs(value) >= BIG_NUMBER:
        return str(value), True
    if isinstance(value, float) and abs(value) >= BIG_NUMBER:
        return str(int(value)) if value.is_integer() else repr(value), True
    return value, False


class OutputBook:
    """一个输出工作簿；按需创建分页（超行数上限自动续写）。"""

    def __init__(self, path, style_mode="unified"):
        self.path = Path(path)
        self.style_mode = style_mode
        self.wb = Workbook(write_only=True)
        self._used_names = set()
        self._pages_merges = []          # 每页的合并 ref（与 sheet 创建顺序一致）
        self.big_numbers_converted = 0
        self._page_seq = {}              # 基名 -> 已建页数

    def _next_sheet_name(self, base):
        n = self._page_seq.get(base, 0) + 1
        self._page_seq[base] = n
        if n == 1:
            return sanitize_sheet_name(base, self._used_names)
        suffix = f"({n})"
        return sanitize_sheet_name(base[: 31 - len(suffix)] + suffix, self._used_names)

    def add_table(self, base_name, header_rows, header_merges, extra_headers=None,
                  source_label=None):
        """header_rows: 模板表头单元格行([[cell,...]])；extra_headers: 追加列的表头文本列表。"""
        return _TablePage(self, base_name, header_rows, header_merges,
                          extra_headers or [], source_label)

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(self.path)
        if any(m for m in self._pages_merges):
            _inject_merges(self.path, self._pages_merges)


class _TablePage:
    def __init__(self, book, base_name, header_rows, header_merges, extra_headers, source_label):
        self.book = book
        self.height = max(1, len(header_rows))
        self.width = max((len(r) for r in header_rows), default=0) + len(extra_headers) \
            + (1 if source_label else 0)
        self.base_name = base_name
        self.row_index = 0
        self._col_widths = {}
        self.page_epoch = 0
        self._new_page(header_rows, header_merges, extra_headers, source_label, first=True)

    # ---------------------------------------------------------------- 分页
    def _new_page(self, header_rows, header_merges, extra_headers, source_label, first):
        book = self.book
        name = book._next_sheet_name(self.base_name if first else self.base_name)
        self.ws = book.wb.create_sheet(name)
        self.row_index = 0
        self.page_epoch = 0                      # 分页代际（数据区合并翻译用）
        merges = self._write_header(header_rows, header_merges, extra_headers, source_label)
        self._page_merges = merges               # 同一 list 对象，存于 book._pages_merges
        book._pages_merges.append(merges)
        self._apply_widths()

    def add_extra_merge(self, c1, r1, c2, r2):
        """登记数据区合并（输出坐标，随 mergeCells 后处理一并注入）。"""
        if (c1, r1) != (c2, r2):
            self._page_merges.append((c1, r1, c2, r2))

    def set_col_widths(self, widths):
        """列宽（1-based 列号 → 字符宽）；分页续写时自动重新应用。"""
        self._col_widths.update(widths or {})
        self._apply_widths()

    def _apply_widths(self):
        for col, w in getattr(self, "_col_widths", {}).items():
            try:
                self.ws.column_dimensions[col_letter(col)].width = w
            except Exception:
                pass

    def _write_header(self, header_rows, header_merges, extra_headers, source_label):
        """写出表头区：模板行原样（含合并结构），追加列纵向合并跨满表头高度。"""
        height = max(1, len(header_rows))
        template_width = max((len(r) for r in header_rows), default=0)
        merges = []
        # 模板行
        for r_idx, row in enumerate(header_rows, start=1):
            cells = []
            for c_idx in range(1, template_width + 1):
                src = row[c_idx - 1] if c_idx - 1 < len(row) else None
                cells.append(self._header_cell(src))
            # 追加列表头（首行写文本，其余行留空 + 纵向合并）
            for text in extra_headers:
                cells.append(self._header_cell(None, text=text) if r_idx == 1
                             else self._header_cell(None, text=None))
            if source_label:
                cells.append(self._header_cell(None, text=source_label) if r_idx == 1
                             else self._header_cell(None, text=None))
            self.ws.append(cells)
            self.row_index += 1
        # 模板合并（调用方已平移到输出坐标并裁剪到表头区高度内）
        for mg in header_merges:
            merges.append((mg.min_col, mg.min_row, mg.max_col, mg.max_row))
        # 追加列纵向合并
        total_width = template_width + len(extra_headers) + (1 if source_label else 0)
        if height > 1:
            for c in range(template_width + 1, total_width + 1):
                merges.append((c, 1, c, height))
        self.ws.freeze_panes = f"A{height + 1}"
        return merges

    def _header_cell(self, src, text=None):
        value = text if text is not None else (getattr(src, "value", None) if src is not None else None)
        wc = WriteOnlyCell(self.ws, value=value if value is not None else "")
        if self.book.style_mode == "preserve" and src is not None:
            _copy_style(wc, src)
        else:
            wc.font = copy(_HEADER_FONT)
            wc.fill = copy(_HEADER_FILL)
            wc.alignment = Alignment(horizontal="center", vertical="center")
            wc.border = copy(_THIN)
        return wc

    # ---------------------------------------------------------------- 数据行
    def write_row(self, cells):
        """cells: 与 self.width 等长的源单元格列表（None 补位）。超行数上限由上层分页。

        快速通道：General 格式且非保留样式 → 直接追加原生值（绕过 WriteOnlyCell，
        大规模合并的主要性能来源；日期值追加时 openpyxl 自动套日期格式）。
        """
        preserve = self.book.style_mode == "preserve"
        out = []
        for src in cells:
            if src is None:
                out.append(None)
                continue
            value = getattr(src, "value", src)          # 原生值 或 单元格对象
            value, converted = big_number_guard(value)
            if converted:
                self.book.big_numbers_converted += 1
            ss = getattr(src, "_style_src", src)        # 序号重排值携带原样式
            has_style = hasattr(ss, "value")
            if preserve and has_style:
                wc = WriteOnlyCell(self.ws, value=value if value is not None else "")
                _copy_style(wc, ss)
                out.append(wc)
                continue
            fmt = getattr(ss, "number_format", None) if has_style else None
            if fmt and fmt != "General":
                wc = WriteOnlyCell(self.ws, value=value if value is not None else "")
                wc.number_format = fmt
                out.append(wc)
            else:
                out.append(value)
        self.ws.append(out)
        self.row_index += 1

    def paginate_restart(self, header_rows, header_merges, extra_headers, source_label):
        self.page_epoch += 1                  # 代际+1：跨页的合并范围将被上层丢弃
        self._new_page(header_rows, header_merges, extra_headers, source_label, first=False)


def _copy_style(dst, src):
    """保留样式模式：从源单元格复制样式属性（read_only 单元格可读这些属性）。"""
    try:
        if getattr(src, "font", None):
            dst.font = copy(src.font)
        if getattr(src, "fill", None) and getattr(src.fill, "patternType", None):
            dst.fill = copy(src.fill)
        if getattr(src, "border", None):
            dst.border = copy(src.border)
        if getattr(src, "alignment", None):
            dst.alignment = copy(src.alignment)
        fmt = getattr(src, "number_format", None)
        if fmt and fmt != "General":
            dst.number_format = fmt
    except Exception:
        pass  # 样式尽力而为，不因样式失败丢数据


def _inject_merges(path, pages_merges):
    """把每页的合并范围注入已保存 xlsx 的 sheet XML（write_only 不支持 merge_cells）。"""
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        sheet_no = 0
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", item.filename):
                sheet_no += 1
                merges = pages_merges[sheet_no - 1] if sheet_no - 1 < len(pages_merges) else []
                if merges:
                    refs = "".join(
                        f'<mergeCell ref="{MergeRange(min_col, min_row, max_col, max_row).ref()}"/>'
                        for min_col, min_row, max_col, max_row in merges
                        if not (min_col == max_col and min_row == max_row)
                    )
                    if refs:
                        block = f'<mergeCells count="{len(merges)}">{refs}</mergeCells>'.encode()
                        xml = data
                        pos = xml.find(b"</sheetData>")
                        if pos >= 0:
                            pos += len(b"</sheetData>")
                        else:
                            pos = xml.find(b"<pageMargins")
                            pos = pos if pos >= 0 else len(xml)
                        data = xml[:pos] + block + xml[pos:]
            zout.writestr(item, data)
    tmp.replace(path)


# MergeRange.ref() 输出 A1 形式（补在 recognize 中未提供）
def _mergeref(self):
    return f"{col_letter(self.min_col)}{self.min_row}:{col_letter(self.max_col)}{self.max_row}"


def col_letter(col):
    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


MergeRange.ref = _mergeref
