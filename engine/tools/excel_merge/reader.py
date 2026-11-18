# -*- coding: utf-8 -*-
"""统一读取层：把 .xlsx / .xls / .csv 抹平成同一种"表"抽象。

- 流式：数据行迭代器逐行产出（value, number_format, style_source)，不整表载入。
- 结构扫描只读文件头部（表头探测窗口 + 表头区高度）。
- 合并单元格范围：xlsx 走 zip 内 XML 流式扫描（read_only 模式拿不到）；
  xls 走 xlrd merged_cells；csv 无合并概念。
"""
import csv
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from . import recognize
from .recognize import MergeRange

SUPPORTED_EXTS = {".xlsx", ".xls", ".csv"}
XLSX_WINDOW = 16
# 识别引擎开关: v2=全局候选评分+样式证据(默认), v1=逐行打分; TOOLBOX_RECOG=v1 可一键回滚
import os as _os
RECOG_ENGINE = _os.environ.get("TOOLBOX_RECOG", "v2").lower()   # 结构扫描读取的行数（探测窗口10 + 表头区高4 + 余量）


# --------------------------------------------------------------- 文件级错误

class FileError(Exception):
    """文件不可读（损坏/被占用/不支持的格式）——进报告，不中断。"""

    def __init__(self, path, reason):
        super().__init__(reason)
        self.path = str(path)
        self.reason = reason


# --------------------------------------------------------------- 结构与表源

@dataclass
class SheetStructure:
    sheet_name: str
    header_row: int                      # 1-based
    height: int                          # 表头区行数
    keys: list                           # 展平内部对齐键
    merges: list                         # 表头区内原始 MergeRange（输出模板用）
    all_merges: list = field(default_factory=list)   # 全表合并范围（数据区还原用）
    col_widths: dict = field(default_factory=dict)   # 列宽（保留样式复制用）
    orientation: str = "vertical"         # vertical 常规 | horizontal 横向表(转置读取)
    native_cross: bool = False            # 原生交叉表: 坐标即原始网格(前置区+首列字段标签+记录沿列)
    data_end_row: int = 0                 # 数据区域结束行(0=到文件末尾); 连续稀疏行尾部截断
    orientation_confident: bool = True    # 方向检测置信度(低→体检提示人工确认)
    pre_rows: list = field(default_factory=list)       # 表头前置区行(值, 共性检测/拼接用)
    title: str = None                    # 识别到的标题文本（不进总表，进报告）
    low_confidence: bool = False
    warnings: list = field(default_factory=list)
    empty: bool = False                  # 整表无数据（含仅表头）


@dataclass
class TableSource:
    """一张可合并的表 = 某文件的一个 Sheet（csv 视为单 Sheet 文件）。"""
    path: Path
    sheet_name: str
    display: str                         # 文件名 或 文件名/Sheet名
    structure: SheetStructure
    _reader: object = None

    # ---- 数据行迭代（含表头之后的所有行，未做任何清理）----
    def iter_data_rows(self):
        selected = getattr(self, "selection", None)
        for row_no, cells in self._reader.iter_data_rows(self):
            if selected:
                if not selected["data_start"] <= row_no <= selected["data_end"]:
                    continue
                cells = cells[selected["field_start"] - 1:selected["field_end"]]
            yield row_no, cells

    # ---- 表头区快照（含单元格对象，供输出模板与样式复制）----
    def header_cells(self):
        rows = self._reader.header_cells(self)
        selected = getattr(self, "selection", None)
        if selected:
            return [row[selected["field_start"] - 1:selected["field_end"]] for row in rows]
        return rows


# --------------------------------------------------------------- CSV 编码探测

def detect_csv_encoding(path):
    raw = Path(path).open("rb").read(4096)
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    for enc in ("utf-8", "gbk"):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "gb18030"


# --------------------------------------------------------------- xlsx 合并范围（XML 流式）

_MERGE_RE = re.compile(rb'<mergeCell ref="([A-Z0-9:]+)"\s*/>')
_COL_RE = re.compile(rb'<col\s([^>]*?)/>')


def _sheet_xml_paths(zf):
    """sheet名 -> worksheets xml 路径（走 workbook.xml + rels，命名空间无关）。"""
    def local(tag):
        return tag.rsplit("}", 1)[-1]

    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rid_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    name_to_rid = {}
    for el in wb_root.iter():
        if local(el.tag) == "sheet":
            name_to_rid[el.get("name")] = el.get(rid_attr)

    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {}
    for el in rels_root.iter():
        if local(el.tag) == "Relationship":
            rid_to_target[el.get("Id")] = el.get("Target")

    out = {}
    for name, rid in name_to_rid.items():
        target = rid_to_target.get(rid, "")
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target
        out[name] = target
    return out


def xlsx_column_widths(path, sheet_name):
    """流式解析 <cols> 块（位于 sheetData 之前，扫描很早结束）：自定义列宽。"""
    with zipfile.ZipFile(path) as zf:
        target = _sheet_xml_paths(zf).get(sheet_name)
        if not target:
            return {}
        widths = {}
        carry = b""
        with zf.open(target) as f:
            while True:
                chunk = f.read(1 << 18)
                if not chunk:
                    break
                data = carry + chunk
                for m in _COL_RE.finditer(data):
                    attrs = dict(re.findall(rb'(\w+)="([^"]*)"', m.group(1)))
                    try:
                        mn, mx = int(attrs[b"min"]), int(attrs[b"max"])
                        w = float(attrs[b"width"])
                    except (KeyError, ValueError):
                        continue
                    for c in range(mn, min(mx, mn + 200) + 1):
                        widths[c] = w
                if b"<sheetData" in data:
                    break
                carry = data[-128:]
        return widths


def xlsx_merge_ranges(path, sheet_name):
    """流式扫描 sheet XML 的 mergeCell（大文件不整块载入）。返回 [MergeRange]。"""
    with zipfile.ZipFile(path) as zf:
        target = _sheet_xml_paths(zf).get(sheet_name)
        if not target:
            return []
        refs = []
        carry = b""
        with zf.open(target) as f:
            while True:
                chunk = f.read(1 << 20)
                if not chunk:
                    break
                data = carry + chunk
                for m in _MERGE_RE.finditer(data):
                    refs.append(m.group(1).decode())
                # 保留尾部防标签被切断
                carry = data[-64:]
        return [MergeRange.parse(r) for r in refs]


# --------------------------------------------------------------- 各格式读取器

def xlsx_style_facts(path, sheet_name, rows, quick_only=False):
    """从包 XML 直接读样式事实(read_only 工作簿没有这些属性):
    冻结窗格(pane@ySplit) / 自动筛选(autoFilter@ref) 为 O(1);
    quick_only=False 再解析 styles.xml 与窗口行单元格样式(加粗/填充行)。"""
    from .recognize2 import StyleFacts
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    try:
        with zipfile.ZipFile(str(path)) as zf:
            target = _sheet_xml_paths(zf).get(sheet_name)
            if not target:
                return StyleFacts()
            root = ET.fromstring(zf.read(target))
            freeze_row = filter_row = 0
            pane = root.find(f".//{NS}sheetView/{NS}pane")
            if pane is not None and pane.get("state") == "frozen":
                freeze_row = int(pane.get("ySplit") or 0) + 1
            af = root.find(f"{NS}autoFilter")
            if af is not None and ":" in (af.get("ref") or ""):
                m = re.match(r"([A-Z]+)([0-9]+)", af.get("ref").split(":")[0])
                if m:
                    filter_row = int(m.group(2))
            if quick_only:
                return StyleFacts(freeze_row=freeze_row, filter_row=filter_row)
            sroot = ET.fromstring(zf.read("xl/styles.xml"))
            bold_fonts, solid_fills = set(), set()
            fonts = sroot.find(NS + "fonts")
            if fonts is not None:
                for i, f in enumerate(fonts.findall(NS + "font")):
                    b = f.find(NS + "b")
                    if b is not None and b.get("val") in (None, "1", "true"):
                        bold_fonts.add(i)
            fills = sroot.find(NS + "fills")
            if fills is not None:
                for i, fl in enumerate(fills.findall(NS + "fill")):
                    pf = fl.find(NS + "patternFill")
                    if pf is not None and pf.get("patternType") == "solid":
                        fg = pf.find(NS + "fgColor")
                        rgb = fg.get("rgb") if fg is not None else None
                        if rgb and rgb not in ("00000000", "FFFFFFFF"):
                            solid_fills.add(i)
            xfs = sroot.find(NS + "cellXfs")
            xf_bold, xf_fill = set(), set()
            if xfs is not None:
                for i, xf in enumerate(xfs.findall(NS + "xf")):
                    if int(xf.get("fontId") or 0) in bold_fonts:
                        xf_bold.add(i)
                    if int(xf.get("fillId") or 0) in solid_fills:
                        xf_fill.add(i)
            bold_rows, filled_rows = [], []
            for row in root.iter(NS + "row"):
                r = int(row.get("r") or 0)
                if r > rows:
                    break
                bold = filled = total = 0
                for c in row.findall(NS + "c"):
                    if c.get("s") is None:
                        continue
                    total += 1
                    s_idx = int(c.get("s"))
                    if s_idx in xf_bold:
                        bold += 1
                    if s_idx in xf_fill:
                        filled += 1
                if total >= 2 and bold / total > 0.6:
                    bold_rows.append(r)
                if total >= 2 and filled / total > 0.6:
                    filled_rows.append(r)
            return StyleFacts(freeze_row=freeze_row, filter_row=filter_row,
                              bold_rows=tuple(bold_rows), filled_rows=tuple(filled_rows))
    except Exception:
        return StyleFacts()


class _XlsxReader:
    def _style_facts(self, sheet_name, rows, quick_only=False):
        """窗口行的样式事实: read_only 工作簿缺这些属性, 从包 XML 直接解析。"""
        return xlsx_style_facts(self.path, sheet_name, rows, quick_only=quick_only)

    def __init__(self, path, workbook=None):
        self.path = Path(path)
        self._owns_workbook = workbook is None
        if workbook is not None:
            self.wb = workbook
            return
        try:
            self.wb = load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            raise FileError(self.path, f"文件无法读取（{type(e).__name__}）") from e

    def close(self):
        if self._owns_workbook:
            self.wb.close()

    def sheets(self):
        return [(ws.title, ws.sheet_state == "visible") for ws in self.wb.worksheets]

    def build_structure(self, sheet_name, force=0, override=None):
        ws = self.wb[sheet_name]
        window = []
        for row in ws.iter_rows(min_row=1, max_row=XLSX_WINDOW, values_only=True):
            window.append(list(row))
            if len(window) >= XLSX_WINDOW:
                break
        if not window or not any(recognize._row_nonempty(r) for r in window):
            return SheetStructure(sheet_name, 0, 0, [], [], empty=True)

        ranges = xlsx_merge_ranges(self.path, sheet_name)
        use_v2 = RECOG_ENGINE in ("v2", "v2full")
        style = self._style_facts(sheet_name, len(window), quick_only=True) if use_v2 else None
        orientation, confident = "vertical", True
        if override in ("vertical", "horizontal"):
            orientation, confident = override, True      # 人工指定优先
        elif force == 0 and window and (ws.max_row or 0) <= 20000 and (ws.max_column or 0) <= 500:
            # 方向判定保留 v1 信号净分制: v2 的双视图分数对比缺校准根基(样本量/
            # 宽度/稀疏度不同), 实测方向错误率反而升高; v2 只接管表头行/高度/置信度
            orientation, confident = recognize.detect_orientation(window)
        if orientation == "horizontal":
            if recognize.cross_table_shape(window):
                return self._structure_native_cross(ws, sheet_name, ranges, confident)
            return self._structure_transposed(ws, sheet_name, ranges, confident)
        low_confidence = False
        extra_warnings = []
        # 组合策略(默认): v1 逐行打分为基线; v1 低置信(前两名贴脸)或 v2 样式
        # 命中(冻结/筛选=用户钦定表头位置)时, 由 v2 全局评分接管 —— 常规表零
        # 扰动, 疑难表获得交界突变/同构性/样式证据更强的判断。
        # TOOLBOX_RECOG=v2full 可全量启用 v2 继续打磨。
        det = recognize.detect_header_row(window)
        if force > 0:
            header_row = force
            height = recognize.detect_region_height(window, header_row, ranges)
        else:
            header_row = det.header_row or 1
            header_row = recognize.extend_header_up(window, header_row, ranges)
            post = recognize.detect_header_row_post(window, header_row, ranges)
            if post:
                header_row, height = post
            else:
                height = recognize.detect_region_height(window, header_row, ranges)
            if use_v2 and RECOG_ENGINE != "v2full" and window:
                from . import recognize2
                if (det.low_confidence or style.freeze_row or style.filter_row)                         and not (style.bold_rows or style.filled_rows):
                    # 样式介入的场景(低置信/冻结/筛选)都升级到完整样式再裁决:
                    # 加粗/填充行是区分同级候选的关键证据
                    style = self._style_facts(sheet_name, len(window))
                guess = recognize2.detect_structure(window, ranges, style)
                styled = guess.header_row and (guess.evidence.get("freeze")
                                               or guess.evidence.get("filter"))
                if styled and guess.header_row != header_row:
                    # 冻结/筛选位置=用户亲手指定的表头, 唯一有权推翻 v1 的证据
                    header_row, height = guess.header_row, guess.height
                    extra_warnings.append(
                        f"已按冻结/筛选位置认定表头在第 {header_row} 行。")
                elif det.low_confidence and guess.header_row                         and guess.header_row != header_row and not styled:
                    # v1 没把握时不抢答: 把 v2 的不同意见放进提示供人工核对
                    extra_warnings.append(
                        f"表头位置把握较低：按第 {header_row} 行处理，"
                        f"另一种判断是第 {guess.header_row} 行，请核对。")
            elif RECOG_ENGINE == "v2full" and window:
                from . import recognize2
                guess = recognize2.detect_structure(window, ranges, style)
                if guess.header_row:
                    header_row, height = guess.header_row, guess.height
                    low_confidence = not guess.confidence
                    hint = guess.runner_up_hint()
                    if hint:
                        extra_warnings.append(hint)
        keys = recognize.flatten_keys(window, header_row, height, ranges)
        st = SheetStructure(
            sheet_name=sheet_name,
            header_row=header_row,
            height=height,
            keys=keys,
            merges=[mg for mg in ranges if mg.min_row >= header_row
                    and mg.max_row <= header_row + height - 1],
            all_merges=ranges,
            col_widths=xlsx_column_widths(self.path, sheet_name),
            orientation_confident=confident,
            pre_rows=[list(r) for r in window[:header_row - 1]],
            title=recognize.detect_title(window, header_row, ranges),
            low_confidence=det.low_confidence or low_confidence,
        )
        warn = recognize.out_of_bounds_merge_warning(ranges, header_row, height)
        if warn:
            st.warnings.append(warn)
        st.warnings.extend(extra_warnings)
        if det.header_row == 0:
            st.warnings.append("未探测到表头，已按第 1 行处理（可手动指定）")
        if not confident:
            st.warnings.append("方向待确认（疑似交叉表），可在体检中手动切换纵向/横向")
        st.data_end_row = detect_data_end_row(ws, header_row, height)
        st.data_end = st.data_end_row  # 别名: fidelity 用 data_end 访问
        return st

    def _structure_native_cross(self, ws, sheet_name, ranges, confident=True):
        """原生交叉表: 前置区+首列字段标签+记录沿列。结构坐标即原始网格,
        与纵向表共用表头/数据区识别, 仅方向语义不同(记录=原列)。"""
        window = [list(r) for r in ws.iter_rows(min_row=1, max_row=XLSX_WINDOW,
                                                values_only=True)]
        det = recognize.detect_header_row(window)
        header_row = det.header_row or 1
        header_row = recognize.extend_header_up(window, header_row, ranges)
        post = recognize.detect_header_row_post(window, header_row, ranges)
        if post:
            header_row, height = post
        else:
            height = recognize.detect_region_height(window, header_row, ranges)
        keys = recognize.flatten_keys(window, header_row, height, ranges)
        st = SheetStructure(
            sheet_name=sheet_name,
            header_row=header_row,
            height=height,
            keys=keys,
            merges=[mg for mg in ranges if mg.min_row >= header_row
                    and mg.max_row <= header_row + height - 1],
            all_merges=ranges,
            col_widths=xlsx_column_widths(self.path, sheet_name),
            orientation="horizontal",
            native_cross=True,
            orientation_confident=confident,
            pre_rows=[list(r) for r in window[:header_row - 1]],
            title=recognize.detect_title(window, header_row, ranges),
            low_confidence=det.low_confidence,
        )
        warn = recognize.out_of_bounds_merge_warning(ranges, header_row, height)
        if warn:
            st.warnings.append(warn)
        if det.header_row == 0:
            st.warnings.append("未探测到表头，已按第 1 行处理（可手动指定）")
        if not confident:
            st.warnings.append("方向待确认（疑似交叉表），可在体检中手动切换纵向/横向")
        st.data_end_row = detect_data_end_row(ws, header_row, height)
        st.data_end = st.data_end_row
        # 记录列范围(标签列之后的表头行非空列): 拆分/预览按"字段=行,记录=列"语义取值
        header_values = window[header_row - 1] if header_row - 1 < len(window) else []
        last = max((c for c in range(2, len(header_values) + 1)
                    if header_values[c - 1] not in (None, "")), default=1)
        st.cross_records = (2, max(2, last))
        return st

    def _structure_transposed(self, ws, sheet_name, ranges, confident=True):
        """横向表：在转置视图（行=原列）上跑同一套识别逻辑。"""
        buf = [list(row) for row in ws.iter_rows(values_only=True)]
        ncols = max((len(r) for r in buf), default=0)
        tv = [[(buf[r][c] if c < len(buf[r]) else None) for r in range(len(buf))]
              for c in range(min(ncols, XLSX_WINDOW))]
        ranges = [mg.transposed() for mg in ranges]
        det = recognize.detect_header_row(tv)
        header_row = det.header_row or 1
        header_row = recognize.extend_header_up(tv, header_row, ranges)
        height = recognize.detect_region_height(tv, header_row, ranges)
        keys = recognize.flatten_keys(tv, header_row, height, ranges)
        st = SheetStructure(
            sheet_name=sheet_name, header_row=header_row, height=height,
            keys=keys, orientation="horizontal", orientation_confident=confident,
            pre_rows=[list(r) for r in tv[:header_row - 1]],
            merges=[mg for mg in ranges if mg.min_row >= header_row
                    and mg.max_row <= header_row + height - 1],
            all_merges=ranges)
        warn = recognize.out_of_bounds_merge_warning(ranges, header_row, height)
        if warn:
            st.warnings.append(warn)
        if det.header_row == 0:
            st.warnings.append("未探测到表头，已按第 1 行处理（可手动指定）")
        return st

    def iter_data_rows(self, source):
        ws = self.wb[source.sheet_name]
        st = source.structure
        start = st.header_row + st.height
        end = st.data_end_row if st.data_end_row > 0 else ws.max_row
        if st.orientation == "horizontal" and not getattr(st, "native_cross", False):
            buf = [list(row) for row in ws.iter_rows()]
            ncols = max((len(r) for r in buf), default=0)
            for c in range(start - 1, ncols):
                yield c + 1, [(buf[r][c] if c < len(buf[r]) else None)
                              for r in range(len(buf))]
            return
        if getattr(st, "native_cross", False):
            # 原生交叉表: 记录=原始列(标签列之后), 整列产出与转置路径同构
            buf = [list(row) for row in ws.iter_rows()]
            ncols = max((len(r) for r in buf), default=0)
            for c in range(1, ncols):
                yield c + 1, [(buf[r][c] if c < len(buf[r]) else None)
                              for r in range(len(buf))]
            return
        for row_no, row in enumerate(ws.iter_rows(min_row=start, max_row=end), start=start):
            yield row_no, [c for c in row]   # ReadOnlyCell 列表（带 number_format/font/fill）

    def header_cells(self, source):
        ws = self.wb[source.sheet_name]
        st = source.structure
        if st.orientation == "horizontal" and not getattr(st, "native_cross", False):
            buf = [list(row) for row in ws.iter_rows()]
            return [[(buf[r][c] if c < len(buf[r]) else None) for r in range(len(buf))]
                    for c in range(st.header_row - 1, st.header_row - 1 + st.height)]
        rows = []
        for row in ws.iter_rows(min_row=st.header_row, max_row=st.header_row + st.height - 1):
            rows.append([c for c in row])
        return rows


class _XlsReader:
    def __init__(self, path):
        import xlrd
        self.path = Path(path)
        try:
            try:
                self.book = xlrd.open_workbook(str(self.path), formatting_info=True)
            except NotImplementedError:
                self.book = xlrd.open_workbook(str(self.path))
        except Exception as e:
            raise FileError(self.path, f"文件无法读取（{type(e).__name__}）") from e
        self._has_merge = hasattr(self.book.sheet_by_index(0), "merged_cells") \
            and getattr(self.book.sheet_by_index(0), "merged_cells", None) is not None

    def close(self):
        pass

    def sheets(self):
        return [(s.name, True) for s in self.book.sheets()]

    def _conv(self, sheet, r, c):
        import xlrd
        try:
            cell = sheet.cell(r, c)
        except Exception:
            return None
        if cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            v = cell.value
            return int(v) if float(v).is_integer() else v
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                tt = xlrd.xldate_as_tuple(cell.value, self.book.datemode)
                return dt_datetime(tt)
            except Exception:
                return cell.value
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)
        return None

    def _style_facts(self, sheet_name, rows, quick_only=False):
        """窗口行的样式事实: 冻结窗格/自动筛选/加粗行/填充行(占比>60%非空格)。
        quick_only=True 只取 O(1) 的冻结/筛选 —— 逐格的加粗/填充扫描较贵,
        仅在需要 v2 全局裁决(低置信)时才做。"""
        from .recognize2 import StyleFacts
        try:
            ws = self.wb[sheet_name]
            freeze_row = 0
            fp = getattr(ws, "freeze_panes", None)
            if fp and isinstance(fp, str):
                m = re.match(r"([A-Z]+)(\d+)", fp)
                if m:
                    freeze_row = int(m.group(2))
            filter_row = 0
            af = getattr(ws.auto_filter, "ref", None) if getattr(ws, "auto_filter", None) else None
            if af and ":" in str(af):
                m = re.match(r"([A-Z]+)(\d+)", str(af).split(":")[0])
                if m:
                    filter_row = int(m.group(2))
            if quick_only:
                return StyleFacts(freeze_row=freeze_row, filter_row=filter_row)
            bold_rows, filled_rows = [], []
            limit_col = 60
            for r in range(1, min(rows, XLSX_WINDOW + 8) + 1):
                bold = filled = total = 0
                for c in range(1, limit_col + 1):
                    cell = ws.cell(r, c)
                    if cell.value in (None, ""):
                        continue
                    total += 1
                    try:
                        if cell.font and cell.font.bold:
                            bold += 1
                    except Exception:
                        pass
                    try:
                        if cell.fill and cell.fill.fill_type == "solid":
                            fg = cell.fill.fgColor
                            if fg and fg.rgb not in (None, "00000000", "FFFFFFFF"):
                                filled += 1
                    except Exception:
                        pass
                if total >= 2 and bold / total > 0.6:
                    bold_rows.append(r)
                if total >= 2 and filled / total > 0.6:
                    filled_rows.append(r)
            return StyleFacts(freeze_row=freeze_row, filter_row=filter_row,
                              bold_rows=tuple(bold_rows), filled_rows=tuple(filled_rows))
        except Exception:
            return StyleFacts()

    def build_structure(self, sheet_name, force=0, override=None):
        sheet = self.book.sheet_by_name(sheet_name)
        if sheet.nrows == 0:
            return SheetStructure(sheet_name, 0, 0, [], [], empty=True)
        window = [[self._conv(sheet, r, c) for c in range(sheet.ncols)]
                  for r in range(min(sheet.nrows, XLSX_WINDOW))]
        ranges = [MergeRange(c1 + 1, r1 + 1, c2 + 1, r2 + 1)
                  for r1, r2, c1, c2 in (sheet.merged_cells if self._has_merge else [])]
        det = recognize.detect_header_row(window)
        header_row = force if force > 0 else recognize.extend_header_up(
            window, det.header_row or 1, ranges)
        height = recognize.detect_region_height(window, header_row, ranges)
        keys = recognize.flatten_keys(window, header_row, height, ranges)
        widths = {}
        try:
            ci = getattr(sheet, "colinfo_map", None) or {}
            widths = {c + 1: round(v.width / 256.0, 2) for c, v in ci.items()}
        except Exception:
            pass
        return SheetStructure(
            sheet_name, header_row, height, keys,
            [mg for mg in ranges if mg.min_row >= header_row and mg.max_row <= header_row + height - 1],
            all_merges=ranges, col_widths=widths,
            pre_rows=[list(r) for r in window[:header_row - 1]],
            title=recognize.detect_title(window, header_row, ranges),
            low_confidence=det.low_confidence,
            warnings=(["未探测到表头，已按第 1 行处理（可手动指定）"] if det.header_row == 0 else []),
        )

    def _row_cells(self, sheet, r, ncols):
        return [_PseudoCell(self._conv(sheet, r, c)) for c in range(ncols)]

    def iter_data_rows(self, source):
        sheet = self.book.sheet_by_name(source.sheet_name)
        st = source.structure
        start = st.header_row + st.height - 1          # 0-based 起始行
        for r in range(start, sheet.nrows):
            yield r + 1, [self._conv(sheet, r, c) for c in range(sheet.ncols)]  # 原生值直传

    def header_cells(self, source):
        sheet = self.book.sheet_by_name(source.sheet_name)
        st = source.structure
        return [self._row_cells(sheet, r, sheet.ncols)
                for r in range(st.header_row - 1, st.header_row - 1 + st.height)]


def dt_datetime(tt):
    import datetime as _dt
    return _dt.datetime(tt[0], tt[1], tt[2], tt[3], tt[4], int(tt[5]))


class _PseudoCell:
    """xls/csv 行单元格的最小替身：值 + General 格式。"""
    __slots__ = ("value", "number_format", "font", "fill", "border", "alignment")

    def __init__(self, value):
        self.value = value
        self.number_format = "General"
        self.font = None
        self.fill = None
        self.border = None
        self.alignment = None


class _CsvReader:
    def __init__(self, path):
        self.path = Path(path)
        self.encoding = detect_csv_encoding(self.path)
        try:
            with self.path.open("r", encoding=self.encoding, newline="") as f:
                head = [row for _, row in zip(range(XLSX_WINDOW), csv.reader(f))]
        except Exception as e:
            raise FileError(self.path, f"文件无法读取（{type(e).__name__}）") from e
        self.window = head
        self.sheet_name = self.path.stem

    def close(self):
        pass

    def sheets(self):
        return [(self.sheet_name, True)]

    def _reopen(self):
        return csv.reader(self.path.open("r", encoding=self.encoding, newline=""))

    def _style_facts(self, sheet_name, rows, quick_only=False):
        """窗口行的样式事实: 冻结窗格/自动筛选/加粗行/填充行(占比>60%非空格)。
        quick_only=True 只取 O(1) 的冻结/筛选 —— 逐格的加粗/填充扫描较贵,
        仅在需要 v2 全局裁决(低置信)时才做。"""
        from .recognize2 import StyleFacts
        try:
            ws = self.wb[sheet_name]
            freeze_row = 0
            fp = getattr(ws, "freeze_panes", None)
            if fp and isinstance(fp, str):
                m = re.match(r"([A-Z]+)(\d+)", fp)
                if m:
                    freeze_row = int(m.group(2))
            filter_row = 0
            af = getattr(ws.auto_filter, "ref", None) if getattr(ws, "auto_filter", None) else None
            if af and ":" in str(af):
                m = re.match(r"([A-Z]+)(\d+)", str(af).split(":")[0])
                if m:
                    filter_row = int(m.group(2))
            if quick_only:
                return StyleFacts(freeze_row=freeze_row, filter_row=filter_row)
            bold_rows, filled_rows = [], []
            limit_col = 60
            for r in range(1, min(rows, XLSX_WINDOW + 8) + 1):
                bold = filled = total = 0
                for c in range(1, limit_col + 1):
                    cell = ws.cell(r, c)
                    if cell.value in (None, ""):
                        continue
                    total += 1
                    try:
                        if cell.font and cell.font.bold:
                            bold += 1
                    except Exception:
                        pass
                    try:
                        if cell.fill and cell.fill.fill_type == "solid":
                            fg = cell.fill.fgColor
                            if fg and fg.rgb not in (None, "00000000", "FFFFFFFF"):
                                filled += 1
                    except Exception:
                        pass
                if total >= 2 and bold / total > 0.6:
                    bold_rows.append(r)
                if total >= 2 and filled / total > 0.6:
                    filled_rows.append(r)
            return StyleFacts(freeze_row=freeze_row, filter_row=filter_row,
                              bold_rows=tuple(bold_rows), filled_rows=tuple(filled_rows))
        except Exception:
            return StyleFacts()

    def build_structure(self, sheet_name, force=0, override=None):
        window = [r for r in self.window]  # 前导空行参与窗口但计分为 0
        if not any(recognize._row_nonempty(r) for r in window):
            return SheetStructure(sheet_name, 0, 0, [], [], empty=True)
        orientation, confident = "vertical", True
        if override in ("vertical", "horizontal"):
            orientation, confident = override, True
        elif force == 0:
            orientation, confident = recognize.detect_orientation(window)
        if orientation == "horizontal":
            rows = [r for r in self._reopen()]
            ncols = max((len(r) for r in rows), default=0)
            tv = [[(rows[r][c] if c < len(rows[r]) else None)
                   for r in range(len(rows))] for c in range(min(ncols, XLSX_WINDOW))]
            det = recognize.detect_header_row(tv)
            header_row = det.header_row or 1
            keys = recognize.flatten_keys(tv, header_row, 1, [])
            return SheetStructure(sheet_name, header_row, 1, keys, [],
                                  orientation="horizontal",
                                  orientation_confident=confident,
                                  pre_rows=[list(r) for r in tv[:header_row - 1]],
                                  low_confidence=det.low_confidence)
        det = recognize.detect_header_row(window)
        header_row = force if force > 0 else (det.header_row or 1)
        keys = recognize.flatten_keys(window, header_row, 1, [])
        return SheetStructure(
            sheet_name, header_row, 1, keys, [],
            pre_rows=[list(r) for r in window[:header_row - 1]],
            title=recognize.detect_title(window, header_row, []),
            low_confidence=det.low_confidence,
            warnings=(["未探测到表头，已按第 1 行处理（可手动指定）"] if det.header_row == 0 else []),
        )

    def iter_data_rows(self, source):
        start = source.structure.header_row          # csv 高度恒 1
        if source.structure.orientation == "horizontal":
            rows = [r for r in self._reopen()]
            ncols = max((len(r) for r in rows), default=0)
            for c in range(start, ncols + 1):
                yield c, [((rows[r][c - 1] if c - 1 < len(rows[r]) else "") or None)
                          for r in range(len(rows))]
            return
        for i, row in enumerate(self._reopen(), start=1):
            if i > start:
                yield i, [v if v != "" else None for v in row]   # 原生值直传

    def header_cells(self, source):
        r = source.structure.header_row
        if source.structure.orientation == "horizontal":
            rows = [row for row in self._reopen()]
            col = [((rows[i][r - 1] if r - 1 < len(rows[i]) else "") or None)
                   for i in range(len(rows))]
            return [[_PseudoCell(v) for v in col]]
        row = self.window[r - 1] if r - 1 < len(self.window) else []
        return [[_PseudoCell(v if v != "" else None) for v in row]]


# --------------------------------------------------------------- 入口

def detect_data_end_row(ws, header_row, height, max_check=200):
    """检测数据区域实际结束行: 连续≥10行完全全空(0非空格) → 截断.

    注意: 只用"完全空行"截断, 不用"稀疏行"截断.
    稀疏行(如仅有序号的行)可能是表格结构的一部分(用户确认保留).
    返回最后一行有效数据的行号(1-based), 0=到文件末尾.
    实现约束: 单次 iter_rows 流式扫描. 只读工作表的 cell() 会逐格重解析
    整个 sheet XML(O(行×列)次全表解析), 在此使用会让预览慢两个数量级.
    """
    data_start = header_row + height
    max_row = ws.max_row or 0
    if data_start > max_row:
        return 0
    max_col = min(ws.max_column or 1, 20)
    empty_run = 0
    last_nonempty_row = 0
    checked = 0
    for row in ws.iter_rows(min_row=data_start,
                            max_row=min(max_row, data_start + max_check - 1),
                            max_col=max_col, values_only=True):
        checked += 1
        if any(v not in (None, "") for v in row):
            empty_run = 0
            last_nonempty_row = data_start + checked - 1
        else:
            empty_run += 1
            if empty_run >= 10:
                return last_nonempty_row
    return 0


def open_reader(path):
    ext = Path(path).suffix.lower()
    if ext not in SUPPORTED_EXTS:
        raise FileError(path, f"不支持的格式 {ext}（支持 xlsx/xls/csv）")
    if not Path(path).exists():
        raise FileError(path, "文件不存在")
    if ext == ".xlsx":
        return _XlsxReader(path)
    if ext == ".xls":
        return _XlsReader(path)
    return _CsvReader(path)


def list_tables(path, include_hidden=False, force_header_row=0, orientation_override=None):
    """打开文件 → [(TableSource, 已跳过原因?)]。损坏文件抛 FileError 由上层记报告。"""
    reader = open_reader(path)
    out = []
    for name, visible in reader.sheets():
        if not visible and not include_hidden:
            out.append((None, f"隐藏 Sheet「{name}」已跳过"))
            continue
        st = reader.build_structure(name, force_header_row, orientation_override)
        src = TableSource(Path(path), name, display_table(Path(path), name),
                          st, reader)
        out.append((src, None))
    return out, reader


def display_table(path, sheet_name):
    stem = Path(path).stem
    if Path(path).suffix.lower() == ".csv":
        return stem
    return f"{stem}/{sheet_name}"
