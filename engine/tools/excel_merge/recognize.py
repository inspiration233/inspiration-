# -*- coding: utf-8 -*-
"""表头智能识别：列名规范化 / 表头行探测 / 多行表头展平键 / 合计小计行剔除。

对应需求 §3.2（V1.2/V1.7 口径）：展平只作内部对齐键，输出结构照表头模板原样保留。
"""
import re
import unicodedata
from dataclasses import dataclass, field

DEFAULT_SUMMARY_KEYWORDS = ("合计", "小计", "总计", "汇总", "累计")
HEADER_SCAN_WINDOW = 10          # 表头探测扫描窗口（可配置 1-50）
HEADER_MAX_HEIGHT = 4            # 多行表头最大高度
LOW_CONFIDENCE_MARGIN = 0.06     # 前两名分差小于此值 → 低置信度提示

_NUM_RE = re.compile(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$")
_DATE_RE = re.compile(r"^\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?$|^\d{1,2}月\d{1,2}日$")


def normalize_name(v):
    """列名规范化（仅用于匹配，不改输出）：NFKC 全角→半角、折叠空格、去首尾、英文小写。"""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def is_numberlike(v):
    """探测用：数值/日期形态（xlsx 的真数值、csv 的数字与日期样式文本）。"""
    if isinstance(v, bool):
        return True
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        return bool(_NUM_RE.match(s) or _DATE_RE.match(s))
    return False  # datetime 等


def is_summary_row(values, keywords=DEFAULT_SUMMARY_KEYWORDS, check_cols=2):
    """合计/小计行判定：前两列任一文本（规范化后）等于或以关键词结尾。
    先做原文子串预筛（C 级快），避免逐行 NFKC 规范化的大规模开销。"""
    for v in list(values)[:check_cols]:
        if isinstance(v, str) and v and any(k in v for k in keywords):
            t = normalize_name(v)
            if t and any(t == k or t.endswith(k) for k in keywords):
                return True
    return False


# --------------------------------------------------------------- 合并单元格范围

@dataclass(frozen=True)
class MergeRange:
    min_col: int
    min_row: int
    max_col: int
    max_row: int

    @classmethod
    def parse(cls, ref: str):
        a, _, b = ref.partition(":")
        b = b or a
        c1, r1 = _coord(a)
        c2, r2 = _coord(b)
        return cls(min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))

    def covers(self, row, col):
        return self.min_row <= row <= self.max_row and self.min_col <= col <= self.max_col

    def is_anchor(self, row, col):
        return row == self.min_row and col == self.min_col

    def horizontal(self):
        return self.max_col > self.min_col

    def vertical(self):
        return self.max_row > self.min_row


def _coord(cell_ref: str):
    m = re.match(r"([A-Z]+)(\d+)", cell_ref.strip().upper())
    if not m:
        return 1, 1
    letters, digits = m.group(1), int(m.group(2))
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return col, digits


# --------------------------------------------------------------- 表头行探测

def _row_nonempty(cells):
    return [v for v in cells if v not in (None, "")]


def score_header_row(cells, width):
    """单行“像表头”的得分：密度 + 文本占比 + 取值唯一性。空行 0 分。"""
    non_empty = _row_nonempty(cells)
    if not non_empty:
        return 0.0
    density = len(non_empty) / max(width, 1)
    texts = [v for v in non_empty if not is_numberlike(v)]
    text_ratio = len(texts) / len(non_empty)
    uniq = len({normalize_name(t) for t in texts}) / max(len(texts), 1)
    return 0.4 * density + 0.4 * text_ratio + 0.2 * uniq


@dataclass
class HeaderDetect:
    header_row: int = 0          # 1-based；0 = 未检出（回退第 1 行并警告）
    scores: dict = field(default_factory=dict)
    low_confidence: bool = False
    title: str = None            # 表头上方的标题/制表行文本（若有）


def detect_header_row(window_rows, window=HEADER_SCAN_WINDOW):
    """window_rows: 前 N 行的值列表（1-based 索引即行号）。平手取更早的行。"""
    width = max((len(r) for r in window_rows), default=0)
    scores = {}
    for i, cells in enumerate(window_rows[:window], start=1):
        scores[i] = score_header_row(cells, width)
    ranked = sorted(((s, -r) for r, s in scores.items() if s > 0), reverse=True)
    if not ranked:
        return HeaderDetect(0, scores, False, None)
    best_score, neg_row = ranked[0]
    best_row = -neg_row
    low = len(ranked) > 1 and (best_score - ranked[1][0]) < LOW_CONFIDENCE_MARGIN
    return HeaderDetect(best_row, scores, low, None)


def detect_title(window_rows, header_row, ranges):
    """表头上方：单格非空且被跨列合并（≥3列）覆盖的行 → 标题文本。"""
    for r in range(1, header_row):
        cells = window_rows[r - 1]
        non_empty = _row_nonempty(cells)
        if len(non_empty) != 1:
            continue
        col = next(i + 1 for i, v in enumerate(cells) if v not in (None, ""))
        for mg in ranges:
            if mg.covers(r, col) and mg.max_col - mg.min_col >= 2:
                return str(non_empty[0])
    return None


def extend_header_up(window_rows, header_row, ranges):
    """多行表头的父级行因合并占位得分偏低 → 选中行后向上回溯：
    上一行 ≥2 个非空纯文本格、且存在跨列/跨行合并结构信号 → 表头实际从该行开始。
    标题/制表行只有 1 个非空格，天然不会被误吸收。
    """
    r = header_row - 1
    while r >= 1:
        cells = window_rows[r - 1]
        non_empty = _row_nonempty(cells)
        if len(non_empty) < 2 or any(is_numberlike(v) for v in non_empty):
            break
        has_merge_signal = any(
            (mg.horizontal() or mg.vertical()) and mg.min_row == r for mg in ranges
        )
        if not has_merge_signal:
            break
        header_row = r
        r -= 1
    return header_row


# --------------------------------------------------------------- 表头区高度与展平键

def detect_region_height(window_rows, header_row, ranges, max_height=HEADER_MAX_HEIGHT):
    """从表头行向下扩展表头区：
    扩展条件 = 下一行非空、纯文本（无数值/日期形态），且
    (a) 有从表头区出发的合并覆盖到该行，或 (b) 上一区域行存在跨列合并（父子结构信号）。
    内容信号优先于合并范围（越界合并不扩展）。
    """
    h = 1
    while header_row + h <= len(window_rows) and h < max_height:
        r = header_row + h
        cells = window_rows[r - 1]
        non_empty = _row_nonempty(cells)
        if not non_empty or any(is_numberlike(v) for v in non_empty):
            break
        covered = any(mg.covers(r, c) and mg.min_row <= header_row + h - 1 for mg in ranges for c in range(1, len(cells) + 1))
        parent_merge_above = any(
            mg.horizontal() and mg.min_row >= header_row and mg.min_row < r
            for mg in ranges
        )
        if not (covered or parent_merge_above):
            break
        h += 1
    return h


def _merged_value(window_rows, ranges, row_idx, col_idx):
    """取表头区 (row_idx, col_idx) 的逻辑值：合并范围内取锚点值，非锚点为 None。"""
    cells = window_rows[row_idx - 1]
    raw = cells[col_idx - 1] if col_idx <= len(cells) else None
    for mg in ranges:
        if mg.covers(row_idx, col_idx):
            if mg.is_anchor(row_idx, col_idx):
                return raw
            anchor_cells = window_rows[mg.min_row - 1]
            return anchor_cells[mg.min_col - 1] if mg.min_col <= len(anchor_cells) else None
    return raw


def flatten_keys(window_rows, header_row, height, ranges):
    """生成每列的内部对齐键：纵向合并取单值；跨列+子级拼 '父/子'；空键由调用方命名 未命名n。"""
    width = max(len(r) for r in window_rows[header_row - 1: header_row - 1 + height])
    keys = []
    for c in range(1, width + 1):
        parts = []
        for r in range(header_row, header_row + height):
            n = normalize_name(_merged_value(window_rows, ranges, r, c))
            if n and n not in parts:
                parts.append(n)
        keys.append("/".join(parts))
    # 同名键（含合并继承出的同名）全部加 (n) 序号；唯一键保持原样
    counts = {}
    for k in keys:
        counts[k] = counts.get(k, 0) + 1
    out, seen = [], {}
    for k in keys:
        if not k:
            out.append("")
            continue
        if counts[k] == 1:
            out.append(k)
        else:
            seen[k] = seen.get(k, 0) + 1
            out.append(f"{k}({seen[k]})")
    # 空键 → 未命名n（数据安全优先，有数据即保留）
    unnamed = 0
    for i, k in enumerate(out):
        if k == "":
            unnamed += 1
            out[i] = f"未命名{unnamed}"
    return out


def out_of_bounds_merge_warning(ranges, header_row, height):
    """合并范围延伸出表头区（拖拽失手覆盖数据行）→ 警告。"""
    for mg in ranges:
        if mg.min_row <= header_row + height - 1 and mg.max_row > header_row + height - 1:
            return "检测到延伸进数据区的表头合并"
    return None
