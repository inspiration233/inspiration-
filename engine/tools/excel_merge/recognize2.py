# -*- coding: utf-8 -*-
"""识别引擎 v2：全局候选评分（D 方案）。

对扫描窗口内每个候选切分 (表头行 r, 表头高 h) 计算全局解释力分，取最优：

  标签性(表头区文本密度/唯一性, 复用 v1 单行打分)
  + 交界突变(表头底线两侧列类型跳变 —— 表头全文本标签, 数据出现数值/日期)
  + 数据同构性(数据前几行彼此类型相似, 且与表头区异构)
  + 合并证据(表头区合并覆盖密度高于数据区)
  + 样式证据(冻结窗格/自动筛选/加粗/填充 —— 用户亲手留下的答案)

置信度 = 最优与次优的分差；次优候选生成人话提示（"表头可能在第 r 或第 s 行"）。
方向判定 = 原窗口与转置窗口各跑一遍评分, 比全局最优总分。

v1（逐行打分+一票规则）保留在 recognize.py，由 reader.RECOG_ENGINE 开关切换。
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field

from .recognize import (
    HEADER_MAX_HEIGHT,
    is_numberlike,
    normalize_name,
    score_header_row,
)

MAX_HEADER_ROW_CANDIDATES = 10      # 表头行候选上限: 真实表表头几乎总在前10行内
DATA_COMPARE_ROWS = 8               # 数据同构性对比行数
CONFIDENT_MARGIN = 0.08             # 最优与次优分差 ≥ 此值 → 高置信

# 权重表集中可解释；调参只改这里
WEIGHTS = {
    "label": 0.24,          # 标签性
    "boundary": 0.26,       # 交界突变
    "homogeneity": 0.20,    # 数据同构性
    "heterogeneity": 0.08,  # 表头/数据异构
    "merge": 0.10,          # 合并证据
    "style": 0.12,          # 样式证据(基础)
}
STYLE_BONUS_FREEZE = 0.20           # 冻结窗格精确对齐候选
STYLE_BONUS_FILTER = 0.20           # 自动筛选起始行对齐候选
STYLE_BONUS_BOLD = 0.10             # 表头区加粗/填充惯例


@dataclass
class StyleFacts:
    """样式层事实（openpyxl 读取, 由 reader 提供；转置视图共用转置后的值）。"""
    freeze_row: int = 0             # 冻结窗格: 数据首行行号(如 "A3"→3)；0=无
    filter_row: int = 0             # 自动筛选区域起始行；0=无
    bold_rows: tuple = ()           # 加粗格占比>60% 的行号(1-based)
    filled_rows: tuple = ()         # 有底色填充占比>60% 的行号

    def transposed(self):
        return StyleFacts()          # 行/列语义互换后样式事实不可靠, 转置视图不消费


@dataclass
class Guess:
    header_row: int                 # 1-based；0 = 无有效候选
    height: int
    score: float
    runner_up: tuple = None         # (行, 分) 次优候选(不同表头行)
    confidence: bool = False
    evidence: dict = field(default_factory=dict)

    def runner_up_hint(self):
        if not self.runner_up or self.runner_up[0] == self.header_row:
            return None
        return (f"表头位置在第 {self.runner_up[0]} 行与第 {self.header_row} 行之间难以完全确定，"
                f"已按第 {self.header_row} 行处理，请核对。")


# ---------------------------------------------------------------- 类型向量

def _kind(v):
    if v in (None, ""):
        return None
    if isinstance(v, bool):
        return "num"
    if isinstance(v, (_dt.datetime, _dt.date)):
        return "date"
    return "num" if is_numberlike(v) else "text"


def _row_kinds(cells):
    return [_kind(v) for v in cells]


def _type_distance(kinds_a, kinds_b):
    """两行类型距离 ∈ [0,1]：共同非空列中类型不同的比例。无共同列 → None。"""
    same = diff = 0
    for a, b in zip(kinds_a, kinds_b):
        if a is None or b is None:
            continue
        if a == b:
            same += 1
        else:
            diff += 1
    total = same + diff
    return diff / total if total else None


def _row_numeric_ratio(cells):
    ne = [v for v in cells if v not in (None, "")]
    if not ne:
        return 0.0
    return sum(1 for v in ne if _kind(v) in ("num", "date")) / len(ne)


def _is_title_row(row, r, ranges, width, sparse):
    """标题行: 非空格≤sparse, 且(无合并 或 非空格被整行横向大合并覆盖)。"""
    ne = [i + 1 for i, v in enumerate(row) if v not in (None, "")]
    if len(ne) > sparse:
        return False
    if not ranges:
        return True
    for mg in ranges:
        if mg.min_row == r and mg.horizontal()                 and (mg.max_col - mg.min_col + 1) >= max(3, width * 0.6)                 and mg.min_col <= (ne[0] if ne else 1) <= mg.max_col:
            return True               # 整行大横合并盖住唯一内容 → 标题
    return all(not mg.covers(r, c) for mg in ranges for c in (ne or [1]))


def _merge_coverage(ranges, r, width):
    """第 r 行被任意合并覆盖的格数占比。"""
    if not ranges or width <= 0:
        return 0.0
    covered = sum(1 for c in range(1, width + 1)
                  if any(mg.covers(r, c) for mg in ranges))
    return covered / width


# ---------------------------------------------------------------- 候选评分

def score_candidate(window, r, h, ranges=None, style=None):
    """候选切分 (表头行 r, 高 h) 的全局解释力分。window: 1-based 行值列表。"""
    ranges = ranges or []
    style = style or StyleFacts()
    width = max((len(row) for row in window), default=0)
    if width == 0 or r < 1 or r + h - 1 > len(window):
        return None, {}

    header_rows = window[r - 1: r + h - 1]
    data_rows = window[r + h - 1: r + h - 1 + DATA_COMPARE_ROWS]
    if not any(any(v not in (None, "") for v in row) for row in data_rows):
        return None, {}              # 候选下方没有任何数据 → 不是完整切分
    # 标题行(稀疏+整行大横合并/无合并)属于前置区, 不得混入表头区 ——
    # 从候选层面判死而非罚分: "多包一行"在 het/boundary 上几乎免费, 罚不住
    sparse0 = max(1, min(2, width // 3))
    if any(_is_title_row(row, rr, ranges, width, sparse0)
           for rr, row in zip(range(r, r + h), header_rows)):
        return None, {}

    ev = {}

    # 1) 标签性: v1 单行分均值; "数字为主"行(数据行误入)与极稀疏行(标题)都重罚
    label = sum(score_header_row(row, width) for row in header_rows) / h
    numeric_penalty = max(_row_numeric_ratio(row) for row in header_rows)
    ev["label"] = round(label, 3)
    if numeric_penalty > 0.5:
        label -= numeric_penalty * 1.5
        ev["numeric_penalty"] = round(numeric_penalty, 2)
    # 表头区里的稀疏行(标题样)要罚: 标题的特征 = 非空格极少, 且要么没合并,
    # 要么被一整条横向大合并(≥60%宽度, 如 A1:AK1 的"全年计划")盖住 ——
    # 父级表头的合并是局部列分组(几列一组), 不是整行大横合并
    sparse = max(1, min(2, width // 3))
    title_like = sum(
        1 for rr, row in zip(range(r, r + h), header_rows)
        if _is_title_row(row, rr, ranges, width, sparse))
    if title_like:
        label -= 0.35 * title_like / h
        ev["sparse_header"] = title_like

    # 2) 交界突变: 表头底线 vs 数据首行; 共同列不足六成 → 证据打折防误判
    boundary = _type_distance(_row_kinds(header_rows[-1]), _row_kinds(data_rows[0]))
    header_ne = len([v for v in header_rows[-1] if v not in (None, "")])
    shared = sum(1 for a, b in zip(_row_kinds(header_rows[-1]), _row_kinds(data_rows[0]))
                 if a is not None and b is not None)
    if boundary is None:             # 无共同列: 表头宽、数据窄 → 弱负
        boundary = -0.2
    elif shared < max(2, int(min(header_ne, width) * 0.6)):
        boundary *= 0.5              # 共同列太少, 突变证据不可靠
        ev["boundary_thin"] = shared
    ev["boundary"] = round(boundary, 3)

    # 3) 数据同构性(抗脏行): 与"众数类型向量"的列级匹配率, 剔除最差 25% 离群行
    #    (合计/孤值行在真实表里合法存在, 两两均值会被它们摧毁; 众数+截尾抗污染)
    d_kinds = [_row_kinds(row) for row in data_rows
               if any(v not in (None, "") for v in row)]
    if len(d_kinds) >= 3:
        n_cols = max(len(k) for k in d_kinds)
        mode_vec = []
        for c in range(n_cols):
            counts = {}
            for k in d_kinds:
                if c < len(k) and k[c] is not None:
                    counts[k[c]] = counts.get(k[c], 0) + 1
            mode_vec.append(max(counts, key=counts.get) if counts else None)
        fit = []
        for k in d_kinds:
            same = diff = 0
            for a, b in zip(k, mode_vec):
                if a is None or b is None:
                    continue
                if a == b:
                    same += 1
                else:
                    diff += 1
            if same + diff:
                fit.append(same / (same + diff))
        fit.sort(reverse=True)
        keep = max(3, len(fit) - len(fit) // 4)
        homogeneity = sum(fit[:keep]) / keep
    else:
        homogeneity = 0.0
    ev["homogeneity"] = round(homogeneity, 3)

    # 4) 表头/数据异构: 表头区与数据区的代表性类型差异
    het_pairs = [_type_distance(_row_kinds(hr), _row_kinds(dr))
                 for hr in header_rows for dr in data_rows]
    het_pairs = [d for d in het_pairs if d is not None]
    heterogeneity = sum(het_pairs) / len(het_pairs) if het_pairs else 0.0
    ev["heterogeneity"] = round(heterogeneity, 3)

    # 5) 合并证据: 表头区平均合并覆盖 - 数据区平均(父子结构在表头侧)
    header_cov = sum(_merge_coverage(ranges, rr, width)
                     for rr in range(r, r + h)) / h
    data_cov = sum(_merge_coverage(ranges, rr, width)
                   for rr in range(r + h, r + h + DATA_COMPARE_ROWS)) / DATA_COMPARE_ROWS
    ev["merge"] = round(max(0.0, header_cov - data_cov), 3)
    merge_ev = min(1.0, max(0.0, header_cov - data_cov) * 2)

    # 6) 样式证据: 用户亲手留下的答案
    style_ev = 0.0
    if style.freeze_row and style.freeze_row == r + h:
        ev["freeze"] = True
    if style.filter_row and style.filter_row == r:
        ev["filter"] = True
    styled = set(style.bold_rows) | set(style.filled_rows)
    styled_in_window = {rr for rr in styled if 1 <= rr <= len(window)}
    if styled_in_window and data_rows and all(
            rr not in styled for rr in range(r + h, r + h + min(3, len(data_rows))))             and styled_in_window <= set(range(r, r + h)):
        ev["styled_header"] = True          # 表头区恰好覆盖全部样式行, 区外干净
    style_ev = (STYLE_BONUS_FREEZE if ev.get("freeze") else 0) \
        + (STYLE_BONUS_FILTER if ev.get("filter") else 0) \
        + (STYLE_BONUS_BOLD if ev.get("styled_header") else 0)

    # 宽表头强纵向信号: 表头行含大量互异文本标签(几十个字段名) ——
    # 宽扁表(字段远多于记录, 如拆分产物)在纵向解释下天然成立
    last_header_ne = [v for v in header_rows[-1] if v not in (None, "")]
    uniq_labels = len({normalize_name(v) for v in last_header_ne
                       if not is_numberlike(v)})
    if uniq_labels >= 8 and uniq_labels >= 0.5 * max(1, len(last_header_ne)):
        ev["wide_labels"] = uniq_labels

    raw = (WEIGHTS["label"] * max(0.0, label)
           + WEIGHTS["boundary"] * boundary
           + WEIGHTS["homogeneity"] * homogeneity
           + WEIGHTS["heterogeneity"] * heterogeneity
           + WEIGHTS["merge"] * merge_ev
           + WEIGHTS["style"] * style_ev * 3       # 样式 bonus 加权放大
           + (0.06 if ev.get("wide_labels") else 0.0)
           - 0.02 * (h - 1))                       # 表头高度先验(弱): 只防无据多包, 不压合并证据支撑的双行
    # 证据量折扣: 数据区非空行越少, 一切证据越不可信 —— 脏表里的孤立行
    # 组合会制造"假交界"(单行 boundary 极端), 样本不足的候选整体降信念
    nonempty_data = sum(1 for row in data_rows
                        if any(v not in (None, "") for v in row))
    discount = 0.7 + 0.3 * min(1.0, nonempty_data / 4)
    ev["data_rows"] = nonempty_data
    return raw * discount, ev


def detect_structure(window, ranges=None, style=None,
                     max_height=HEADER_MAX_HEIGHT):
    """全局最优切分：每个表头行取其最高分候选，再排行取最优与次优(不同行)。"""
    by_row = {}
    limit = min(len(window), MAX_HEADER_ROW_CANDIDATES)
    for r in range(1, limit + 1):
        for h in range(1, max_height + 1):
            if r + h > len(window):
                break
            score, ev = score_candidate(window, r, h, ranges, style)
            if score is None:
                continue
            if r not in by_row or score > by_row[r][1]:
                by_row[r] = (h, score, ev)
    if not by_row:
        return Guess(0, 0, 0.0)
    ranked = sorted(by_row.items(), key=lambda kv: -kv[1][1])
    (r1, (h1, s1, ev1)) = ranked[0]
    runner = (ranked[1][0], round(ranked[1][1][1], 3)) if len(ranked) > 1 else None
    guess = Guess(r1, h1, s1, runner, evidence=ev1)
    # 冻结/筛选是用户亲手指定的表头位置 → 命中即高置信
    guess.confidence = ((not runner or (s1 - runner[1]) >= CONFIDENT_MARGIN)
                        or bool(ev1.get("freeze") or ev1.get("filter")))
    if not guess.confidence:
        guess.evidence["low_confidence"] = True
    return guess


def detect_orientation(window, ranges=None, style=None):
    """方向 v2：原窗口与转置窗口各取全局最优分，对比。

    样式事实只属于原方向（冻结行/加粗行是行列语义的）；转置视图不带样式，
    防止同一证据两头计票。原方向领先不足 CONFIDENT_MARGIN → 低置信交人工。
    """
    from .recognize import transpose_window
    if not window:
        return "vertical", True
    width = max((len(r) for r in window), default=0)
    nonempty = sum(1 for r in window if any(v not in (None, "") for v in r))
    if width < 3 or nonempty < 3:
        return "vertical", True

    # 数据样本太少的视图(如拆分产物每组一两行)评分不可信 —— 不足即弃比
    def _data_rows(rows, guess):
        return max(0, len(rows) - (guess.header_row + guess.height - 1))             if guess and guess.header_row else 0

    g_v = detect_structure(window, ranges, style)
    t = transpose_window(window)
    t_ranges = [mg.transposed() for mg in (ranges or [])]
    g_h = detect_structure(t, t_ranges, None)
    v_data = _data_rows(window, g_v)
    h_data = _data_rows(t, g_h)
    if v_data < 4 and h_data < 4:
        return "vertical", True         # 双侧样本都枯竭: 小表, 默认纵向即可
    if g_h.header_row == 0 or h_data < 4:
        return "vertical", v_data >= 4
    if g_v.header_row == 0 or v_data < 4:
        return "horizontal", h_data >= 4
    # 宽标签方向偏向: 表头行含大量互异字段名(几十个) 是该方向的铁证 ——
    # 纵向表的字段名排成一行, 横向配置表的字段名排成一列(转置后成行)
    v_wide = 0.15 if g_v.evidence.get("wide_labels") else 0.0
    h_wide = 0.15 if g_h.evidence.get("wide_labels") else 0.0
    margin = 2 * CONFIDENT_MARGIN       # 方向比表头更怕错: 差距要翻倍才定向
    if g_v.score + v_wide - g_h.score - h_wide >= margin:
        return "vertical", True
    if g_h.score + h_wide - g_v.score - v_wide >= margin:
        return "horizontal", True
    return "vertical", False


def flatten_top_header_keys(window, guess, ranges):
    """给 preview/识别结果用的表头行说明(与 v1 flatten_keys 等价路径由 reader 处理)。"""
    return None
