# -*- coding: utf-8 -*-
"""列对齐：严格（键集合必须一致）/ 宽容（按内部键对齐，多收缺空）。需求 §3.2.3。

匹配规则（纯结构、无语义）：
1. 精确匹配：内部键完全相等；
2. 末段回退：多行表头的复合键（父/子）与平铺键互相对齐——
   双方未精确匹配的键中，复合键的末段（"品名" of "商品信息/品名"）与
   平铺键相等且两侧均唯一时配对（需求 §3.2.2 输出呈现规则 3）。
"""


class AlignError(Exception):
    """严格模式表头不一致。携带文件与列级差异信息。"""

    def __init__(self, display, missing, extra):
        self.display = display
        self.missing = missing
        self.extra = extra
        parts = [f"「{display}」表头与模板不一致"]
        if missing:
            parts.append("缺少: " + "、".join(missing))
        if extra:
            parts.append("多出: " + "、".join(extra))
        super().__init__("；".join(parts))


def pair_columns(template_keys, source_keys):
    """源列 → 模板列 的下标映射（None = 源的新列）。"""
    t_exact = {}
    for i, k in enumerate(template_keys):
        t_exact.setdefault(k, i)
    s_exact = {}
    for j, k in enumerate(source_keys):
        s_exact.setdefault(k, j)

    pair = {}
    used_t = set()
    # 1) 精确
    for j, k in enumerate(source_keys):
        if k in t_exact and t_exact[k] not in used_t:
            pair[j] = t_exact[k]
            used_t.add(t_exact[k])

    # 2) 末段回退（双方未匹配的键）
    t_seg = {}
    for i, k in enumerate(template_keys):
        if i in used_t:
            continue
        t_seg.setdefault(k.rsplit("/", 1)[-1], []).append(i)
    s_free = [j for j in range(len(source_keys)) if j not in pair]
    s_seg = {}
    for j in s_free:
        s_seg.setdefault(source_keys[j].rsplit("/", 1)[-1], []).append(j)
    for seg, t_list in t_seg.items():
        s_list = s_seg.get(seg, [])
        if len(t_list) == 1 and len(s_list) == 1:
            pair[s_list[0]] = t_list[0]
            used_t.add(t_list[0])
    return pair


def align_sources(sources, mode, template_idx=0):
    """返回 (输出列键列表, [每源映射])。映射 = 输出列下标 → 该源列下标（None=该源缺此列）。"""
    template = sources[template_idx]
    tkeys = list(template.structure.keys)

    pairs = [pair_columns(tkeys, s.structure.keys) for s in sources]

    if mode == "strict":
        for s, pr in zip(sources, pairs):
            missing = [tkeys[i] for i in range(len(tkeys)) if i not in pr.values()]
            extra = [s.structure.keys[j] for j in range(len(s.structure.keys))
                     if j not in pr]
            if missing or extra:
                raise AlignError(s.display, missing, extra)
        return tkeys, [pr_inverse(pr, len(tkeys)) for pr in pairs]

    # 宽容：模板键序在前，未配对的源键按文件顺序追加
    out_keys = list(tkeys)
    for s, pr in zip(sources, pairs):
        for j, k in enumerate(s.structure.keys):
            if j not in pr and k not in out_keys:
                out_keys.append(k)
    # 输出列 → 源列
    maps = []
    for s, pr in zip(sources, pairs):
        pos_by_tkey = {tkeys[ti]: j for j, ti in pr.items()}
        extra_pos = {}
        for j, k in enumerate(s.structure.keys):
            if j not in pr:
                extra_pos[k] = j
        row = []
        for k in out_keys:
            if k in pos_by_tkey:
                row.append(pos_by_tkey[k])
            elif k in extra_pos:
                row.append(extra_pos[k])
            else:
                row.append(None)
        maps.append(row)
    return out_keys, maps


def pr_inverse(pair, tlen):
    """源→模板 映射翻转为 模板列 → 源列（未配对为 None）。"""
    inv = [None] * tlen
    for j, ti in pair.items():
        inv[ti] = j
    return inv
