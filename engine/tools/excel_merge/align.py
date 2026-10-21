# -*- coding: utf-8 -*-
"""列对齐：严格（键集合必须一致）/ 宽容（按内部键对齐，多收缺空）。需求 §3.2.3。

匹配规则（纯结构、无语义）：
1. 精确匹配：内部键完全相等；
2. 末段回退：多行表头的复合键（父/子）与平铺键互相对齐——
   双方未精确匹配的键中，复合键的末段（"品名" of "商品信息/品名"）与
   平铺键相等且两侧均唯一时配对（需求 §3.2.2 输出呈现规则 3）。
"""
from collections import Counter
import json

from .recognize import normalize_name


def key_parts(key, *, signature=False):
    """Canonical hierarchy without treating a literal slash as a new level."""
    if isinstance(key, tuple):
        return key
    parsed = None
    if signature:
        try:
            parsed = json.loads(key)
        except (TypeError, ValueError):
            pass
    parts = (parsed if isinstance(parsed, list) and parsed and all(isinstance(p, str) for p in parsed)
             else [str(key)] if signature else str(key).split("/"))
    return tuple(normalize_name(part) for part in parts)


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


def pair_columns(template_keys, source_keys, allow_leaf=True):
    """源列 → 模板列 的下标映射（None = 源的新列）。"""
    template_keys = [key_parts(key) for key in template_keys]
    source_keys = [key_parts(key) for key in source_keys]
    t_exact = {}
    for i, k in enumerate(template_keys):
        t_exact.setdefault(k, []).append(i)
    s_exact = {}
    for j, k in enumerate(source_keys):
        s_exact.setdefault(k, []).append(j)

    pair = {}
    used_t = set()
    # 1) 精确
    for j, k in enumerate(source_keys):
        if k in t_exact and len(t_exact[k]) == len(s_exact[k]) == 1:
            pair[j] = t_exact[k][0]
            used_t.add(t_exact[k][0])

    if not allow_leaf:
        return pair
    # 2) 末段回退（双方未匹配的键）
    t_seg = {}
    for i, k in enumerate(template_keys):
        if i in used_t:
            continue
        t_seg.setdefault(k[-1], []).append(i)
    s_free = [j for j in range(len(source_keys)) if j not in pair]
    s_seg = {}
    for j in s_free:
        s_seg.setdefault(source_keys[j][-1], []).append(j)
    for seg, t_list in t_seg.items():
        s_list = s_seg.get(seg, [])
        if (len(t_list) == 1 and len(s_list) == 1
                and ((len(template_keys[t_list[0]]) > 1) != (len(source_keys[s_list[0]]) > 1))):
            pair[s_list[0]] = t_list[0]
            used_t.add(t_list[0])
    return pair


def align_sources(sources, mode, template_idx=0, allow_leaf=True):
    """返回 (输出列键列表, [每源映射])。映射 = 输出列下标 → 该源列下标（None=该源缺此列）。"""
    template = sources[template_idx]
    tkeys = list(template.structure.keys)
    def keys(source):
        signatures = getattr(source.structure, "key_signatures", None)
        return [key_parts(key, signature=signatures is not None)
                for key in (signatures if signatures is not None else source.structure.keys)]
    match_keys = keys(template)
    pairs = []
    for source in sources:
        skeys = keys(source)
        pair = ({i: i for i in range(len(skeys))} if source is template else
                pair_columns(match_keys, skeys, allow_leaf))
        explicit = getattr(source, "field_pairs", {})
        empty = getattr(source, "unmapped_targets", set())
        if (explicit or empty) and getattr(source, "field_map_template", template) is not template:
            raise ValueError(f"「{source.display}」字段对应依据的模板不在当前分组中，请为该分组重新选择模板")
        assigned = set(explicit.values()) | set(empty)
        pair = {si: ti for si, ti in pair.items() if si not in explicit and ti not in assigned}
        pair.update(explicit)
        scounts, tcounts = Counter(skeys), Counter(match_keys)
        ambiguous_sources = [i for i, key in enumerate(skeys) if scounts[key] > 1 and i not in pair]
        ambiguous_targets = [i for i, key in enumerate(match_keys)
                             if tcounts[key] > 1 and key in scounts and i not in pair.values() and i not in empty]
        if ambiguous_sources or ambiguous_targets:
            raise ValueError(f"「{source.display}」存在重名字段，请按位置明确对应所有重名字段，不能按出现顺序猜测")
        pairs.append(pair)

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
    extra_slots = {}
    for s, pr in zip(sources, pairs):
        for j, k in enumerate(keys(s)):
            if j not in pr and k not in extra_slots:
                extra_slots[k] = len(out_keys)
                out_keys.append(s.structure.keys[j])
    # 输出列 → 源列
    maps = []
    for s, pr in zip(sources, pairs):
        row = [None] * len(out_keys)
        for j, ti in pr.items():
            row[ti] = j
        for j, k in enumerate(keys(s)):
            if j not in pr:
                row[extra_slots[k]] = j
        maps.append(row)
    return out_keys, maps


def pr_inverse(pair, tlen):
    """源→模板 映射翻转为 模板列 → 源列（未配对为 None）。"""
    inv = [None] * tlen
    for j, ti in pair.items():
        inv[ti] = j
    return inv
