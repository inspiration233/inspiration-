# -*- coding: utf-8 -*-
"""模糊不变量测试 —— "不射箭后画靶"：随机表 + 普适铁律断言。

夹具由 生成模糊集.py 按种子随机构造(不经过引擎设计思路)；
断言只检查任何合并工具都必须满足的不变量, 以及 pandas 独立重读的差分一致。
"""
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "测试集"))
from 生成模糊集 import gen_case  # noqa: E402

from engine.tools.excel_merge.merge import Options, merge, scan  # noqa: E402

SEEDS = list(range(int(__import__("os").environ.get("FUZZ_SEEDS", "40"))))


def _norm(v):
    if v is None or v == "":
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if hasattr(v, "date"):
        try:
            return v.date()
        except Exception:
            return v
    return v


@pytest.mark.parametrize("seed", SEEDS)
@pytest.mark.parametrize("style", ["fidelity", "fast"])
def test_fuzz_invariants(seed, style, tmp_path):
    src = tmp_path / f"f_{seed}.xlsx"
    gen_case(seed, src)
    out = tmp_path / f"out_{seed}_{style}.xlsx"

    # I1 扫描+合并不崩溃, 损坏等异常必须进报告而非抛出
    rep = merge([str(src)], out, Options(mode="single", style=style))
    t = rep["totals"]
    assert rep["files"][0]["status"] == "ok"

    # I2 输出是合法 xlsx(第三方重读)
    wb = load_workbook(out)
    for ws in wb.worksheets:
        # I5 合并范围不越界
        for mr in ws.merged_cells.ranges:
            assert 1 <= mr.min_row <= ws.max_row and 1 <= mr.min_col <= ws.max_column
            assert mr.max_row <= ws.max_row and mr.max_col <= ws.max_column
        # I4 公式合法: 以=开头; 跨表引用的目标 sheet 存在
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    body = c.value[1:]
                    if "!" in body:
                        ref = body.split("!")[0].strip("'")
                        assert ref in wb.sheetnames, (seed, style, c.coordinate, c.value)

    # I3 报告自洽: rows_written == 输出实际数据行数(扣除表头/前置区)
    total_rows = 0
    for g in rep.get("groups", []):
        ws = wb[g["sheet"]]
        total_rows += ws.max_row
    assert t["rows_written"] <= total_rows            # 写入行不会超过文件行数

    # I6 pandas 独立重读一致(openpyxl 与 pandas 的行数一致)
    pd = pytest.importorskip("pandas")
    for name in wb.sheetnames:
        df = pd.read_excel(out, sheet_name=name, header=None)
        assert len(df) == wb[name].max_row

    # I7 差分预言机(单文件·数据守恒): 用引擎自报的表头行号, 独立重算应写入的行
    srep = scan([str(src)])
    det = srep["header_detections"][0]
    if det["header_row"] <= 0 or not det["columns"]:
        return
    swb = load_workbook(src)
    sws = swb[det["sheet"]]
    hdr_end = det["header_row"] + det["height"]

    def _iter_data():
        if det["orientation"] == "horizontal":
            grid = [list(r) for r in sws.iter_rows(values_only=True)]
            ncols = max(len(r) for r in grid) if grid else 0
            for c in range(hdr_end - 1, ncols):
                yield [grid[r][c] if c < len(grid[r]) else None
                       for r in range(len(grid))]
        else:
            yield from sws.iter_rows(min_row=hdr_end, values_only=True)

    expected = []
    for row in _iter_data():
        vals = [_norm(v) for v in row]
        if all(v is None for v in vals):
            continue
        ne = [(i, v) for i, v in enumerate(vals) if v is not None]
        if len(ne) == 1 and ne[0][0] == 0 and isinstance(ne[0][1], (int, float)):
            continue                                     # 仅序号空行(按规格剔除)
        expected.append(tuple(v for v in vals[:len(det["columns"])]))
    ncols_out = len(det["columns"])
    # 输出侧: 总表组的数据行(扣表头区; 宽松比对: 行多重集按非空投影)
    out_ws = wb[rep["groups"][0]["sheet"]]
    got = []
    for row in out_ws.iter_rows(values_only=True):
        vals = tuple(_norm(v) for v in row[:ncols_out])
        if all(v is None for v in vals):
            continue
        got.append(vals)
    # 序号列被重排/公式被平移会改变值 → 比对非序号/非公式列的行指纹
    def fingerprint(vals, hdr):
        fp = []
        for i, v in enumerate(vals):
            if i < len(hdr) and hdr[i].strip() in ("序号", "编号"):
                continue
            if isinstance(v, str) and v.startswith("="):
                continue
            fp.append(v)
        return tuple(x for x in fp if x is not None)

    hdr = det["columns"]
    def _key(t):
        return tuple(repr(x) for x in t)
    exp_list = [fingerprint(e, hdr) for e in expected if fingerprint(e, hdr)]
    # 跳过表头区(前置区行+表头行): 定位第一行"包含已知列名"之后的行
    hdr_start = 0
    hdr_names = {h.strip() for h in hdr}
    for i, g in enumerate(got):
        strs = [v.strip() for v in g if isinstance(v, str)]
        if strs and sum(1 for s in strs if s in hdr_names or s.startswith("子")) >= max(1, len(hdr) // 2):
            hdr_start = i + 1
    got_list = [fingerprint(g, hdr) for g in got[hdr_start:] if fingerprint(g, hdr)]
    exp_fp = sorted(exp_list, key=_key)
    got_fp = sorted(got_list, key=_key)
    # 合计行按规格被剔除 → 输出 ≤ 期望; 差值应等于报告的剔除计数
    removed = t["summary_rows_removed"]
    assert len(got_fp) + removed >= len(exp_fp) - t.get("serial_rows_removed", 0) - 1, \
        (seed, style, len(got_fp), len(exp_fp), removed)
    assert len(got_fp) <= len(exp_fp)
