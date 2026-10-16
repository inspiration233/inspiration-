# -*- coding: utf-8 -*-
"""复杂统计表适应性 + 数据区合并还原（2026-10-15 需求方实测反馈）。"""
from engine.tools.excel_merge.merge import Options, merge, scan

from conftest import load, rows_of, scenario

F = scenario("场景A_多文件纵向合并") / "部门F_复杂统计表.xlsx"   # 手工维护夹具(用户真实样例)


def test_complex_table_detection():
    r = scan([str(F)])
    det = next(d for d in r["header_detections"] if d["sheet"] == "Sheet2")
    # 三层前置区(大标题/统计说明/盖章联系人行)全部跳过, 表头定位第4行
    assert det["header_row"] == 4
    assert det["height"] == 2                                  # 双行套娃表头
    assert len(det["columns"]) == 19
    assert det["columns"][0] == "序号"
    assert "数量x/单位/部门" in det["columns"]
    assert "详细信息/年份" in det["columns"]
    assert "详细信息/补充说明" in det["columns"]


def test_data_area_merge_restoration(out_path):
    r = merge([str(F)], out_path, Options(mode="workbook", style="preserve"))
    assert r["totals"]["data_merges_preserved"] >= 19
    ws = load(out_path)["Sheet2"]
    assert ws.max_row == 18                                    # 2表头 + 16数据
    merged = {str(x) for x in ws.merged_cells.ranges}
    # 表头套娃结构
    for ref in ("A1:A2", "F1:G1", "H1:S1"):
        assert ref in merged, ref
    # 数据区合并: 填报单位整块 / 分组列 / 年份两行一组
    for ref in ("B3:B18", "C3:C18", "F3:F6", "F7:F11", "I3:I4", "I5:I6"):
        assert ref in merged, ref
    # V2.0: 公式保留并平移(数据行6→3, 位移-3)
    assert str(ws["C3"].value).startswith("=") and "G7" in str(ws["C3"].value)
    # 列宽复制
    assert (ws.column_dimensions["H"].width or 0) > 0


def test_unified_mode_keeps_clean_data(out_path):
    """统一样式 + 纵向合并模式不做数据区合并还原（干净数据原则）。
    注意：工作簿模式是页面拷贝语义，即使统一样式也还原数据区合并。"""
    r = merge([str(F)], out_path, Options(mode="single", style="unified"))
    assert r["totals"].get("data_merges_preserved", 0) == 0
    ws = load(out_path).active
    merged = {str(x) for x in ws.merged_cells.ranges}
    assert "B3:B18" not in merged                              # 仅表头区合并
