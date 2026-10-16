# -*- coding: utf-8 -*-
"""V2.0 原生保真：公式平移 / 数据验证搬运 / 行高 / 异构分组。"""
from engine.tools.excel_merge.merge import Options, merge

from conftest import load, rows_of, scenario


def test_dropdown_validation_preserved(out_path):
    """下拉完整保留：内联列表照搬；引用隐藏配置页的搬运到输出隐藏页并重接线。"""
    f = scenario("边缘情况") / "下拉列表.xlsx"
    r = merge([str(f)], out_path, Options(mode="single", style="fidelity"))
    t = r["totals"]
    assert t["validations_kept"] == 2

    wb = load(out_path)
    ws = wb["总表"]
    dvs = ws.data_validations.dataValidation
    formulas = {dv.formula1 for dv in dvs}
    assert '"日常,紧急,专项"' in formulas                    # 内联列表
    assert any(str(x).startswith("'_下拉配置'") for x in formulas)   # 重接线
    sqrefs = {str(dv.sqref) for dv in dvs}
    assert "C2:C9" in sqrefs and "B2:B9" in sqrefs

    assert "_下拉配置" in wb.sheetnames
    cfg = wb["_下拉配置"]
    assert cfg.sheet_state == "hidden"
    assert [cfg.cell(i, 1).value for i in (1, 2, 3)] == ["待审批", "已审批", "已驳回"]

    # 自定义行高还原
    assert ws.row_dimensions[2].height == 22


def test_formula_translation_complex_table(out_path):
    """部门F 复杂统计表：公式本体保留且引用按行位移平移。"""
    f = scenario("场景A_多文件纵向合并") / "部门F_复杂统计表.xlsx"
    r = merge([str(f)], out_path, Options(mode="workbook", style="fidelity"))
    ws = load(out_path)["Sheet2"]
    c3 = str(ws["C3"].value)
    assert c3.startswith("=")
    assert "G7" in c3 and "G3" in c3                       # 行位移 -3 (数据行6→3)
    assert r["totals"]["formulas_to_value"] == 0           # 块内引用全部平移成功


def test_heterogeneous_grouping(out_path):
    """异构自动分组：明细们进总表、统计表独立成页且结构完整。"""
    d = scenario("场景A_多文件纵向合并")
    r = merge([str(d / "部门A_销售明细.xlsx"), str(d / "部门F_复杂统计表.xlsx")],
              out_path, Options(mode="single", style="fidelity"))
    wb = load(out_path)
    assert wb.sheetnames == ["总表", "部门F_复杂统计表"]
    n_main = len(rows_of(wb["总表"]))
    assert n_main == 17 + 17                               # 部门A + 部门F/销售明细
    ws2 = wb["部门F_复杂统计表"]
    merged = {str(x) for x in ws2.merged_cells.ranges}
    assert "B3:B18" in merged and "H1:S1" in merged        # 套娃结构完整


def test_fast_mode_still_works(out_path):
    """高速模式（原统一样式）保持干净数据行为。"""
    d = scenario("场景A_多文件纵向合并")
    r = merge([str(d / "部门F_复杂统计表.xlsx")], out_path,
              Options(mode="workbook", style="fast"))
    ws = load(out_path)["Sheet2"]
    assert not str(ws["C3"].value).startswith("=")         # 公式=缓存值
    assert r["totals"]["formulas_translated"] == 0
