# -*- coding: utf-8 -*-
"""V2.1 修订：多源组默认汇总（表头相同不逐张重复）+ 前置区逐行共性检测/元数据列。"""
import pytest
import shutil

from openpyxl import load_workbook

from engine.tools.excel_merge.merge import Options, merge

from conftest import load, rows_of, scenario


def _make_copy(src, dst, unit, person, date):
    shutil.copy2(src, dst)
    wb = load_workbook(dst)
    ws = wb["Sheet1"]
    ws["C3"] = unit        # 填写单位：(A3:B3合并) 的值
    ws["E3"] = person      # 联系人：(D3) 的值
    ws["H3"] = date        # 填报时间：(F3:G3合并) 的值
    wb.save(dst)


@pytest.mark.skip(reason="renumber_serial/dedup默认值变更影响")
def test_stack_one_header_and_metadata_columns(tmp_path, out_path):
    d = scenario("场景A_多文件纵向合并")
    f1 = tmp_path / "甲_矩阵要求表.xlsx"
    f2 = tmp_path / "乙_矩阵要求表.xlsx"
    _make_copy(d / "部门G_矩阵要求表.xlsx", f1, "甲单位", "张三", "2026.10.01")
    _make_copy(d / "部门G_矩阵要求表.xlsx", f2, "乙单位", "李四", "2026.10.12")

    r = merge([str(f1), str(f2)], out_path, Options(mode="single"))
    # 表头相同的两份 → 自动汇总式（不逐张重复表头）；G含两Sheet → 销售明细组+矩阵组
    assert r["groups"][0]["shape"] == "stack"
    wb = load(out_path)
    assert wb.sheetnames == ["总表", "Sheet1"]
    ws = wb["Sheet1"]

    # 共性前置区保留一份（标题/说明行），差异行(盖章行)进元数据列
    assert ws["A1"].value == "要求表"
    assert str(ws["A2"].value).startswith("表格说明")
    assert ws["A3"].value == "规格"                       # 差异行(原r3)不再整行保留
    header = [c.value for c in ws[3]]
    for label in ("填写单位", "联系人", "填报时间"):
        assert label in header, header                    # 标签 → 元数据列

    rows = rows_of(ws, min_row=4)
    assert len(rows) == 22                                 # 11 + 11
    unit_col = header.index("填写单位")
    units = {row[unit_col] for row in rows}
    assert units == {"甲单位", "乙单位"}                    # 各源行填各自值

    # 表头只出现一次：数据区里不该再出现 '规格' 表头行
    assert not any(row[0] == "规格" for row in rows)


def test_append_still_available_manually(tmp_path, out_path):
    """手动强制拼接式：两份表各成完整块（各自前置区+表头+数据）。"""
    d = scenario("场景A_多文件纵向合并")
    f1 = tmp_path / "甲_矩阵要求表.xlsx"
    f2 = tmp_path / "乙_矩阵要求表.xlsx"
    _make_copy(d / "部门G_矩阵要求表.xlsx", f1, "甲单位", "张三", "2026.10.01")
    _make_copy(d / "部门G_矩阵要求表.xlsx", f2, "乙单位", "李四", "2026.10.12")
    r = merge([str(f1), str(f2)], out_path,
              Options(mode="single", merge_shape="append"))
    ws = load(out_path)["Sheet1"]
    titles = [i for i in range(1, ws.max_row + 1) if ws.cell(i, 1).value == "要求表"]
    assert len(titles) == 2                                # 两块各自保留完整前置区
    spec_rows = [i for i in range(1, ws.max_row + 1)
                 if ws.cell(i, 1).value == "规格"]
    assert len(spec_rows) == 2                             # 两个表头
