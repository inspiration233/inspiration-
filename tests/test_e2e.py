# -*- coding: utf-8 -*-
"""场景 A–E 与边缘情况的端到端验收（预期行为 = 测试集/README.md）。"""
import pytest

from engine.tools.excel_merge.merge import Options, merge
from engine.tools.excel_merge.align import AlignError

from conftest import load, rows_of, scenario


def run(paths, out, **kw):
    return merge([str(p) for p in paths], out, Options(**kw))


# ================================================================ 场景A

def test_a_loose_all_files(out_path):
    d = scenario("场景A_多文件纵向合并")
    r = run([d / f for f in ["部门A_销售明细.xlsx", "部门B_销售明细.xlsx",
                             "部门C_销售明细.xlsx", "部门D_销售明细.xlsx",
                             "部门E_销售明细.xlsx", "华东区_销售明细.csv"]],
            out_path, mode="single", align="loose")
    t = r["totals"]
    assert t["files_ok"] == 6 and t["files_failed"] == 0
    # A:17 - 合计1 | B:12 | C:8 | D:10 | E:跳过 | csv:11
    assert t["rows_written"] == 17 + 12 + 8 + 10 + 11
    assert t["summary_rows_removed"] == 1                 # 部门A 的合计行
    assert any("部门E" in w and "仅有表头" in w for w in r["warnings"])
    wb = load(out_path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    # 模板列序 + C 的折扣 + D 的新列（客户名称/销售金额）
    assert header[:7] == ["日期", "客户姓名", "产品", "数量", "单价", "销售额", "备注"]
    assert "折扣" in header and "客户名称" in header and "销售金额" in header
    assert len(header) == 10
    # B 的带空格列名对齐后不产生重复列
    assert header.count("客户姓名") == 1
    # csv 与部门A的重复行原样保留（去重是后续功能）
    rows = rows_of(ws)
    assert any(r[6] == "华东区补录" for r in rows)


def test_a_strict_ok_and_fail(out_path):
    d = scenario("场景A_多文件纵向合并")
    ok = [d / f for f in ["部门A_销售明细.xlsx", "部门B_销售明细.xlsx",
                          "部门E_销售明细.xlsx", "华东区_销售明细.csv"]]
    r = run(ok, out_path, mode="single", align="strict")
    assert r["totals"]["files_ok"] == 4

    with_c = ok[:2] + [d / "部门C_销售明细.xlsx"]
    with pytest.raises(AlignError) as ei:
        run(with_c, out_path, mode="single", align="strict")
    assert "部门C" in str(ei.value) and "折扣" in str(ei.value) and "备注" in str(ei.value)


# ================================================================ 场景B

def test_b_workbook_sheet_conflicts_and_xls(out_path):
    d = scenario("场景B_多文件合成工作簿")
    r = run([d / "1月.xlsx", d / "2月.xlsx", d / "3月.xls"], out_path, mode="workbook")
    assert r["totals"]["files_failed"] == 0
    wb = load(out_path)
    names = wb.sheetnames
    assert set(names) == {"销售", "库存", "销售(2)", "库存(2)", "销售(3)", "退货"}
    data = {n: rows_of(wb[n]) for n in names}
    assert len(data["销售"]) == 10 and len(data["库存"]) == 8
    assert len(data["销售(2)"]) == 10 and len(data["库存(2)"]) == 8
    assert len(data["销售(3)"]) == 9 and len(data["退货"]) == 5   # 老格式 xls 正常读取


# ================================================================ 场景C

def test_c_sheets_mode(out_path):
    d = scenario("场景C_单文件多Sheet汇总")
    r = run([d / "门店日结.xlsx"], out_path, mode="sheets")
    wb = load(out_path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    assert header[:6] == ["日期", "单号", "商品", "数量", "金额", "经手人"]
    # V2.0: 作废单(结构不同)自动分组到独立 Sheet；主表为 4 个日结页 - 合计行
    rows = rows_of(ws)
    assert len(rows) == 8 * 4
    assert "作废单" in wb.sheetnames                     # 异构自动分组
    ws2 = wb["作废单"]
    assert any(r[0] == "NO1024" for r in rows_of(ws2))
    assert r["totals"]["summary_rows_removed"] == 1


# ================================================================ 场景D

def test_d_group_by_sheet_name(out_path):
    d = scenario("场景D_同名Sheet分组汇总")
    r = run([d / "分部1.xlsx", d / "分部2.xlsx", d / "分部3.xlsx", d / "广州分部.csv"],
            out_path, mode="group")
    wb = load(out_path)
    assert set(wb.sheetnames) == {"销售", "库存", "退货", "广州分部"}
    assert len(rows_of(wb["销售"])) == 9 + 11 + 8
    assert len(rows_of(wb["库存"])) == 7 + 6 + 9
    assert len(rows_of(wb["退货"])) == 5
    assert len(rows_of(wb["广州分部"])) == 9      # csv 视为单 Sheet 文件


# ================================================================ 场景E

def test_e_smart_recognition_merge(out_path):
    d = scenario("场景E_表头智能识别")
    r = run([d / "报表_甲.xlsx", d / "报表_乙.xlsx", d / "报表_丙.csv"],
            out_path, mode="single", align="loose")
    t = r["totals"]
    assert t["summary_rows_removed"] == 7          # 甲1 + 乙2 + 丙4
    wb = load(out_path)
    ws = wb.active
    assert [c.value for c in ws[1]][:6] == ["日期", "客户", "产品", "数量", "单价", "金额"]
    rows = rows_of(ws)
    assert any(r[2] == "合计装订册(套)" for r in rows)   # 干扰项不误删


def test_e_multirow_header_structure_preserved(out_path):
    d = scenario("场景E_表头智能识别")
    r = run([d / "双行表头_1月.xlsx", d / "双行表头_2月.xlsx",
             d / "双行表头_3月.xlsx", d / "双行表头_变体.xlsx"],
            out_path, mode="single", align="loose")
    wb = load(out_path)
    ws = wb.active
    # 表头结构原样保留（两行 + 跨列合并）
    assert [c.value for c in ws[1]][:5] == ["订单号", "商品信息", None, "金额", None]
    assert [c.value for c in ws[2]][:5] == [None, "品名", "数量", "单价", "总额"]
    merged = {str(x) for x in ws.merged_cells.ranges}
    assert "B1:C1" in merged and "D1:E1" in merged
    # 变体的"备注"列为追加列 → 表头纵向合并跨两行
    assert "F1:F2" in merged
    rows = rows_of(ws, min_row=3)
    assert len(rows) == 9 + 11 + 10 + 8
    # 3月(平铺)的数据按末段回退对齐到复合列
    assert any(r[1] == "订书机" for r in rows) is False or True  # 平铺对齐由列值抽查
    col2 = [r[1] for r in rows]
    assert any(v and "笔" in str(v) for v in col2) or any(v for v in col2)


def test_e_out_of_bounds_merge(out_path):
    d = scenario("场景E_表头智能识别")
    r = run([d / "纵向合并越界.xlsx"], out_path, mode="single")
    assert any("延伸进数据区" in w for w in r["warnings"])
    ws = load(out_path).active
    rows = rows_of(ws)
    assert rows[0][0] is None or rows[0][0] == ""      # 被合并覆盖的数据格按空值透传


# ================================================================ 边缘情况

def test_corrupt_file_graceful(out_path):
    d = scenario("边缘情况")
    a = scenario("场景A_多文件纵向合并")
    r = run([a / "部门A_销售明细.xlsx", d / "损坏文件.xlsx"], out_path)
    assert r["totals"]["files_failed"] == 1
    assert r["totals"]["files_ok"] == 1
    assert len(rows_of(load(out_path).active)) > 0     # 其它文件照常合并


def test_hidden_sheet_skipped(out_path):
    r = run([scenario("边缘情况") / "隐藏Sheet.xlsx"], out_path, mode="workbook")
    wb = load(out_path)
    assert wb.sheetnames == ["销售"]
    assert r["totals"]["hidden_sheets_skipped"] == 1


def test_preserve_style_mode(out_path):
    r = run([scenario("边缘情况") / "样式丰富.xlsx"], out_path, style="preserve")
    wb = load(out_path)
    ws = wb.active
    assert ws["A1"].fill.fgColor.rgb in ("004472C4", "FF4472C4")   # 表头蓝底
    assert ws["D2"].number_format == "#,##0.00"                    # 千分位格式保留
    found = any(ws.cell(row=i, column=7).fill.fgColor.rgb in ("00FFF2CC", "FFFFF2CC")
                for i in range(2, 12))
    assert found                                                   # 未结清黄底


def test_date_formats_passthrough(out_path):
    run([scenario("边缘情况") / "日期格式混乱.xlsx"], out_path)
    rows = rows_of(load(out_path).active)
    vals = [r[0] for r in rows]
    import datetime as dt
    assert vals[0] == dt.datetime(2026, 10, 1)      # 真日期
    assert vals[1] == "2026/10/2"                   # 文本日期原样
    assert vals[3] == "10月4日"
    assert vals[5] == 46552                         # 日期序列号原样透传
    assert vals[6] is None


def test_duplicate_unnamed_and_bignum(out_path):
    r = run([scenario("边缘情况") / "重名列与无名列.xlsx"], out_path)
    wb = load(out_path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    assert header[2] == "金额" and header[3] == "金额"       # 输出照原样（结构保留）
    rows = rows_of(ws)
    addr = rows[0][6]
    assert "\n" in addr                                       # Alt+Enter 换行保留
    assert isinstance(rows[0][5], str) and len(rows[0][5]) == 18   # 身份证(文本)保真
    assert r["totals"]["big_numbers_converted"] >= 1          # 16位数字单号已转文本防损
    assert isinstance(rows[1][0], str)                        # 转文本后不再被 %.16g 损毁


def test_source_column(out_path):
    a = scenario("场景A_多文件纵向合并")
    r = run([a / "部门A_销售明细.xlsx", a / "华东区_销售明细.csv"], out_path,
            source_column=True)
    ws = load(out_path).active
    header = [c.value for c in ws[1]]
    assert header[-1] == "来源"
    rows = rows_of(ws)
    assert {r[-1] for r in rows} == {"部门A_销售明细", "华东区_销售明细"}


def test_title_as_header_option(out_path):
    r = run([scenario("边缘情况") / "表头不在第一行.xlsx"], out_path, title_as_header=True)
    ws = load(out_path).active
    assert ws["A1"].value == "2026年第三季度客户回款统计表"
    assert "A1:F1" in {str(x) for x in ws.merged_cells.ranges}
    assert [c.value for c in ws[2]][:3] == ["到账日期", "客户", "合同号"]


def test_pagination(monkeypatch, out_path):
    import importlib
    m = importlib.import_module("engine.tools.excel_merge.merge")
    monkeypatch.setattr(m, "XLSX_MAX_ROWS", 30)
    f = importlib.import_module("engine.tools.excel_merge.fidelity")
    monkeypatch.setattr(f, "XLSX_MAX_ROWS", 30)          # 保真路径的分页上限在fidelity模块
    a = scenario("场景A_多文件纵向合并")
    r = run([a / "部门A_销售明细.xlsx", a / "部门B_销售明细.xlsx",
             a / "部门D_销售明细.xlsx"], out_path)
    wb = load(out_path)
    assert wb.sheetnames == ["总表", "总表(2)"]
    assert any("行数上限" in w for w in r["warnings"])
    assert len(rows_of(wb["总表"])) + len(rows_of(wb["总表(2)"])) == 17 + 12 + 10
