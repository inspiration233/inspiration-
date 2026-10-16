# -*- coding: utf-8 -*-
"""方向自动识别（横向表转正）+ 序号空行清理与重排。"""
from engine.tools.excel_merge.merge import Options, merge, scan

from conftest import load, rows_of, scenario


def test_horizontal_table_detection_and_merge(out_path):
    d = scenario("边缘情况")
    r = scan([str(d / "横向表.xlsx")])
    det = r["header_detections"][0]
    assert det["orientation"] == "horizontal"
    assert det["columns"] == ["客户", "联系人", "签订日期", "金额", "回款状态", "备注"]

    r = merge([str(d / "横向表.xlsx")], out_path, Options(mode="single"))
    ws = load(out_path).active
    assert [c.value for c in ws[1]] == ["客户", "联系人", "签订日期", "金额", "回款状态", "备注"]
    rows = rows_of(ws)
    assert len(rows) == 3                                    # 3 条记录(原3列)
    assert rows[0][0] == "上海岚图贸易有限公司"              # 记录 = 原列B
    import datetime as dt
    assert isinstance(rows[0][2], dt.datetime)               # 日期类型保真


def test_vertical_tables_not_misjudged():
    cases = [
        ("场景A_多文件纵向合并", "部门A_销售明细.xlsx"),
        ("场景E_表头智能识别", "双行表头_1月.xlsx"),
        ("场景A_多文件纵向合并", "部门F_复杂统计表.xlsx"),
    ]
    for d, f in cases:
        r = scan([str(scenario(d) / f)])
        assert all(x["orientation"] == "vertical" for x in r["header_detections"]), f


def test_serial_rows_cleanup_and_renumber(out_path):
    d = scenario("边缘情况")
    r = merge([str(d / "序号空行.xlsx")], out_path, Options(mode="single"))
    t = r["totals"]
    assert t["serial_rows_removed"] == 3                     # 3 行仅序号的空行
    assert t["summary_rows_removed"] == 1                    # 尾部合计
    rows = rows_of(load(out_path).active)
    assert len(rows) == 4
    assert [row[0] for row in rows] == [1, 2, 3, 4]          # 序号重排连续
    assert [row[1] for row in rows] == ["复印纸采购", "打印机维修", "团建费用", "快递费"]


def test_serial_options_off_keeps_original(out_path):
    d = scenario("边缘情况")
    r = merge([str(d / "序号空行.xlsx")], out_path,
              Options(mode="single", strip_serial_rows=False,
                      renumber_serial=False, remove_summary=False))
    rows = rows_of(load(out_path).active)
    assert len(rows) == 8                                    # 空行序号行/合计全保留
    assert rows[0][0] == 1 and rows[2][0] == 3               # 原始序号不动
