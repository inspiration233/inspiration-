# -*- coding: utf-8 -*-
"""泛用性验收测试集自检脚本 — 验证生成的文件完整性和结构正确性（不调用引擎代码）"""
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
PASS, FAIL, WARN = 0, 0, 0


def ok(msg):
    global PASS; PASS += 1
    print(f"  ✓ {msg}")


def bad(msg):
    global FAIL; FAIL += 1
    print(f"  ✗ {msg}")


def warn(msg):
    global WARN; WARN += 1
    print(f"  ⚠ {msg}")


def check_case(cid):
    d = ROOT / cid
    if not d.exists():
        bad(f"{cid}: 目录不存在")
        return
    cj = d / "case.json"
    if not cj.exists():
        bad(f"{cid}: 缺少 case.json")
        return
    data = json.loads(cj.read_text(encoding="utf-8"))
    xlsx_files = list(d.glob("*.xlsx"))
    if not xlsx_files:
        bad(f"{cid}: 没有 Excel 文件")
        return
    ok(f"{cid}: {len(xlsx_files)}个xlsx + case.json")

    # 逐文件检查：能否被 openpyxl 打开
    for f in xlsx_files:
        try:
            wb = load_workbook(f)
            ok(f"  {f.name}: {len(wb.sheetnames)}个Sheet")
        except Exception as e:
            bad(f"  {f.name}: 无法打开 — {e}")


def check_W1():
    """W1 完整流程数据一致性"""
    d = ROOT / "W1"
    if not (d / "W1_原始计划表.xlsx").exists():
        return bad("W1: 缺少原始文件")
    wb = load_workbook(d / "W1_原始计划表.xlsx")
    ws = wb.active
    # 检查表头在第4行
    if ws.cell(4, 1).value != "任务编号":
        return bad("W1: 表头不在第4行")
    # 检查任务数
    data_rows = sum(1 for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value)
    if data_rows != 12:
        bad(f"W1: 原始表任务数{data_rows}≠12")
    else:
        ok(f"W1: 原始表12个任务")
    # 检查公式
    if ws.cell(5, 6).value == "=E5*100":
        ok("W1: 预算公式=E*100")
    else:
        warn(f"W1: 预算公式={ws.cell(5, 6).value}")
    # 检查下拉
    dvs = ws.data_validations.dataValidation
    if any(dv.formula1 == '"进行中,已完成,已取消"' for dv in dvs):
        ok("W1: 状态下拉存在")
    # 检查拆分文件
    splits = list(d.glob("W1_拆分_*.xlsx"))
    if len(splits) == 5:
        ok(f"W1: 5个单位拆分文件")
        for f in splits:
            wbs = load_workbook(f)
            wss = wbs.active
            unit = wss.cell(2, 3).value  # 第2行第3列=责任单位
            all_same = all(wss.cell(r, 3).value == unit
                           for r in range(2, wss.max_row + 1) if wss.cell(r, 1).value)
            if all_same:
                ok(f"  拆分_{unit}: 所有行单位一致")
            else:
                bad(f"  拆分文件中混入其他单位数据")
    # 检查反馈变更
    fb = load_workbook(d / "W1_反馈_财务部.xlsx")
    qty = fb.active.cell(2, 5).value  # T002在第2行, 第5列=数量
    if qty == 25:
        ok("W1: 财务部反馈T002数量=25")
    else:
        bad(f"W1: 财务部反馈T002数量={qty}(期望25)")
    # 预期输出
    exp = load_workbook(d / "W1_预期_全年更新.xlsx")
    for r in range(2, exp.active.max_row + 1):
        if exp.active.cell(r, 1).value == "T002":
            if exp.active.cell(r, 5).value == 25:
                ok("W1: 全年更新版T002=25")
            else:
                bad("W1: 全年更新版T002≠25")
        if exp.active.cell(r, 1).value == "T999":
            ok("W1: 全年更新版含新增T999")
            break


def check_fidelity():
    """保真案例结构检查"""
    # F19: 数字格式
    d = ROOT / "F19"
    if (d / "F19_数字格式.xlsx").exists():
        wb = load_workbook(d / "F19_数字格式.xlsx")
        ws = wb.active
        if ws["B2"].number_format == "0.0%":
            ok("F19: 百分比格式")
        if ws["D2"].number_format == "@":
            ok("F19: 长编号文本格式")
        if ws["E2"].number_format == "000":
            ok("F19: 前导零格式")
    # F21: 打印设置
    d = ROOT / "F21"
    if (d / "F21_打印设置.xlsx").exists():
        wb = load_workbook(d / "F21_打印设置.xlsx")
        ws = wb.active
        if ws.freeze_panes == "B2":
            ok("F21: 冻结B2")
        if ws.page_setup.orientation == "landscape":
            ok("F21: 横向打印")
        if ws.print_title_rows == "1:1":
            ok("F21: 打印标题行")
    # F22: 批注和隐藏页
    d = ROOT / "F22"
    if (d / "F22_批注链接隐藏页.xlsx").exists():
        wb = load_workbook(d / "F22_批注链接隐藏页.xlsx")
        ws = wb["数据"]
        if ws["B2"].comment:
            ok("F22: 批注存在")
        if "配置" in wb.sheetnames and wb["配置"].sheet_state == "hidden":
            ok("F22: 隐藏配置页")
        dvs = ws.data_validations.dataValidation
        if any("配置" in str(dv.formula1) for dv in dvs):
            ok("F22: 下拉引用隐藏页")
    # F24: 隐藏行列
    d = ROOT / "F24"
    if (d / "F24_行高列宽隐藏.xlsx").exists():
        wb = load_workbook(d / "F24_行高列宽隐藏.xlsx")
        ws = wb.active
        if ws.column_dimensions["D"].hidden:
            ok("F24: D列隐藏")
        if ws.row_dimensions[8].hidden:
            ok("F24: r8隐藏")
        if ws.row_dimensions[2].height == 30:
            ok("F24: r2行高30")


def main():
    print("=" * 50)
    print("泛用性验收测试集自检")
    print("=" * 50)

    all_cases = sorted([d.name for d in ROOT.iterdir()
                        if d.is_dir() and d.name[0] in "WSF"])
    print(f"\n发现 {len(all_cases)} 个案例目录\n")

    for cid in all_cases:
        check_case(cid)

    print("\n--- 重点案例深度检查 ---\n")
    check_W1()
    check_fidelity()

    print(f"\n{'=' * 50}")
    print(f"结果: {PASS} 通过 / {FAIL} 失败 / {WARN} 警告")
    if FAIL > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
