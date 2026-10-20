# -*- coding: utf-8 -*-
"""修订脚本：修正已确认的内容错误 + 补齐W2~W6完整流程 + 修正保真标准"""
import json, random, shutil, datetime as dt
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

ROOT = Path(__file__).resolve().parent
UNITS = ["综合办公室", "财务部", "工程技术部", "安全管理部", "市场营销部"]
TASKS_POOL = ["完成年度预算编制", "组织安全培训演练", "设备巡检维护", "季度报表汇总",
              "完成资料汇总", "成本合计核算", "供应商评估审查", "新员工入职办理",
              "项目验收资料整理", "应急预案修订", "办公耗材采购", "固定资产盘点"]
THIN = Border(*[Side(style="thin", color="999999")] * 4)
HDR_F = Font(bold=True, size=10, name="微软雅黑")
HDR_B = PatternFill("solid", fgColor="D9E1F2")

def _hdr(ws, row, vals, start_col=1):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font, c.fill, c.border = HDR_F, HDR_B, THIN
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

def _cj(d, data):
    (d / "case.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ============================================================ W1 修正
def fix_W1():
    d = ROOT / "W1"
    if not d.exists(): return

    # 修正1: 工程技术部反馈中 T999 公式（原引用E7错误，应为T999所在行自身的E列）
    fb = d / "W1_反馈_工程技术部.xlsx"
    if fb.exists():
        wb = load_workbook(fb)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value == "T999":
                ws.cell(r, 6).value = f"=E{r}*100"  # 引用自身行的E列
        _save(wb, fb)

    # 修正2: 全年更新版 T002 预算（25*100=2500或公式=E*100）
    exp = d / "W1_预期_全年更新.xlsx"
    if exp.exists():
        wb = load_workbook(exp)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value == "T002":
                ws.cell(r, 6).value = f"=E{r}*100"  # 公式而非硬编码
        _save(wb, exp)

    # 修正3: 拆分文件补齐前置区、下拉、条件格式、冻结、列宽（原生复刻）
    for f in d.glob("W1_拆分_*.xlsx"):
        wb = load_workbook(f)
        ws = wb.active
        # 补冻结
        ws.freeze_panes = "A2"
        # 补下拉
        dv = DataValidation(type="list", formula1='"进行中,已完成,已取消"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"G2:G{ws.max_row}")
        # 补条件格式
        ws.conditional_formatting.add(f"G2:G{ws.max_row}",
            CellIsRule(operator="equal", formula=['"已完成"'],
                       fill=PatternFill("solid", fgColor="C6EFCE")))
        # 补列宽
        for col, w in zip("ABCDEFG", [10, 22, 14, 8, 10, 14, 10]):
            ws.column_dimensions[col].width = w
        _save(wb, f)

    # 修正4: 预期汇总也补齐保真属性
    for fname in ["W1_预期_6月汇总.xlsx", "W1_预期_全年更新.xlsx"]:
        f = d / fname
        if f.exists():
            wb = load_workbook(f)
            ws = wb.active
            ws.freeze_panes = "A2"
            for col, w in zip("ABCDEFG", [10, 22, 14, 8, 10, 14, 10]):
                ws.column_dimensions[col].width = w
            _save(wb, f)

    # 修正5: 更新 case.json
    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["corrections"] = [
        "T999公式修正为引用自身行E列",
        "T002预算修正为公式=E*100(非硬编码2000)",
        "拆分文件补齐冻结/下拉/条件格式/列宽",
        "预期汇总补齐冻结/列宽",
    ]
    data["fidelity_requirements"] = {
        "前置区": "原始表3行前置区在拆分后应保留",
        "下拉": "状态下拉(进行中/已完成/已取消)在拆分和更新后保留",
        "条件格式": "已完成=绿色填充在拆分和更新后保留",
        "冻结": "冻结窗格在拆分和更新后保留",
        "列宽": "原始列宽在拆分和更新后保留",
        "公式": "预算金额=E*100公式语义不变",
    }
    _cj(d, data)

    # 重新生成工程技术部反馈（修正T999行号引用）
    # 找到T999在拆分文件中的行号并修正
    fb2 = d / "W1_反馈_工程技术部.xlsx"
    if fb2.exists():
        wb = load_workbook(fb2)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value == "T999":
                ws.cell(r, 6).value = f"=E{r}*100"
                # 补保真属性
                ws.freeze_panes = "A2"
                dv = DataValidation(type="list", formula1='"进行中,已完成,已取消"', allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(f"G2:G{ws.max_row}")
                for col, w in zip("ABCDEFG", [10, 22, 14, 8, 10, 14, 10]):
                    ws.column_dimensions[col].width = w
        _save(wb, fb2)

    print("✓ W1 修正完成")


# ============================================================ W3 修正（双行表头保留）
def fix_W3():
    d = ROOT / "W3"
    if not d.exists(): return

    for f in d.glob("W3_拆分_*.xlsx"):
        wb = load_workbook(f)
        ws = wb.active
        # 检查是否已有合并（如果有就跳过）
        if len(list(ws.merged_cells.ranges)) >= 2:
            continue
        # 重建为双行表头结构
        unit = ws.title
        rows_data = []
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value:
                rows_data.append([ws.cell(r, c).value for c in range(1, 9)])

        wb2 = Workbook(); ws2 = wb2.active; ws2.title = unit
        # 双行表头
        _hdr(ws2, 1, ["编号", "任务名称", "基本信息", "基本信息", "进度", "进度", "进度", "备注"])
        ws2.merge_cells("C1:D1")
        ws2.merge_cells("E1:G1")
        _hdr(ws2, 2, ["", "", "责任单位", "完成时限", "计划数", "完成数", "完成率", ""])
        for i, row in enumerate(rows_data):
            r = i + 3
            ws2.append(row)
            # 修正公式行号
            if ws2.cell(r, 7).value and "=" in str(ws2.cell(r, 7).value):
                ws2.cell(r, 7).value = f"=F{r}/E{r}"
        # 条件格式
        ws2.conditional_formatting.add(f"G3:G{ws2.max_row}",
            CellIsRule(operator="lessThan", formula=["0.5"],
                       fill=PatternFill("solid", fgColor="FFC7CE")))
        ws2.freeze_panes = "A3"
        for col, w in zip("ABCDEFGH", [10, 20, 14, 12, 10, 10, 10, 12]):
            ws2.column_dimensions[col].width = w
        _save(wb2, f)

    # 更新 case.json
    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["fidelity_requirements"] = {
        "双行表头": "拆分文件保留C1:D1和E1:G1合并结构",
        "条件格式": "完成率<0.5红色填充保留",
        "公式": "完成率公式=F/E行号随位置调整",
        "冻结": "冻结A3(双行表头下方)",
    }
    _cj(d, data)
    print("✓ W3 修正完成（双行表头保留）")


# ============================================================ W5 修正（真值核对）
def fix_W5():
    d = ROOT / "W5"
    if not d.exists(): return
    wb = load_workbook(d / "W5_原始_Sheet月份.xlsx")
    actual_counts = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        cnt = sum(1 for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value)
        actual_counts[sn] = cnt

    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["ground_truth"]["tasks_per_sheet"] = [actual_counts[sn] for sn in wb.sheetnames]
    data["ground_truth"]["total_tasks"] = sum(actual_counts.values())
    data["correction"] = f"真值修正: 实际各Sheet任务数为{actual_counts}"
    _cj(d, data)
    print(f"✓ W5 真值修正: {actual_counts}")


# ============================================================ W2~W6 补齐完整流程
def complete_W2():
    d = ROOT / "W2"
    if not d.exists(): return
    # 补反馈文件（财务部改了1月数量）
    wb = load_workbook(d / "W2_拆分_财务部.xlsx")
    ws = wb.active
    ws.cell(2, 3).value = 99  # 改1月数量
    _save(wb, d / "W2_反馈_财务部.xlsx")

    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["expected_changes"] = [{"task": "任务1", "field": "1月数量", "old": "原值", "new": 99}]
    data["expected_monthly"] = "1月汇总中财务部任务1数量=99"
    data["expected_annual"] = "全年合计公式包含新数量"
    data["business_premise"] = "以起始月归入月度汇总；横向月份按列值直接读取"
    data["verification_status"] = "未验证（需人工在Excel中核对公式和数值）"
    _cj(d, data)
    print("✓ W2 补齐反馈+预期")


def complete_W4():
    d = ROOT / "W4"
    if not d.exists(): return
    # 补拆分文件（按单位拆，保留纵向合并结构）
    wb_orig = load_workbook(d / "W4_原始_纵向合并.xlsx")
    ws_orig = wb_orig.active

    # 读取各单位的数据范围
    unit_data = {}
    for r in range(2, ws_orig.max_row + 1):
        unit = ws_orig.cell(r, 1).value
        if unit:
            current_unit = unit
        if current_unit:
            unit_data.setdefault(current_unit, []).append(
                [ws_orig.cell(r, c).value for c in range(1, 7)])

    for unit, rows in unit_data.items():
        wb2 = Workbook(); ws2 = wb2.active; ws2.title = unit
        _hdr(ws2, 1, ["所属单位", "项目名称", "子任务", "开始月份", "结束月份", "金额"])
        first_r = 2
        for i, row in enumerate(rows):
            ws2.append(row)
        # 重建纵向合并（每列在首行有值、后续行为空的列做合并）
        col_has_merge = []
        for c in range(1, 7):
            vals = [ws2.cell(r, c).value for r in range(2, ws2.max_row + 1)]
            if vals[0] is not None and all(v is None for v in vals[1:]):
                col_has_merge.append(c)
        for c in col_has_merge:
            if ws2.max_row > 2:
                ws2.merge_cells(start_row=2, start_column=c,
                               end_row=ws2.max_row, end_column=c)
        _save(wb2, d / f"W4_拆分_{unit}.xlsx")

    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["business_premise"] = "按项目整块拆走（子任务跟随项目），纵向合并保留"
    data["expected_split"] = "每个单位一个文件，项目内子任务保持多行+纵向合并"
    data["verification_status"] = "未验证"
    _cj(d, data)
    print("✓ W4 补齐拆分文件（保留合并结构）")


def complete_W6():
    d = ROOT / "W6"
    if not d.exists(): return
    # 补预期全年更新（C003从3月→5月）
    wb = load_workbook(d / "W6_反馈_月份调整.xlsx")
    _save(wb, d / "W6_预期_全年更新.xlsx")  # 反馈即为更新后状态

    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["business_premise"] = "跨月任务按起始月归入月度汇总"
    data["expected_monthly"] = {
        "3月原汇总": "含C001(1-3月)和C003(原3月)",
        "3月新汇总": "C003移出后仅含C001",
        "5月新汇总": "新增C003(调整为5月)和原有的C004(4-6月)",
    }
    data["expected_annual_update"] = "C003起始月3→5, 结束月3→5, 其余不变"
    data["verification_status"] = "未验证"
    _cj(d, data)
    print("✓ W6 补齐预期更新")


# ============================================================ F23 修正
def fix_F23():
    d = ROOT / "F23"
    if not d.exists(): return
    # 重新生成：完成率用百分比验证而非整数
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "完成率", "状态", "截止日期"])
    for i in range(8):
        ws.append([f"CF{i+1:03d}", 0.3 + i * 0.1,
                   "已完成" if i > 4 else "进行中",
                   dt.date(2026, i + 1, 15)])
    ws["B2"].number_format = "0%"
    ws.conditional_formatting.add("B2:B9",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.8"],
                   fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add("B2:B9",
        CellIsRule(operator="lessThan", formula=["0.5"],
                   fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add("D2:D9",
        FormulaRule(formula=['DATE(2026,10,1)-D2>0'],
                    font=Font(color="FF0000", bold=True)))
    # 修正验证：十进制0~1（匹配完成率数据）
    dv = DataValidation(type="decimal", operator="between",
                        formula1="0", formula2="1",
                        error="请输入0~1之间的小数", errorTitle="数值超出范围")
    ws.add_data_validation(dv)
    dv.add("B2:B9")
    # 增加自定义公式验证
    dv2 = DataValidation(type="custom",
                         formula1='AND(B2>=0,B2<=1,ISNUMBER(B2))',
                         error="完成率必须是0~1之间的数字",
                         errorTitle="自定义验证失败")
    ws.add_data_validation(dv2)
    dv2.add("B2:B9")
    _save(wb, d / "F23_条件格式验证.xlsx")

    _cj(d, {
        "case_id": "F23", "type": "保真",
        "checks": {
            "绿色": "完成率≥80%的B列单元格",
            "红色": "完成率<50%的B列单元格",
            "过期红字": "截止日期在2026-10-01之前",
            "十进制验证": "B列限制0~1小数(修正: 原整数1-100与0.3~1.0数据不匹配)",
            "自定义公式验证": "AND(B2>=0,B2<=1,ISNUMBER(B2))",
        },
        "correction": "原版整数1-100验证与0.3~1.0完成率数据矛盾，修正为十进制0~1+自定义公式",
        "verification_status": "未验证",
    })
    print("✓ F23 修正（验证规则与数据匹配）")


# ============================================================ D41 修正
def fix_D41():
    d = ROOT / "D41"
    if not d.exists(): return
    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["expected"] = "保留空行（原生复刻要求：排版空行是原始体验的一部分）"
    data["correction"] = "默认预期从'可保留可删除'修正为'保留'"
    _cj(d, data)
    print("✓ D41 默认预期修正为保留")


# ============================================================ C36 修正
def fix_C36():
    d = ROOT / "C36"
    if not d.exists(): return
    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["expected_changes"] = [
        {"task": "H001", "field": "数量", "old": 10, "new": 12},
        {"task": "H002", "field": "UUID", "old": "UUID-DEF456", "new": None,
         "note": "UUID丢失，需回退到编号匹配"},
        {"task": "新行", "field": "UUID", "old": None, "new": "UUID-ABC123",
         "note": "UUID与H001重复，冲突"},
    ]
    data["correction"] = "补全H001数量10→12变更；UUID列从隐藏变为可见也是变更"
    _cj(d, data)
    print("✓ C36 补全所有变更")


# ============================================================ D40 增强
def fix_D40():
    d = ROOT / "D40"
    if not d.exists(): return
    wb = Workbook(); ws = wb.active; ws.title = "原始"
    _hdr(ws, 1, ["A值", "B值", "C值", "D值", "SUM"])
    for i in range(5):
        r = i + 2
        ws.append([i+1, i*10, i*100, i*1000, f"=SUM(A{r}:C{r})"])  # SUM包含A,B,C

    # 变体：删除B列后SUM范围变化
    ws2 = wb.create_sheet("删除B列后")
    _hdr(ws2, 1, ["A值", "C值", "D值", "SUM"])
    for i in range(5):
        r = i + 2
        # 原SUM(A:C)=A+B+C, 删B后变为SUM(A:B)=A+C(因为B列消失了, 原C变成了新B)
        # 正确的等价公式应该是 A+C(跳过已删除的B)
        ws2.append([i+1, i*100, i*1000, f"=A{r}+B{r}"])  # 新B=原C

    # 变体：无关列D移入SUM范围
    ws3 = wb.create_sheet("D列移入SUM")
    _hdr(ws3, 1, ["A值", "B值", "C值", "D值", "SUM"])
    for i in range(5):
        r = i + 2
        ws3.append([i+1, i*10, i*100, i*1000, f"=SUM(A{r}:D{r})"])  # 扩大到D

    _save(wb, d / "D40_列换序引用.xlsx")
    _cj(d, {
        "case_id": "D40", "type": "保真",
        "checks": {
            "原始SUM": "=SUM(A:C) 包含A+B+C",
            "删除B后": "=A+B (新B=原C), 参与计算的内容变为A+C",
            "D列移入": "=SUM(A:D) 包含A+B+C+D, D的值参与计算",
            "关键": "公式引用范围变化后，参与计算的单元格集合是否正确",
        },
        "verification_status": "未验证",
    })
    print("✓ D40 增强SUM范围变化案例")


# ============================================================ 主入口
if __name__ == "__main__":
    fix_W1()
    fix_W3()
    fix_W5()
    complete_W2()
    complete_W4()
    complete_W6()
    fix_F23()
    fix_D40()
    fix_D41()
    fix_C36()
    print("\n修订完成")
