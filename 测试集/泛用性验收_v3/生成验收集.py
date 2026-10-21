# -*- coding: utf-8 -*-
"""
泛用性与原生保真验收测试集生成器（v3）

原则：
1. 标准答案独立于本项目引擎 — 只用 openpyxl 原生 API 生成，不调用 engine/ 任何代码
2. 先定义业务事实（真值 JSON），再生成匹配的 Excel 文件
3. 固定随机种子，可精确重建
4. 歧义案例的真值是"需要确认什么"，不是编造的唯一答案

用法：
    python 生成验收集.py                    # 生成全部 24 案例
    python 生成验收集.py --cases W1,W2      # 只生成指定案例
    python 生成验收集.py --perf 10000       # 额外生成 10000 条性能文件
"""
import argparse
import json
import random
import datetime as dt
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

ROOT = Path(__file__).resolve().parent
rng = random.Random(20261020)

# ============================================================ 公共数据池
UNITS = ["综合办公室", "财务部", "工程技术部", "安全管理部", "市场营销部"]
UNITS_SHORT = ["综合办", "财务", "工程", "安全", "市场"]
MONTHS = [f"{m}月" for m in range(1, 13)]
THIN = Border(*[Side(style="thin", color="999999")] * 4)
HDR_FONT = Font(bold=True, size=10, name="微软雅黑")
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")

TASKS_POOL = [
    "完成年度预算编制", "组织安全培训演练", "设备巡检维护", "季度报表汇总",
    "完成资料汇总", "成本合计核算", "供应商评估审查", "新员工入职办理",
    "项目验收资料整理", "应急预案修订", "办公耗材采购", "固定资产盘点",
    "客户回访满意度调查", "技术文档归档", "质量体系内审", "合同续签谈判",
    "系统数据备份", "消防安全检查", "年报审计配合", "制度修订发布",
]


def _hdr(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = THIN
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _case_json(case_dir, data):
    p = case_dir / "case.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return p


# ============================================================ W1: 标准月度计划→拆分→反馈→汇总
def gen_W1(d: Path):
    """标准纵向表：月份单列、单位单列、12任务×5单位，含公式、下拉、条件格式"""
    tasks = []
    for i, name in enumerate(TASKS_POOL[:12]):
        month = (i % 12) + 1
        unit = UNITS[i % 5]
        qty = (i + 1) * 10
        tasks.append({"id": f"T{i+1:03d}", "name": name, "unit": unit,
                      "month": month, "qty": qty, "budget": qty * 100,
                      "status": "进行中" if i % 3 else "已完成"})

    # --- 原始表 ---
    wb = Workbook(); ws = wb.active; ws.title = "年度计划"
    ws.append(["2026年度重点工作任务计划表"]); ws.merge_cells("A1:G1")
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["说明：金额单位为元，数量为计划数。"]); ws.merge_cells("A2:G2")
    ws.append([f"编制单位：综合办公室    日期：2026-01-15"]); ws.merge_cells("A3:G3")
    _hdr(ws, 4, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
    for i, t in enumerate(tasks):
        r = i + 5
        ws.append([t["id"], t["name"], t["unit"], f"{t['month']}月",
                   t["qty"], f"=E{r}*100", t["status"]])
    dv = DataValidation(type="list", formula1='"进行中,已完成,已取消"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"G5:G{4 + len(tasks)}")
    ws.conditional_formatting.add(f"G5:G{4 + len(tasks)}",
        CellIsRule(operator="equal", formula=['"已完成"'],
                   fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    for col, w in zip("ABCDEFG", [10, 22, 14, 8, 10, 14, 10]):
        ws.column_dimensions[col].width = w
    _save(wb, d / "W1_原始计划表.xlsx")

    # --- 拆分文件（每个单位一份，只含该单位的任务）---
    for unit in UNITS:
        unit_tasks = [t for t in tasks if t["unit"] == unit]
        if not unit_tasks:
            continue
        wb2 = Workbook(); ws2 = wb2.active; ws2.title = unit
        _hdr(ws2, 1, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
        for i, t in enumerate(unit_tasks):
            r = i + 2
            ws2.append([t["id"], t["name"], t["unit"], f"{t['month']}月",
                        t["qty"], f"=E{r}*100", t["status"]])
        _save(wb2, d / f"W1_拆分_{unit}.xlsx")

    # --- 反馈文件（3个单位有反馈，1个改了数量，1个新增任务，1个未反馈）---
    fb_changes = []
    # 财务部: 改了 T002 数量 20→25
    wb3 = Workbook(); ws3 = wb3.active; ws3.title = "财务部"
    _hdr(ws3, 1, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
    fin_tasks = [t for t in tasks if t["unit"] == "财务部"]
    for i, t in enumerate(fin_tasks):
        r = i + 2
        qty = 25 if t["id"] == "T002" else t["qty"]
        if t["id"] == "T002":
            fb_changes.append({"task": "T002", "field": "计划数量", "old": 20, "new": 25})
        ws3.append([t["id"], t["name"], t["unit"], f"{t['month']}月",
                    qty, f"=E{r}*100", t["status"]])
    _save(wb3, d / "W1_反馈_财务部.xlsx")

    # 工程技术部: 新增任务 T999
    wb4 = Workbook(); ws4 = wb4.active; ws4.title = "工程技术部"
    _hdr(ws4, 1, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
    eng_tasks = [t for t in tasks if t["unit"] == "工程技术部"]
    for i, t in enumerate(eng_tasks):
        r = i + 2
        ws4.append([t["id"], t["name"], t["unit"], f"{t['month']}月",
                    t["qty"], f"=E{r}*100", t["status"]])
    ws4.append(["T999", "新增：设备升级改造", "工程技术部", "6月", 5, "=E7*100", "进行中"])
    fb_changes.append({"task": "T999", "field": "新增", "old": None, "new": "新增任务"})
    _save(wb4, d / "W1_反馈_工程技术部.xlsx")

    # 市场营销部: 将 T005 状态改为已完成
    wb5 = Workbook(); ws5 = wb5.active; ws5.title = "市场营销部"
    _hdr(ws5, 1, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
    mkt_tasks = [t for t in tasks if t["unit"] == "市场营销部"]
    for i, t in enumerate(mkt_tasks):
        r = i + 2
        status = "已完成" if t["id"] == "T005" else t["status"]
        if t["id"] == "T005":
            fb_changes.append({"task": "T005", "field": "状态", "old": "进行中", "new": "已完成"})
        ws5.append([t["id"], t["name"], t["unit"], f"{t['month']}月",
                    t["qty"], f"=E{r}*100", status])
    _save(wb5, d / "W1_反馈_市场营销部.xlsx")

    # --- 预期输出 ---
    # 月度汇总（以6月为例）
    june_tasks = [t for t in tasks if t["month"] == 6]
    june_new = [{"id": "T999", "name": "新增：设备升级改造", "unit": "工程技术部",
                 "month": 6, "qty": 5, "budget": 500, "status": "进行中"}]
    all_june = june_tasks + june_new
    wb6 = Workbook(); ws6 = wb6.active; ws6.title = "6月汇总"
    _hdr(ws6, 1, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
    for i, t in enumerate(all_june):
        ws6.append([t["id"], t["name"], t["unit"], "6月", t["qty"],
                    t.get("budget", t["qty"] * 100), t["status"]])
    _save(wb6, d / "W1_预期_6月汇总.xlsx")

    # 全年更新版
    updated = {t["id"]: dict(t) for t in tasks}
    for c in fb_changes:
        if c["task"] in updated:
            if c["field"] == "计划数量":
                updated[c["task"]]["qty"] = c["new"]
            elif c["field"] == "状态":
                updated[c["task"]]["status"] = c["new"]
    all_updated = list(updated.values()) + [
        {"id": "T999", "name": "新增：设备升级改造", "unit": "工程技术部",
         "month": 6, "qty": 5, "budget": 500, "status": "进行中"}]
    wb7 = Workbook(); ws7 = wb7.active; ws7.title = "年度计划(更新)"
    _hdr(ws7, 1, ["任务编号", "任务名称", "责任单位", "月份", "计划数量", "预算金额(元)", "状态"])
    for t in sorted(all_updated, key=lambda x: x["id"]):
        ws7.append([t["id"], t["name"], t["unit"], f"{t['month']}月",
                    t["qty"], t.get("budget", t["qty"] * 100), t["status"]])
    _save(wb7, d / "W1_预期_全年更新.xlsx")

    return _case_json(d, {
        "case_id": "W1", "type": "完整流程",
        "scenario": "标准年度计划表，月份在单列，按单位拆分后收集反馈",
        "structure": {"header_row": 4, "pre_rows": 3, "columns": 7,
                      "month_location": "单列(D列)", "unit_location": "单列(C列)"},
        "ground_truth": {"total_tasks": len(tasks), "units": UNITS,
                         "tasks_by_unit": {u: len([t for t in tasks if t["unit"] == u]) for u in UNITS},
                         "june_tasks": len(all_june)},
        "expected_split": "每个单位一个文件，只含该单位任务",
        "expected_monthly": "6月汇总含原有6月任务+新增T999",
        "expected_annual_update": "T002数量20→25, T005状态→已完成, 新增T999",
        "feedback_changes": fb_changes,
        "no_feedback_units": ["综合办公室", "安全管理部"],
        "ambiguity": [],
        "invariants": ["未反馈单位的任务不变", "任务编号不被重排",
                       "预算金额公式=E*100语义不变", "状态下拉可选值不变"],
    })


# ============================================================ W2: 月份横向12列
def gen_W2(d: Path):
    """月份在列方向（1月~12月横向），按单位拆分"""
    tasks_data = {}
    for unit in UNITS[:3]:
        for i in range(4):
            task_name = TASKS_POOL[i + UNITS.index(unit)]
            for m in range(1, 13):
                tasks_data[(unit, task_name, m)] = rng.randint(5, 50)

    wb = Workbook(); ws = wb.active; ws.title = "月度计划"
    ws.append(["各单位月度任务数量分配表"]); ws.merge_cells("A1:M1")
    ws["A1"].font = Font(bold=True, size=13)
    _hdr(ws, 2, ["责任单位", "任务名称"] + MONTHS + ["合计"])
    r = 3
    for unit in UNITS[:3]:
        for i in range(4):
            task_name = TASKS_POOL[i + UNITS.index(unit)]
            row = [unit if i == 0 else None, task_name]
            row += [tasks_data.get((unit, task_name, m), 0) for m in range(1, 13)]
            row += [f"=SUM(C{r}:N{r})"]
            ws.append(row)
            r += 1
    ws.append(["合计"] + [None] + [f"=SUM({get_column_letter(3+j)}3:{get_column_letter(3+j)}{r-1})" for j in range(12)] + [f"=SUM(O3:O{r-1})"])
    for col in "AB":
        ws.column_dimensions[col].width = 16
    _save(wb, d / "W2_原始_横向月份.xlsx")

    for unit in UNITS[:3]:
        wb2 = Workbook(); ws2 = wb2.active; ws2.title = unit
        _hdr(ws2, 1, ["任务名称"] + MONTHS + ["合计"])
        rr = 2
        for i in range(4):
            task_name = TASKS_POOL[i + UNITS.index(unit)]
            row = [task_name] + [tasks_data.get((unit, task_name, m), 0) for m in range(1, 13)]
            row += [f"=SUM(B{rr}:M{rr})"]
            ws2.append(row)
            rr += 1
        _save(wb2, d / f"W2_拆分_{unit}.xlsx")

    return _case_json(d, {
        "case_id": "W2", "type": "完整流程",
        "scenario": "月份横向排列（1月~12月为列），每行=单位×任务，按单位拆分",
        "structure": {"header_row": 2, "pre_rows": 1, "columns": 15,
                      "month_location": "横向列(C~N列)", "unit_location": "A列分组"},
        "ground_truth": {"units": UNITS[:3], "tasks_per_unit": 4, "months": 12},
        "expected_split": "每个单位4行任务，保留月份列+合计公式",
        "ambiguity": ["合计行是否应包含在拆分文件中（原始表最后一行是跨单位合计）"],
        "invariants": ["合计公式=SUM(C:N)范围正确", "月份数值不变"],
    })


# ============================================================ W3: 双行表头+合并单元格
def gen_W3(d: Path):
    """双行表头，父级跨列合并，子级细列"""
    wb = Workbook(); ws = wb.active; ws.title = "任务台账"
    ws.append(["任务执行情况统计表"]); ws.merge_cells("A1:H1")
    ws.append(["填报说明：请各单位如实填报执行进度。"]); ws.merge_cells("A2:H2")
    _hdr(ws, 3, ["编号", "任务名称", "基本信息", "基本信息", "进度", "进度", "进度", "备注"])
    ws.merge_cells("C3:D3")  # 基本信息 跨 C-D
    ws.merge_cells("E3:G3")  # 进度 跨 E-G
    _hdr(ws, 4, ["", "", "责任单位", "完成时限", "计划数", "完成数", "完成率", ""])
    tasks = []
    for i, name in enumerate(TASKS_POOL[:8]):
        unit = UNITS[i % 4]
        plan = (i + 1) * 10
        done = int(plan * (0.3 + 0.1 * i))
        t = {"id": f"A{i+1:03d}", "name": name, "unit": unit,
             "deadline": f"2026-{(i%12)+1:02d}-15", "plan": plan, "done": done}
        tasks.append(t)
        r = i + 5
        ws.append([t["id"], t["name"], t["unit"], t["deadline"],
                   t["plan"], t["done"], f"=F{r}/E{r}", None])
    ws.conditional_formatting.add(f"G5:G{4+len(tasks)}",
        CellIsRule(operator="lessThan", formula=["0.5"],
                   fill=PatternFill("solid", fgColor="FFC7CE")))
    _save(wb, d / "W3_原始_双行表头.xlsx")

    for unit in UNITS[:4]:
        unit_tasks = [t for t in tasks if t["unit"] == unit]
        if not unit_tasks:
            continue
        wb2 = Workbook(); ws2 = wb2.active; ws2.title = unit
        _hdr(ws2, 1, ["编号", "任务名称", "责任单位", "完成时限", "计划数", "完成数", "完成率", "备注"])
        for i, t in enumerate(unit_tasks):
            r = i + 2
            ws2.append([t["id"], t["name"], t["unit"], t["deadline"],
                        t["plan"], t["done"], f"=F{r}/E{r}", None])
        _save(wb2, d / f"W3_拆分_{unit}.xlsx")

    return _case_json(d, {
        "case_id": "W3", "type": "完整流程",
        "structure": {"header_rows": [3, 4], "pre_rows": 2,
                      "merges": ["C3:D3(基本信息)", "E3:G3(进度)"],
                      "columns": 8},
        "ground_truth": {"tasks": len(tasks),
                         "flattened_keys": ["编号", "任务名称", "基本信息/责任单位",
                                            "基本信息/完成时限", "进度/计划数", "进度/完成数",
                                            "进度/完成率", "备注"]},
        "invariants": ["完成率公式=F/E不变", "条件格式(完成率<0.5红色)保留",
                       "双行表头结构在拆分后保留"],
    })


# ============================================================ W4: 单位纵向合并+任务多行
def gen_W4(d: Path):
    """单位用纵向合并单元格分组，一项任务占多行（子任务）"""
    wb = Workbook(); ws = wb.active; ws.title = "项目分解"
    _hdr(ws, 1, ["所属单位", "项目名称", "子任务", "开始月份", "结束月份", "金额"])
    data = [
        ("综合办公室", "制度建设", ["起草", "征求意见", "发布"], "1月", "3月", 30000),
        ("财务部",     "预算改革", ["调研", "方案", "实施"], "2月", "6月", 80000),
        ("工程技术部", "系统升级", ["需求", "开发", "测试", "上线"], "4月", "9月", 150000),
        ("安全管理部", "应急演练", ["预案", "演练", "总结"], "5月", "7月", 20000),
    ]
    r = 2
    unit_ranges = {}
    for unit, proj, subs, start, end, amt in data:
        first_r = r
        for j, sub in enumerate(subs):
            ws.append([unit if j == 0 else None, proj if j == 0 else None,
                       sub, start if j == 0 else None, end if j == 0 else None,
                       amt // len(subs) if j == 0 else None])
            r += 1
        if len(subs) > 1:
            for col in [1, 2, 4, 5, 6]:
                ws.merge_cells(start_row=first_r, start_column=col,
                               end_row=r - 1, end_column=col)
        unit_ranges[unit] = (first_r, r - 1)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    _save(wb, d / "W4_原始_纵向合并.xlsx")

    return _case_json(d, {
        "case_id": "W4", "type": "结构案例",
        "scenario": "单位在A列纵向合并，每个项目含多个子任务占多行",
        "structure": {"unit_location": "A列纵向合并单元格",
                      "task_rows": "一任务多行(子任务)"},
        "ground_truth": {"projects": 4, "total_subtasks": 13,
                         "unit_ranges": {u: f"r{v[0]}-r{v[1]}" for u, v in unit_ranges.items()}},
        "ambiguity": ["拆分时子任务行是否应独立(拆散项目)，还是按项目整块拆走"],
        "invariants": ["金额只在每单位首行(合并单元格锚点)，子行金额为空",
                       "纵向合并在拆分后应还原为该单位独立的合并"],
    })


# ============================================================ W5: 月份来自Sheet名
def gen_W5(d: Path):
    """每月一个Sheet，Sheet名=月份"""
    wb = Workbook()
    for m in range(1, 4):
        ws = wb.create_sheet(f"{m}月")
        _hdr(ws, 1, ["任务编号", "任务名称", "责任单位", "数量", "金额"])
        count = 3 + m
        for i in range(count):
            idx = (m - 1) * 5 + i
            ws.append([f"M{m:02d}-{i+1:02d}", TASKS_POOL[idx % len(TASKS_POOL)],
                       UNITS[i % 5], rng.randint(1, 20), rng.randint(1000, 50000)])
    wb.remove(wb["Sheet"])
    _save(wb, d / "W5_原始_Sheet月份.xlsx")

    return _case_json(d, {
        "case_id": "W5", "type": "结构案例",
        "scenario": "月份在Sheet名中（1月/2月/3月各一个Sheet），不在单元格内",
        "ground_truth": {"sheets": ["1月", "2月", "3月"],
                         "tasks_per_sheet": [3, 4, 5]},
        "ambiguity": ["汇总时月份信息来自Sheet名而非列值，工具是否能正确提取"],
    })


# ============================================================ W6: 跨月任务+反馈调整月份
def gen_W6(d: Path):
    """任务有起止月份，反馈中将某任务从3月调整到5月"""
    tasks = [
        {"id": "C001", "name": "年度审计", "unit": "财务部", "start": 1, "end": 3, "qty": 1},
        {"id": "C002", "name": "系统上线", "unit": "工程技术部", "start": 2, "end": 5, "qty": 3},
        {"id": "C003", "name": "安全培训", "unit": "安全管理部", "start": 3, "end": 3, "qty": 5},
        {"id": "C004", "name": "市场推广", "unit": "市场营销部", "start": 4, "end": 6, "qty": 8},
    ]
    wb = Workbook(); ws = wb.active; ws.title = "跨月任务"
    _hdr(ws, 1, ["编号", "任务名称", "责任单位", "起始月", "结束月", "数量"])
    for t in tasks:
        ws.append([t["id"], t["name"], t["unit"], f"{t['start']}月", f"{t['end']}月", t["qty"]])
    _save(wb, d / "W6_原始_跨月.xlsx")

    # 反馈: C003 从3月调整到5月
    wb2 = Workbook(); ws2 = wb2.active; ws2.title = "跨月任务"
    _hdr(ws2, 1, ["编号", "任务名称", "责任单位", "起始月", "结束月", "数量"])
    for t in tasks:
        start = 5 if t["id"] == "C003" else t["start"]
        end = 5 if t["id"] == "C003" else t["end"]
        ws2.append([t["id"], t["name"], t["unit"], f"{start}月", f"{end}月", t["qty"]])
    _save(wb2, d / "W6_反馈_月份调整.xlsx")

    return _case_json(d, {
        "case_id": "W6", "type": "完整流程",
        "scenario": "跨月任务，反馈将C003从3月调整到5月",
        "ground_truth": {"tasks": tasks,
                         "feedback_change": {"task": "C003", "field": "起始月+结束月",
                                             "old": "3月→3月", "new": "5月→5月"}},
        "ambiguity": ["跨月任务(C002 2-5月)应归入哪个月份的汇总——起始月？结束月？每个月都算？"],
        "invariants": ["C001/C002/C004 不变", "月份调整后C003的数量不变"],
    })


# ============================================================ S系列: 结构与关系案例

def gen_S07(d):
    """表头在第12行"""
    wb = Workbook(); ws = wb.active
    for i in range(11):
        ws.append([f"前置区第{i+1}行" if i == 0 else None])
    _hdr(ws, 12, ["编号", "名称", "数量"])
    for i in range(5):
        ws.append([f"S{i+1}", TASKS_POOL[i], i + 1])
    _save(wb, d / "S07_表头第12行.xlsx")
    return _case_json(d, {"case_id": "S07", "type": "结构", "scenario": "表头在第12行，前面11行前置区"})


def gen_S08(d):
    """同名末级列：计划/金额 vs 实际/金额"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "计划", "计划", "实际", "实际", "差异"])
    ws.append(["", "数量", "金额", "数量", "金额", ""])
    for i in range(5):
        r = i + 3
        ws.append([f"P{i+1}", 100, 10000, 90, 9000, f"=E{r}-C{r}"])
    ws.merge_cells("B1:C1"); ws.merge_cells("D1:E1")
    _save(wb, d / "S08_同名末级列.xlsx")
    return _case_json(d, {
        "case_id": "S08", "type": "结构",
        "scenario": "计划/金额 与 实际/金额 末级名称相同、父级不同",
        "flattened_keys": ["编号", "计划/数量", "计划/金额", "实际/数量", "实际/金额", "差异"],
        "invariants": ["差异公式=E-C(实际数量-计划数量)不变"],
    })


def gen_S09(d):
    """单位名称只在标题中"""
    wb = Workbook(); ws = wb.active
    ws.append(["财务部2026年工作任务清单"]); ws.merge_cells("A1:D1")
    _hdr(ws, 2, ["序号", "任务名称", "数量", "备注"])
    for i in range(6):
        ws.append([i + 1, TASKS_POOL[i + 5], rng.randint(1, 10), None])
    _save(wb, d / "S09_单位在标题.xlsx")
    return _case_json(d, {
        "case_id": "S09", "type": "关系",
        "scenario": "单位名称只出现在标题行，不在数据列中",
        "unit": "财务部",
        "ambiguity": ["工具是否能从标题中提取单位信息用于拆分/汇总"],
    })


def gen_S10(d):
    """空白单位+无法确定归属"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "责任单位", "数量"])
    rows = [
        ("U001", "任务A", "综合办公室", 10),
        ("U002", "任务B", None, 5),         # 空白单位
        ("U003", "任务C", "待定", 3),        # 无法确定
        ("U004", "任务D", "", 8),            # 空字符串
    ]
    for r in rows:
        ws.append(list(r))
    _save(wb, d / "S10_空白单位.xlsx")
    return _case_json(d, {
        "case_id": "S10", "type": "关系",
        "scenario": "单位列有空白、'待定'、空字符串",
        "ambiguity": ["U002/U004(空白单位)不应默认分配给任何单位",
                      "U003(待定)需人工确认归属"],
    })


def gen_S11(d):
    """同名不同任务（有编号/无编号/重复编号）"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务名称", "单位", "数量"])
    ws.append(["D001", "安全培训", "安全部", 10])
    ws.append(["D002", "安全培训", "综合办", 5])     # 同名不同编号不同单位
    ws.append([None,  "安全培训", "财务部", 3])       # 同名无编号
    ws.append(["D001", "安全培训", "市场部", 8])       # 重复编号！
    _save(wb, d / "S11_同名任务.xlsx")
    return _case_json(d, {
        "case_id": "S11", "type": "关系",
        "scenario": "同名任务'安全培训'出现4次：不同编号、无编号、重复编号",
        "ambiguity": ["D001出现两次(不同单位)——是数据错误还是两个独立任务？",
                      "无编号的安全培训如何唯一标识？",
                      "答案：不应自动去重或合并，应列出冲突待人工确认"],
    })


def gen_S12(d):
    """数据区大块纵向合并（不应误认为表头）"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["地区", "任务", "数量"])
    for region in ["华东", "华东", "华东", "华南", "华南", "华北"]:
        ws.append([region, TASKS_POOL[rng.randint(0, 19)], rng.randint(1, 20)])
    # A列纵向合并
    ws.merge_cells("A2:A4")  # 华东 3行
    ws.merge_cells("A5:A6")  # 华南 2行
    _save(wb, d / "S12_数据区合并.xlsx")
    return _case_json(d, {
        "case_id": "S12", "type": "反例",
        "scenario": "A列地区用纵向合并(华东3行/华南2行/华北1行) — 这是数据不是表头",
        "invariants": ["A列合并不能被当作多行表头",
                       "拆分后地区分组应保留合并结构"],
    })


# ============================================================ S系列: 反馈变化案例

def gen_S13(d):
    """反馈只改排序，不改内容"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "数量"])
    data = [("F001", "任务A", 10), ("F002", "任务B", 20), ("F003", "任务C", 30)]
    for r in data:
        ws.append(list(r))
    _save(wb, d / "S13_原始.xlsx")

    wb2 = Workbook(); ws2 = wb2.active
    _hdr(ws2, 1, ["编号", "任务", "数量"])
    for r in reversed(data):  # 逆序
        ws2.append(list(r))
    _save(wb2, d / "S13_反馈_仅排序.xlsx")
    return _case_json(d, {
        "case_id": "S13", "type": "反馈",
        "scenario": "反馈文件只是把行序打乱（逆序），内容完全没变",
        "expected": "变更清单应为空（零修改），不应产生任何内容变更",
    })


def gen_S14(d):
    """反馈只改样式（红字），不改数据"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "数量", "状态"])
    data = [("S001", "任务A", 10, "进行中"), ("S002", "任务B", 20, "已完成")]
    for r in data:
        ws.append(list(r))
    _save(wb, d / "S14_原始.xlsx")

    wb2 = Workbook(); ws2 = wb2.active
    _hdr(ws2, 1, ["编号", "任务", "数量", "状态"])
    for r in data:
        ws2.append(list(r))
    ws2["C2"].font = Font(color="FF0000")  # 只改颜色
    ws2["D3"].fill = PatternFill("solid", fgColor="FFFF00")  # 只改填充
    _save(wb2, d / "S14_反馈_仅样式.xlsx")
    return _case_json(d, {
        "case_id": "S14", "type": "反馈",
        "scenario": "反馈只改了字体颜色和单元格填充，数据值完全没变",
        "expected": "数据层面变更清单为空；样式变更可另行报告但不影响数据",
    })


def gen_S15(d):
    """原表已有红字（不是本次修改）"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "数量"])
    ws.append(["R001", "原表红字任务", 10])
    ws.append(["R002", "普通任务", 20])
    ws["B2"].font = Font(color="FF0000")  # 原始表就有红字
    _save(wb, d / "S15_原表红字.xlsx")

    wb2 = Workbook(); ws2 = wb2.active
    _hdr(ws2, 1, ["编号", "任务", "数量"])
    ws2.append(["R001", "原表红字任务", 10])
    ws2.append(["R002", "普通任务", 25])  # 只改了R002数量
    ws2["B2"].font = Font(color="FF0000")  # 红字保留
    _save(wb2, d / "S15_反馈.xlsx")
    return _case_json(d, {
        "case_id": "S15", "type": "反馈",
        "scenario": "原表R001就是红字，反馈只改了R002数量。红字不应被当作本次修改",
        "expected": "变更只有1条：R002数量20→25。R001红字是原有状态",
    })


def gen_S16(d):
    """同单位两个反馈版本"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "数量"])
    ws.append(["V001", "任务A", 10])
    ws.append(["V002", "任务B", 20])
    _save(wb, d / "S16_原始.xlsx")

    for ver, change in [("v1", [("V001", "任务A", 12)]), ("v2", [("V001", "任务A", 15)])]:
        wb2 = Workbook(); ws2 = wb2.active
        _hdr(ws2, 1, ["编号", "任务", "数量"])
        data = {"V001": ("任务A", change[0][2]), "V002": ("任务B", 20)}
        for k, (name, qty) in data.items():
            ws2.append([k, name, qty])
        _save(wb2, d / f"S16_反馈_{ver}.xlsx")
    return _case_json(d, {
        "case_id": "S16", "type": "反馈",
        "scenario": "同一单位提交了两个版本的反馈（V001数量分别改为12和15）",
        "ambiguity": ["应以哪个版本为准？需人工确认，不能自动取最新或最大值"],
    })


def gen_S17(d):
    """未反馈单位不应被清空"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "单位", "数量"])
    for i, unit in enumerate(UNITS):
        ws.append([f"N{i+1:03d}", TASKS_POOL[i], unit, 10 * (i + 1)])
    _save(wb, d / "S17_原始.xlsx")

    # 只有财务部反馈
    wb2 = Workbook(); ws2 = wb2.active
    _hdr(ws2, 1, ["编号", "任务", "单位", "数量"])
    ws2.append(["N002", TASKS_POOL[1], "财务部", 99])
    _save(wb2, d / "S17_反馈_仅财务部.xlsx")
    return _case_json(d, {
        "case_id": "S17", "type": "反馈",
        "scenario": "5个单位只有财务部反馈（N002数量改为99），其余4个未反馈",
        "expected": "N002数量→99；N001/N003/N004/N005保持原值不变。未反馈≠删除",
    })


def gen_S18(d):
    """多单位对同一任务给出不同修改"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "任务", "参与单位", "数量"])
    ws.append(["X001", "联合检查", "综合办+安全部", 20])
    _save(wb, d / "S18_原始.xlsx")

    for unit, qty in [("综合办", 15), ("安全部", 25)]:
        wb2 = Workbook(); ws2 = wb2.active
        _hdr(ws2, 1, ["编号", "任务", "参与单位", "数量"])
        ws2.append(["X001", "联合检查", unit, qty])
        _save(wb2, d / f"S18_反馈_{unit}.xlsx")
    return _case_json(d, {
        "case_id": "S18", "type": "反馈",
        "scenario": "联合任务X001（原数量20），综合办反馈改为15，安全部反馈改为25",
        "ambiguity": ["两个单位给出不同数量（15 vs 25），以谁为准？",
                      "不能自动取平均(20)或取某一方，需人工确认",
                      "也不能未经约定把数量均分"],
    })


# ============================================================ F系列: 保真案例

def gen_F19(d):
    """数字格式与特殊值"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "百分比", "货币", "长编号", "前导零", "日期", "空值"])
    ws.append(["F001", 0.356, 12345.678, "123456789012345678", "007", dt.date(2026, 10, 1), None])
    ws.append(["F002", 0.05, -5000.5, "9" * 18, "000", dt.date(2026, 12, 31), ""])
    ws["B2"].number_format = "0.0%"
    ws["B3"].number_format = "0.00%"
    ws["C2"].number_format = "#,##0.00"
    ws["C3"].number_format = '"¥"#,##0.00'
    ws["D2"].number_format = "@"
    ws["D3"].number_format = "@"
    ws["E2"].number_format = "000"
    ws["F2"].number_format = "yyyy年m月d日"
    ws["F3"].number_format = "yyyy-mm-dd"
    _save(wb, d / "F19_数字格式.xlsx")
    return _case_json(d, {
        "case_id": "F19", "type": "保真",
        "checks": {
            "百分比": "B2=0.0%显示35.6%, B3=0.00%显示5.00%",
            "货币": "C2=#,##0.00显示12,345.68, C3=¥-5,000.50",
            "长编号": "18位数字必须保留全部位数(文本格式@)",
            "前导零": "007/000必须保留前导零",
            "日期": "yyyy年m月d日格式",
            "空值": "G2=None, G3=空字符串(不同)",
        },
    })


def gen_F20(d):
    """公式与跨Sheet引用"""
    wb = Workbook()
    ws1 = wb.active; ws1.title = "数据"
    _hdr(ws1, 1, ["编号", "数量", "单价"])
    for i in range(5):
        ws1.append([f"GS{i+1}", i + 1, 10.5 * (i + 1)])
    ws2 = wb.create_sheet("汇总")
    ws2.append(["总计"])
    ws2.append(["=SUM(数据!B2:B6)"])
    ws2.append(["=SUM(数据!C2:C6)"])
    ws2.append(["=数据!B2*数据!C2"])
    ws2.append(["=AVERAGE(数据!B2:B6,数据!C2:C6)"])
    ws1["E2"] = "=B2*C2"
    ws1["E3"] = "=SUM($B$2:$B$6)"
    _save(wb, d / "F20_公式跨表.xlsx")
    return _case_json(d, {
        "case_id": "F20", "type": "保真",
        "checks": {
            "SUM跨Sheet": "汇总!A2 = SUM(数据!B2:B6) = 15",
            "混合引用": "汇总!A4 = 数据!B2*数据!C2 = 10.5",
            "绝对引用": "数据!E3 = SUM($B$2:$B$6) = 15",
            "行内公式": "数据!E2 = B2*C2 = 10.5",
        },
        "note": "openpyxl生成的公式无缓存值，Excel打开后自动计算",
    })


def gen_F21(d):
    """打印设置与冻结窗格"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号"] + [f"列{i}" for i in range(2, 10)])
    for r in range(2, 30):
        ws.append([f"PR{r:03d}"] + [rng.randint(1, 100) for _ in range(8)])
    ws.freeze_panes = "B2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    ws.print_area = "A1:I29"
    ws.oddHeader.center.text = "第 &[Page] 页"
    ws.oddFooter.right.text = "共 &N 页"
    _save(wb, d / "F21_打印设置.xlsx")
    return _case_json(d, {
        "case_id": "F21", "type": "保真",
        "checks": {"冻结": "B2(首行+首列)", "方向": "横向", "纸张": "A4",
                   "缩放": "宽度适配1页", "打印标题": "第1行",
                   "打印区域": "A1:I29", "页眉": "第X页", "页脚": "共N页"},
    })


def gen_F22(d):
    """批注、超链接、隐藏配置页"""
    wb = Workbook()
    ws = wb.active; ws.title = "数据"
    _hdr(ws, 1, ["编号", "任务", "链接"])
    ws.append(["H001", "查看详情", None])
    ws["C2"].hyperlink = "https://example.com/detail"
    ws["C2"].value = "点击查看"
    ws["B2"].comment = Comment("这是一个批注\n第二行内容", "管理员")
    cfg = wb.create_sheet("配置")
    cfg["A1"] = "选项A"; cfg["A2"] = "选项B"; cfg["A3"] = "选项C"
    cfg.sheet_state = "hidden"
    dv = DataValidation(type="list", formula1="配置!$A$1:$A$3")
    ws.add_data_validation(dv)
    dv.add("B2:B10")
    _save(wb, d / "F22_批注链接隐藏页.xlsx")
    return _case_json(d, {
        "case_id": "F22", "type": "保真",
        "checks": {"批注": "B2有2行文字批注", "超链接": "C2指向example.com",
                   "隐藏配置页": "下拉引用隐藏Sheet'配置'",
                   "下拉": "B2:B10可选选项A/B/C"},
    })


def gen_F23(d):
    """条件格式+自定义验证"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "完成率", "状态", "截止日期"])
    for i in range(8):
        ws.append([f"CF{i+1:03d}", 0.3 + i * 0.1,
                   "已完成" if i > 4 else "进行中",
                   dt.date(2026, i + 1, 15)])
    ws.conditional_formatting.add("B2:B9",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.8"],
                   fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add("B2:B9",
        CellIsRule(operator="lessThan", formula=["0.5"],
                   fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add("D2:D9",
        FormulaRule(formula=[f'DATE(2026,10,1)-D2>0'],
                    font=Font(color="FF0000", bold=True)))
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="100",
                        error="请输入1-100的整数", errorTitle="数值超出范围")
    ws.add_data_validation(dv)
    dv.add("B2:B9")
    _save(wb, d / "F23_条件格式验证.xlsx")
    return _case_json(d, {
        "case_id": "F23", "type": "保真",
        "checks": {"绿色": "完成率≥80%", "红色": "完成率<50%",
                   "过期红字": "截止日期在2026-10-01之前",
                   "整数验证": "B列限制1-100"},
    })


def gen_F24(d):
    """行高列宽+隐藏行列+分组"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号", "A", "B", "C", "D", "E", "备注"])
    for i in range(10):
        ws.append([f"RW{i+1:02d}", i, i * 2, i * 3, i * 4, i * 5, None])
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[5].height = 12
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["D"].hidden = True
    ws.row_dimensions[8].hidden = True
    _save(wb, d / "F24_行高列宽隐藏.xlsx")
    return _case_json(d, {
        "case_id": "F24", "type": "保真",
        "checks": {"行高": "r2=30, r5=12", "列宽": "F=3, G=30",
                   "隐藏列": "D列隐藏", "隐藏行": "r8隐藏"},
    })


# ============================================================ 主入口
CASES = {
    "W1": gen_W1, "W2": gen_W2, "W3": gen_W3,
    "W4": gen_W4, "W5": gen_W5, "W6": gen_W6,
    "S07": gen_S07, "S08": gen_S08, "S09": gen_S09,
    "S10": gen_S10, "S11": gen_S11, "S12": gen_S12,
    "S13": gen_S13, "S14": gen_S14, "S15": gen_S15,
    "S16": gen_S16, "S17": gen_S17, "S18": gen_S18,
    "F19": gen_F19, "F20": gen_F20, "F21": gen_F21,
    "F22": gen_F22, "F23": gen_F23, "F24": gen_F24,
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cases", default="", help="逗号分隔的案例ID，空=全部")
    ap.add_argument("--perf", type=int, default=0, help="额外生成性能测试文件条数")
    args = ap.parse_args()

    ids = args.cases.split(",") if args.cases else list(CASES.keys())
    generated = []
    for cid in ids:
        cid = cid.strip()
        if cid not in CASES:
            print(f"⚠ 未知案例 {cid}")
            continue
        d = ROOT / cid
        if d.exists():
            target = d.resolve()
            if target.parent != ROOT.resolve() or target.name not in CASES:
                raise ValueError(f"拒绝重建测试集之外的目录: {target}")
            shutil.rmtree(target)
        CASES[cid](d)
        files = list(d.rglob("*.xlsx")) + list(d.rglob("*.json"))
        generated.append((cid, len(files)))
        print(f"✓ {cid} ({len(files)}个文件)")

    if args.perf > 0:
        gen_perf(args.perf)

    print(f"\n完成: {len(generated)}个案例, 共{sum(n for _, n in generated)}个文件")


def gen_perf(n):
    """可选性能测试文件"""
    d = ROOT / "性能"
    wb = Workbook(); ws = wb.active; ws.title = "大批量"
    _hdr(ws, 1, ["编号", "任务名称", "单位", "月份", "数量", "金额"])
    for i in range(n):
        ws.append([f"P{i:06d}", TASKS_POOL[i % len(TASKS_POOL)],
                   UNITS[i % 5], f"{(i % 12) + 1}月", i % 100 + 1, i * 10])
    _save(wb, d / f"性能_{n}条.xlsx")
    print(f"✓ 性能测试 {n}条")


if __name__ == "__main__":
    main()
