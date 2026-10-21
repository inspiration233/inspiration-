# -*- coding: utf-8 -*-
"""从原始夹具重建 W2/W3/W6 的反馈与预期；不在生成阶段声明验收通过。"""
import json
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent
THIN = Border(*[Side(style="thin", color="999999")] * 4)
HDR_F = Font(bold=True, size=10, name="微软雅黑")
HDR_B = PatternFill("solid", fgColor="D9E1F2")


def _hdr(ws, row, values):
    for col, value in enumerate(values, 1):
        cell = ws.cell(row, col, value)
        cell.font, cell.fill, cell.border = HDR_F, HDR_B, THIN


def _save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def _case(directory):
    data = json.loads((directory / "case.json").read_text(encoding="utf-8"))
    for key in ("fidelity_check_results", "correction", "expected_monthly_summary",
                "expected_monthly", "expected_annual", "expected_annual_update"):
        data.pop(key, None)
    data["verification_status"] = "待独立自检；Excel/WPS显示效果未人工验收"
    return data


def _cj(directory, data):
    (directory / "case.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _artifacts(original, splits, feedback, monthly, annual):
    return {"original": [original], "split": splits, "feedback": [feedback],
            "expected_monthly": [monthly], "expected_annual": [annual]}


def rebuild_W2():
    """唯一改动：财务部/组织安全培训演练/1月，18→99。"""
    directory = ROOT / "W2"
    original_name = "W2_原始_横向月份.xlsx"
    wb = load_workbook(directory / original_name)
    source = wb.active
    records, unit = [], None
    for row in range(3, source.max_row + 1):
        unit = source.cell(row, 1).value or unit
        task = source.cell(row, 2).value
        if task is not None:
            records.append({"row": row, "unit": unit, "task": task,
                            "months": [source.cell(row, col).value for col in range(3, 15)]})
    keys = [(item["unit"], item["task"]) for item in records]
    if len(keys) != len(set(keys)):
        raise ValueError("W2 原始数据存在重复的单位+任务，无法唯一定位反馈")
    target_key = ("财务部", "组织安全培训演练")
    targets = [item for item in records if (item["unit"], item["task"]) == target_key]
    if len(targets) != 1 or targets[0]["months"][0] != 18:
        raise ValueError("W2 应有且仅有财务部组织安全培训演练，1月基线为18")
    target = targets[0]

    # 始终以拆分原件为模板，按原始表恢复12个月，避免上一轮误改2月残留。
    feedback_name = "W2_反馈_财务部.xlsx"
    feedback = load_workbook(directory / "W2_拆分_财务部.xlsx")
    output = feedback.active
    expected_header = ["任务名称"] + [f"{month}月" for month in range(1, 13)] + ["合计"]
    if [output.cell(1, col).value for col in range(1, 15)] != expected_header:
        raise ValueError("W2 财务部拆分表列结构与夹具约定不一致")
    finance = {item["task"]: item for item in records if item["unit"] == "财务部"}
    feedback_tasks = [output.cell(row, 1).value for row in range(2, output.max_row + 1)]
    if len(feedback_tasks) != len(finance) or set(feedback_tasks) != set(finance):
        raise ValueError("W2 财务部拆分表的任务集合与原始表不一致")
    for row in range(2, output.max_row + 1):
        item = finance[output.cell(row, 1).value]
        for offset, value in enumerate(item["months"], 2):
            output.cell(row, offset, value)
        source_formula = source.cell(item["row"], 15)
        output.cell(row, 14, Translator(source_formula.value, origin=source_formula.coordinate)
                    .translate_formula(output.cell(row, 14).coordinate))
        if (item["unit"], item["task"]) == target_key:
            output.cell(row, 2, 99)
    _save(feedback, directory / feedback_name)

    monthly_name = "W2_预期_1月汇总.xlsx"
    monthly = Workbook()
    sheet = monthly.active
    sheet.title = "1月汇总"
    _hdr(sheet, 1, ["责任单位", "任务名称", "1月数量"])
    for item in records:
        value = 99 if (item["unit"], item["task"]) == target_key else item["months"][0]
        sheet.append([item["unit"], item["task"], value])
    _save(monthly, directory / monthly_name)

    # 年度结果复制原始表，仅变更一个值，保留标题、单位合并和总计公式。
    annual_name = "W2_预期_全年更新.xlsx"
    source.cell(target["row"], 3, 99)
    _save(wb, directory / annual_name)
    data = _case(directory)
    data.update({
        "type": "完整流程",
        "artifacts": _artifacts(original_name,
            [f"W2_拆分_{unit_name}.xlsx" for unit_name in dict.fromkeys(item["unit"] for item in records)],
            feedback_name, monthly_name, annual_name),
        "selection": {"target_year": 2026, "target_month": 1,
                      "month_source": "column_header", "month_rule": "exact_month",
                      "snapshot": "after_feedback"},
        "business_premise": "横向月份按列标题读取；本案例提取1月。单位在合并区域内继承，全年保留原始总计。",
        "expected_changes": [{"unit": target_key[0], "task": target_key[1],
                              "field": "1月数量", "old": 18, "new": 99}],
        "invariants": ["除财务部组织安全培训演练1月18→99外，所有任务和月份值保持原始值",
                       "财务部反馈2月为34，不能残留上一轮误改的99",
                       "拆分/反馈行合计引用B:M；年度行合计引用C:N，原始总计公式保留"],
        "summary_policy": "本案例单位拆分不包含跨单位总计；年度更新保留原始总计及公式",
        "ambiguity": [],
    })
    _cj(directory, data)


def _write_W3(path, title, records):
    """暂用双行表头，最终由原生保真模块保留原始前置区并重定位公式。"""
    wb = Workbook()
    sheet = wb.active
    sheet.title = title
    _hdr(sheet, 1, ["编号", "任务名称", "基本信息", "基本信息", "进度", "进度", "进度", "备注"])
    sheet.merge_cells("C1:D1")
    sheet.merge_cells("E1:G1")
    _hdr(sheet, 2, [None, None, "责任单位", "完成时限", "计划数", "完成数", "完成率", None])
    for row, item in enumerate(records, 3):
        done = 15 if item["id"] == "A002" else item["done"]
        sheet.append([item["id"], item["name"], item["unit"], item["deadline"],
                      item["plan"], done, f"=F{row}/E{row}", item["note"]])
    if records:
        sheet.conditional_formatting.add(f"G3:G{sheet.max_row}",
            CellIsRule(operator="lessThan", formula=["0.5"],
                       fill=PatternFill("solid", fgColor="FFC7CE")))
    _save(wb, path)


def rebuild_W3():
    directory = ROOT / "W3"
    original_name = "W3_原始_双行表头.xlsx"
    wb = load_workbook(directory / original_name)
    source = wb.active
    records = []
    for row in range(5, source.max_row + 1):
        if source.cell(row, 1).value:
            values = [source.cell(row, col).value for col in range(1, 9)]
            records.append(dict(zip(("id", "name", "unit", "deadline", "plan", "done", "rate", "note"), values)))
    wb.close()
    by_id = {item["id"]: item for item in records}
    if len(by_id) != len(records) or set(by_id) != {f"A{index:03d}" for index in range(1, 9)}:
        raise ValueError("W3 原始任务应为唯一的A001至A008")
    if by_id["A002"]["done"] != 8:
        raise ValueError("W3 A002完成数的原始基线应为8")
    finance = [item for item in records if item["unit"] == "财务部"]
    if {item["id"] for item in finance} != {"A002", "A006"}:
        raise ValueError("W3 财务部任务集合应为A002/A006")
    march = [item for item in records if str(item["deadline"]).startswith("2026-03-")]
    if [item["id"] for item in march] != ["A003"]:
        raise ValueError("W3 2026年3月按完成时限筛选应仅为A003")
    feedback_name = "W3_反馈_财务部.xlsx"
    monthly_name = "W3_预期_3月汇总.xlsx"
    annual_name = "W3_预期_全年更新.xlsx"
    _write_W3(directory / feedback_name, "财务部", finance)
    _write_W3(directory / monthly_name, "3月汇总", march)
    _write_W3(directory / annual_name, "全年更新", records)

    # 只删除本案例目录中已被明确替代的旧预期，不通配、不递归。
    stale = directory / "W3_预期_月度汇总.xlsx"
    resolved_directory = directory.resolve()
    if stale.resolve().parent != resolved_directory or resolved_directory != (ROOT.resolve() / "W3"):
        raise ValueError("拒绝删除W3案例目录以外的旧预期文件")
    if stale.exists():
        stale.unlink()
    data = _case(directory)
    data.update({
        "type": "完整流程",
        "artifacts": _artifacts(original_name,
            [f"W3_拆分_{unit}.xlsx" for unit in dict.fromkeys(item["unit"] for item in records)],
            feedback_name, monthly_name, annual_name),
        "selection": {"target_year": 2026, "target_month": 3,
                      "month_source": "完成时限", "month_rule": "deadline_month",
                      "snapshot": "after_feedback"},
        "business_premise": "按完成时限筛选2026年3月；财务部反馈只含A002/A006；A002属2月，其变更只影响年度结果，3月只含A003。",
        "expected_changes": [{"task_id": "A002", "unit": "财务部", "field": "完成数",
                              "old": 8, "new": 15}],
        "expected_monthly_ids": {"3月": ["A003"]},
        "invariants": ["A002完成数8→15，其他任务的名称、单位、期限、计划数和完成数不变",
                       "反馈任务集合仅A002/A006；完成率公式引用当前任务行的F/E",
                       "拆分、反馈、月度和年度结果保留原始1至4行，数据从第5行开始"],
        "fidelity_requirements": {
            "前置区和双行表头": "保留原始1至4行及其合并范围，数据从第5行开始",
            "条件格式": "完成率条件格式按任务行映射保留",
            "公式": "完成率=F/E，引用同一任务输出行",
            "冻结": "与原始工作表一致",
        },
    })
    _cj(directory, data)


def rebuild_W6():
    directory = ROOT / "W6"
    original_name = "W6_原始_跨月.xlsx"
    wb = load_workbook(directory / original_name)
    source = wb.active
    records = [[source.cell(row, col).value for col in range(1, 7)]
               for row in range(2, source.max_row + 1) if source.cell(row, 1).value]
    if len(records) != 4 or {item[0] for item in records} != {"C001", "C002", "C003", "C004"}:
        raise ValueError("W6 原始任务应为唯一的C001至C004")
    target = next(item for item in records if item[0] == "C003")
    if target[3:5] != ["3月", "3月"]:
        raise ValueError("W6 C003起始月和结束月的原始基线应均为3月")
    splits = []
    for unit in dict.fromkeys(item[2] for item in records):
        split_name = f"W6_拆分_{unit}.xlsx"
        splits.append(split_name)
        split = Workbook()
        sheet = split.active
        sheet.title = unit
        for col in range(1, 7):
            original_cell = source.cell(1, col)
            target_cell = sheet.cell(1, col, original_cell.value)
            target_cell.font = copy(original_cell.font)
            target_cell.fill = copy(original_cell.fill)
            target_cell.border = copy(original_cell.border)
            target_cell.alignment = copy(original_cell.alignment)
            target_cell.protection = copy(original_cell.protection)
            target_cell.number_format = original_cell.number_format
        for row, item in enumerate((item for item in records if item[2] == unit), 2):
            sheet.append(item)
        _save(split, directory / split_name)
    feedback_name = "W6_反馈_月份调整.xlsx"
    annual_name = "W6_预期_全年更新.xlsx"
    target_row = next(row for row in range(2, source.max_row + 1)
                      if source.cell(row, 1).value == "C003")
    source.cell(target_row, 4, "5月")
    source.cell(target_row, 5, "5月")
    # 保留原始布局；反馈与年度预期均为同一条明确月份调整后的状态。
    wb.save(directory / feedback_name)
    _save(wb, directory / annual_name)
    updated = [list(item) for item in records]
    next(item for item in updated if item[0] == "C003")[3:5] = ["5月", "5月"]
    monthly_name = "W6_预期_月度汇总.xlsx"
    monthly = Workbook()
    sheet = monthly.active
    sheet.title = "月度汇总"
    _hdr(sheet, 1, ["月份", "编号", "任务名称", "责任单位", "数量"])
    for item in sorted(updated, key=lambda item: (int(item[3].removesuffix("月")), item[0])):
        sheet.append([item[3], item[0], item[1], item[2], item[5]])
    _save(monthly, directory / monthly_name)
    data = _case(directory)
    data.update({
        "type": "完整流程",
        "artifacts": _artifacts(original_name, splits, feedback_name, monthly_name, annual_name),
        "selection": {"target_year": 2026, "target_months": [1, 2, 3, 4, 5],
                      "month_source": "起始月", "month_rule": "start_month",
                      "snapshot": "after_feedback"},
        "business_premise": "跨月任务按起始月归入月度汇总；这是本测试假设，非产品最终口径。所有预期均为反馈后快照。",
        "expected_changes": [{"task_id": "C003", "field": field, "old": "3月", "new": "5月"}
                             for field in ("起始月", "结束月")],
        "expected_monthly_ids": {"1月": ["C001"], "2月": ["C002"], "3月": [],
                                 "4月": ["C004"], "5月": ["C003"]},
        "invariants": ["C001/C002/C004全部字段不变，C003数量仍为5",
                       "反馈后C003只出现在5月，3月无任务，不重复计入3月和5月"],
        "ambiguity": [],
    })
    _cj(directory, data)


def fix_D40():
    d = ROOT / "D40"
    if not d.exists():
        return

    wb = Workbook()

    # 原始数据
    ws1 = wb.active; ws1.title = "原始"
    _hdr(ws1, 1, ["A值", "B值", "C值", "D值", "SUM(A:C)"])
    for i in range(5):
        r = i + 2
        a, b, c, dd = 2, 10, 100, 1000
        ws1.append([a, b, c, dd, f"=SUM(A{r}:C{r})"])  # SUM=112

    # 换列：A/D/C/B 顺序，公式仍只计算原A/B/C（即新的A/D/C列）
    ws2 = wb.create_sheet("换列后")
    _hdr(ws2, 1, ["A值", "D值", "C值", "B值", "原SUM内容"])
    for i in range(5):
        r = i + 2
        a, b, c, dd = 2, 10, 100, 1000
        # 换列后：新A=原A(2), 新B=原D(1000), 新C=原C(100), 新D=原B(10)
        # 原SUM(A:C)=2+10+100=112
        # 换列后公式仍需计算 2+10+100=112，即新A+新D+新C
        ws2.append([a, dd, c, b, f"=A{r}+D{r}+C{r}"])  # =2+10+100=112

    _save(wb, d / "D40_列换序引用.xlsx")

    _cj(d, {
        "case_id": "D40", "type": "保真",
        "scenario": "原始列序A/B/C/D，SUM(A:C)=2+10+100=112。换列后A/D/C/B，公式仍需计算原A/B/C",
        "checks": {
            "原始": "SUM(A:C)=A+B+C=2+10+100=112",
            "换列后": "公式=A+D+C=2+10+100=112（新D=原B）",
            "关键": "参与计算的单元格集合不变（原A、原B、原C），只是物理位置变了",
            "预期值": 112,
        },
        "verification_status": "由独立自检计算公式结果并验证为112；Excel/WPS实际重算效果未人工验收",
    })
    print("✓ D40 增加换列含义不变案例")


def update_case_roles():
    """标明结构专项的范围；状态来自独立自检，不保存生成时的通过结论。"""
    for case_file in ROOT.glob("*/case.json"):
        data = json.loads(case_file.read_text(encoding="utf-8"))
        changed = "fidelity_check_results" in data
        data.pop("fidelity_check_results", None)
        if case_file.parent.name in ("W4", "W5"):
            data["type"] = "结构案例"
            data["incomplete_reason"] = "仅验证结构识别/拆分，不作为反馈、月度和年度完整流程金标准"
            for key in ("expected_monthly_summary", "expected_annual_update"):
                data.pop(key, None)
            changed = True
        if case_file.parent.name == "D41":
            data["ambiguity"] = []
            data["business_premise"] = "保留数据区原有空行及其位置，用于维持原始视觉分组；这是已明确的原生保真要求。"
            data["expected"] = "保留空行：任务之间原有的空行保持原位，不自动删除或合并。"
            changed = True
        if changed:
            _cj(case_file.parent, data)


def main():
    rebuild_W2()
    rebuild_W3()
    rebuild_W6()
    fix_D40()
    update_case_roles()
    from 原生保真 import repair
    repair(ROOT)
    print("W2/W3/W6内容预期与原生保真重建完成；请运行独立自检。")


if __name__ == "__main__":
    main()
