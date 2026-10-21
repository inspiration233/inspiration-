# -*- coding: utf-8 -*-
"""修复 W1/W3 夹具的原生版式；不调用产品引擎，也不推断业务变更。

在全部内容生成步骤之后调用 ``repair(ROOT)``。当前文件中的任务 ID、顺序和
业务字段是输入；只重建版式，以及案例已经约定的预算/完成率派生公式。
本模块不生成验收通过记录，独立自检必须重新读取输出。
"""
from copy import copy, deepcopy
import json
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.cell_range import MultiCellRange


SPECS = {
    "W1": {
        "original": "W1_原始计划表.xlsx",
        "id_pattern": r"T\d{3}",
        "width": 7,
        "formula_column": 6,
        "formula": "=E{row}*100",
        "new_task_format_id": "T001",
        "files": [
            "W1_拆分_综合办公室.xlsx", "W1_拆分_财务部.xlsx",
            "W1_拆分_工程技术部.xlsx", "W1_拆分_安全管理部.xlsx",
            "W1_拆分_市场营销部.xlsx", "W1_反馈_财务部.xlsx",
            "W1_反馈_工程技术部.xlsx", "W1_反馈_市场营销部.xlsx",
            "W1_预期_6月汇总.xlsx", "W1_预期_全年更新.xlsx",
        ],
    },
    "W3": {
        "original": "W3_原始_双行表头.xlsx",
        "id_pattern": r"A\d{3}",
        "width": 8,
        "formula_column": 7,
        "formula": "=F{row}/E{row}",
        "new_task_format_id": "A001",
        "files": [
            "W3_拆分_综合办公室.xlsx", "W3_拆分_财务部.xlsx",
            "W3_拆分_工程技术部.xlsx", "W3_拆分_安全管理部.xlsx",
            "W3_反馈_财务部.xlsx", "W3_预期_3月汇总.xlsx",
            "W3_预期_全年更新.xlsx",
        ],
    },
}


def _records(ws, spec):
    """按夹具既定 ID 读取，兼容修复前后的表头位置；拒绝重复 ID。"""
    records = []
    seen = set()
    for row in ws.iter_rows():
        value = row[0].value
        if not isinstance(value, str) or not re.fullmatch(spec["id_pattern"], value):
            continue
        if value in seen:
            raise ValueError(f"{ws.title}: 重复任务 ID {value}")
        seen.add(value)
        if any(c.value is not None for c in row[spec["width"]:]):
            raise ValueError(f"{ws.title}/{value}: 有既定字段以外的数据，拒绝丢弃")
        values = [ws.cell(row[0].row, col).value
                  for col in range(1, spec["width"] + 1)]
        records.append((value, row[0].row, values))
    return records


def _business_values(records, formula_column):
    """派生公式之外的字段必须逐项不变，包括空值、取消状态与备注。"""
    return [
        (task_id, tuple(v for col, v in enumerate(values, start=1)
                        if col != formula_column))
        for task_id, _row, values in records
    ]


def _copy_row_style(source, source_row, target, target_row, width):
    # target 从同一原文件重新加载，样式/主题表与 source 一致。
    for col in range(1, width + 1):
        src = source.cell(source_row, col)
        dst = target.cell(target_row, col)
        dst._style = copy(src._style)
        dst.comment = copy(src.comment) if src.comment else None
        dst.hyperlink = copy(src.hyperlink) if src.hyperlink else None
    if source_row in source.row_dimensions:
        dim = copy(source.row_dimensions[source_row])
        dim.index = target_row
        dim.parent = target
        target.row_dimensions[target_row] = dim


def _restore_rules(source, target, last_row):
    """这些夹具只含 G 列整段规则；复制完整规则属性，仅重定数据范围。"""
    source_ref = f"G5:G{source.max_row}"
    output_ref = f"G5:G{last_row}" if last_row >= 5 else None
    target.data_validations = deepcopy(source.data_validations)
    target.data_validations.dataValidation = []
    target.conditional_formatting = ConditionalFormattingList()

    for dv in source.data_validations.dataValidation:
        if str(dv.sqref) != source_ref:
            raise ValueError(f"{source.title}: 未约定的数据验证范围 {dv.sqref}")
        if output_ref:
            rule = deepcopy(dv)
            rule.sqref = MultiCellRange(output_ref)
            target.add_data_validation(rule)

    for cf in source.conditional_formatting:
        if str(cf.sqref) != source_ref:
            raise ValueError(f"{source.title}: 未约定的条件格式范围 {cf.sqref}")
        if output_ref:
            for rule in source.conditional_formatting[cf]:
                target.conditional_formatting.add(output_ref, deepcopy(rule))
    return output_ref


def _rule_contract(source):
    return {
        "data_validations": [
            {
                "type": dv.type, "operator": dv.operator,
                "formula1": dv.formula1, "formula2": dv.formula2,
                "allow_blank": dv.allowBlank,
                "show_error_message": dv.showErrorMessage,
                "source_range": str(dv.sqref),
            }
            for dv in source.data_validations.dataValidation
        ],
        "conditional_formats": [
            {
                "type": rule.type, "operator": rule.operator,
                "formula": list(rule.formula or []),
                "source_range": str(cf.sqref),
                "fill_rgb": (
                    rule.dxf.fill.fgColor.rgb
                    if rule.dxf is not None and rule.dxf.fill is not None
                    and rule.dxf.fill.fgColor.type == "rgb" else None
                ),
            }
            for cf in source.conditional_formatting
            for rule in source.conditional_formatting[cf]
        ],
    }


def _repair_file(source_path, source, source_rows, path, spec):
    current = load_workbook(path, data_only=False)
    try:
        if len(current.worksheets) != 1:
            raise ValueError(f"{path.name}: 本夹具要求单工作表，拒绝丢弃其他页")
        sheet_name = current.active.title
        records = _records(current.active, spec)
        before = _business_values(records, spec["formula_column"])
    finally:
        current.close()

    # 以原始工作簿为底稿保留完整样式表、主题、标题、打印和工作表属性。
    # 只替换第 5 行起的数据，不从简化输出反向猜测原有版式。
    output = load_workbook(source_path, data_only=False)
    try:
        target = output.active
        target.title = sheet_name
        if target.max_row >= 5:
            target.delete_rows(5, target.max_row - 4)
        for row_index in list(target.row_dimensions):
            if row_index >= 5:
                del target.row_dimensions[row_index]

        style_sources = {}
        for out_row, (task_id, _old_row, values) in enumerate(records, start=5):
            style_id = task_id if task_id in source_rows else spec["new_task_format_id"]
            source_row = source_rows[style_id]
            _copy_row_style(source, source_row, target, out_row, spec["width"])
            for col, value in enumerate(values, start=1):
                target.cell(out_row, col).value = value
            target.cell(out_row, spec["formula_column"]).value = spec["formula"].format(
                row=out_row
            )
            style_sources[task_id] = style_id

        last_row = 4 + len(records)
        rule_range = _restore_rules(source, target, last_row)
        after = _business_values(_records(target, spec), spec["formula_column"])
        if after != before:
            raise ValueError(f"{path.name}: 版式修复意外改变业务字段")
        output.save(path)
        last_col = chr(ord("A") + spec["width"] - 1)
        return {
            "file": path.name,
            "sheet": sheet_name,
            "header_range": f"A1:{last_col}4",
            "data_range": f"A5:{last_col}{last_row}" if records else None,
            "task_ids": [task_id for task_id, _, _ in records],
            "style_source_ids": style_sources,
            "data_validation_ranges": (
                [rule_range] * len(source.data_validations.dataValidation)
                if records else []
            ),
            "conditional_format_ranges": (
                [rule_range] * len(source.conditional_formatting)
                if records else []
            ),
        }
    finally:
        output.close()


def repair(root=None):
    """最后阶段原地修复已生成的 W1/W3；输出可读契约，不写入通过快照。"""
    root = Path(root) if root is not None else Path(__file__).resolve().parent
    contracts = {}
    for case_id, spec in SPECS.items():
        case_dir = root / case_id
        case_path = case_dir / "case.json"
        case = json.loads(case_path.read_text(encoding="utf-8"))
        source_path = case_dir / spec["original"]
        source_book = load_workbook(source_path, data_only=False)
        try:
            if len(source_book.worksheets) != 1:
                raise ValueError(f"{spec['original']}: 夹具原表应只有一个工作表")
            source = source_book.active
            if any(m.max_row >= 5 for m in source.merged_cells.ranges):
                raise ValueError(f"{case_id}: 夹具新增数据区合并，需明确映射再修复")
            source_rows = {task_id: row for task_id, row, _ in _records(source, spec)}
            if spec["new_task_format_id"] not in source_rows:
                raise ValueError(f"{case_id}: 新任务格式参考行不存在")
            missing = [name for name in spec["files"] if not (case_dir / name).is_file()]
            if missing:
                raise FileNotFoundError(f"{case_id}: 请先完成内容生成：{missing}")

            file_contracts = [
                _repair_file(source_path, source, source_rows, case_dir / name, spec)
                for name in spec["files"]
            ]
            contract = {
                "version": 1,
                "source_file": spec["original"],
                "source_sheet": source.title,
                "preserved_header_rows": [1, 2, 3, 4],
                "data_start_row": 5,
                "header_merges": sorted(str(m) for m in source.merged_cells.ranges),
                "freeze_panes": source.freeze_panes,
                "column_dimensions": "全部与原表一致（含默认值、宽度、隐藏及分组属性）",
                "row_dimensions": "前4行原样；已有任务按ID取原表行，新任务按指定参考ID",
                "cell_styles": "表头和数据行全部样式属性按原表复制，含字体/填充/数字格式",
                "print_settings": "与原表一致，含页面设置、页边距、页眉页脚及打印范围",
                "derived_formula_column": spec["formula_column"],
                "derived_formula": spec["formula"],
                "new_task_format_source_id": spec["new_task_format_id"],
                "empty_data_policy": "只保留原表前4行，不创建空的数据验证或条件格式范围",
                "rules": _rule_contract(source),
                "files": file_contracts,
            }
            contracts[case_id] = contract
            case.pop("fidelity_check_results", None)
            if case_id == "W1":
                case["ground_truth"]["june_tasks"] = 1
                case["expected_monthly_ids"] = {"6月": ["T006", "T999"]}
                case["selection"] = {
                    "target_year": 2026, "target_month": 6,
                    "month_rule": "exact_month", "month_source": "月份",
                    "snapshot": "after_feedback",
                }
                case.pop("corrections", None)
                case["artifacts"] = {
                    "original": [spec["original"]],
                    "split": [n for n in spec["files"] if "_拆分_" in n],
                    "feedback": [n for n in spec["files"] if "_反馈_" in n],
                    "expected_monthly": ["W1_预期_6月汇总.xlsx"],
                    "expected_annual": ["W1_预期_全年更新.xlsx"],
                }
            case["fidelity_contract"] = contract
            case["fidelity_requirements"] = {
                "前置区和表头": "完整保留原表第1-4行，数据统一从第5行开始",
                "样式": "全部源属性按原任务ID复制；新增任务使用契约指定的参考行",
                "下拉和条件格式": "复制原表完整规则，范围仅覆盖实际数据行；不得重复",
                "冻结": f"与原表一致：{source.freeze_panes!r}",
                "公式": spec["formula"] + "，row 为每条任务的实际输出行",
            }
            case_path.write_text(
                json.dumps(case, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
        finally:
            source_book.close()
    return contracts


if __name__ == "__main__":
    repair()
