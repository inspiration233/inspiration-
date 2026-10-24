"""Product acceptance against existing v3 inputs, not fixture self-validation."""
import hashlib
from pathlib import Path

import pytest
from openpyxl import load_workbook

from engine.tools.excel_merge.preview import preview
from engine.tools.excel_merge.native_split import split_preview, split_export
from engine.tools.excel_merge.reconcile import reconcile_preview, reconcile_export
from engine.tools.excel_merge.workflow import workflow_preview, workflow_export


ROOT = Path(__file__).resolve().parents[1] / "测试集" / "泛用性验收_v3"
BOOKS = sorted(p for p in ROOT.glob("*/*.xlsx") if p.parent.name != "报告")


def spec(path, header, start, end=None, columns=None):
    wb = load_workbook(path)
    ws = wb.active
    result = {"path": str(path), "sheet": ws.title,
              "selection": {"axis": "rows", "header_start": header,
                            "header_end": start - 1, "data_start": start,
                            "data_end": ws.max_row if end is None else end,
                            "field_start": 1, "field_end": columns or ws.max_column}}
    wb.close()
    return result


def values(ws, start):
    return [[c.value for c in row] for row in ws.iter_rows(min_row=start)]


def red(cell):
    fill = cell.fill
    return bool(fill and fill.fill_type == "solid" and str(fill.fgColor.rgb or "")[-6:] == "FF0000")


@pytest.mark.parametrize("path", BOOKS, ids=lambda p: f"{p.parent.name}/{p.name}")
def test_every_v3_workbook_can_be_previewed(path):
    result = preview({"path": str(path), "row_limit": 5, "col_limit": 8})
    assert result["sheets"] and result["sheet"]
    assert result["max_row"] >= 1
    assert result["max_column"] >= 1
    assert isinstance(result["cells"], list)


@pytest.mark.parametrize("case,source,header,start,count", [
    ("W1", "W1_原始计划表.xlsx", 4, 5, 5),
    ("W3", "W3_原始_双行表头.xlsx", 3, 5, 4),
])
def test_native_split_matches_independent_v3_answers(tmp_path, case, source, header, start, count):
    folder = ROOT / case
    payload = {**spec(folder / source, header, start),
               "group_fields": [{"kind": "field", "index": 3}], "out_dir": str(tmp_path)}
    result = split_export(payload)
    assert len(result["files"]) == count
    original = load_workbook(folder / source)
    for entry in result["files"]:
        expected = load_workbook(folder / f"{case}_拆分_{entry['label']}.xlsx")
        actual = load_workbook(entry["path"])
        ws, gold = actual.active, expected.active
        assert values(ws, start) == values(gold, start)
        assert ws.freeze_panes == gold.freeze_panes
        assert str(ws.merged_cells) == str(gold.merged_cells)
        assert values(ws, 1)[:start-1] == values(gold, 1)[:start-1]
        for column in original.active.column_dimensions:
            assert ws.column_dimensions[column].width == original.active.column_dimensions[column].width
        for row in range(1, start):
            for column in range(1, ws.max_column + 1):
                assert ws.cell(row, column)._style == original.active.cell(row, column)._style
        if case == "W1":
            rules = ws.data_validations.dataValidation
            assert rules and all(r.formula1 == '"进行中,已完成,已取消"' for r in rules)
            assert all(any(f"G{r}" in rule for rule in rules) for r in range(start, ws.max_row+1))
            assert len(ws.conditional_formatting) > 0
        expected.close()
        actual.close()
    original.close()


def test_w1_full_feedback_and_june_output(tmp_path):
    folder = ROOT / "W1"
    payload = {**spec(folder / "W1_原始计划表.xlsx", 4, 5),
               "feedback": [spec(p, 4, 5) for p in sorted(folder.glob("W1_反馈_*.xlsx"))],
               "key_fields": [1], "update_fields": [2, 3, 4, 5, 7],
               "compare_fields": [2, 3, 4, 5],
               "filtered_output": {"filters": [{"field": 4, "op": "month", "value": "2026-06"}]},
               "result_form": "direct", "out": str(tmp_path / "全年.xlsx")}
    digest = hashlib.sha256(Path(payload["path"]).read_bytes()).hexdigest()
    report = workflow_preview(payload)
    # 新记录默认自动追加, 不再需要确认
    assert report["can_export"], report["issues"]
    ready = report
    assert len(ready["missing"]) == 5
    assert any(c["key"] == ["T002"] and c["after"] == 25 and c["marked"] for c in ready["changes"])
    assert any(c["key"] == ["T005"] and c["after"] == "已完成" and not c["marked"] for c in ready["changes"])
    result = workflow_export(payload)
    assert result["exported"] and len(result["outputs"]) == 3
    annual = load_workbook(result["output"])
    expected = load_workbook(folder / "W1_预期_全年更新.xlsx")
    assert values(annual.active, 5) == values(expected.active, 5)
    assert red(annual.active["E6"])
    assert red(annual.active["G9"])    # 未选标红列的改动也默认红填充
    monthly = load_workbook(result["filtered_output_path"])
    gold_month = load_workbook(folder / "W1_预期_6月汇总.xlsx")
    assert values(monthly.active, 5) == values(gold_month.active, 5)
    assert monthly.active["F6"].value == "=E6*100"
    assert hashlib.sha256(Path(payload["path"]).read_bytes()).hexdigest() == digest
    for wb in (annual, expected, monthly, gold_month):
        wb.close()


def test_w3_progress_is_updated_without_plan_red(tmp_path):
    folder = ROOT / "W3"
    payload = {**spec(folder / "W3_原始_双行表头.xlsx", 3, 5),
               "feedback": [spec(folder / "W3_反馈_财务部.xlsx", 3, 5)],
               "key_fields": [1], "update_fields": [6], "compare_fields": [2, 3, 4, 5],
               "result_form": "direct", "out": str(tmp_path / "更新.xlsx")}
    report = reconcile_preview(payload)
    assert report["can_export"], report["issues"]
    assert len(report["changes"]) == 1 and not report["changes"][0]["marked"]
    reconcile_export(payload)
    wb = load_workbook(payload["out"])
    # 改动默认红填充(不再要求先选标红列)
    assert wb.active["F6"].value == 15 and red(wb.active["F6"])
    assert wb.active["G6"].value == "=F6/E6"
    assert "C3:D3" in str(wb.active.merged_cells)
    wb.close()


def test_w2_select_january_columns_and_confirm_group_total(tmp_path):
    source = ROOT / "W2" / "W2_原始_横向月份.xlsx"
    payload = {**spec(source, 2, 3, end=14), "blank_group": "previous",
               "group_fields": [{"kind": "field", "index": 1}], "output_fields": [1, 2, 3],
               "out_dir": str(tmp_path)}
    report = split_preview(payload)
    assert [g["count"] for g in report["groups"]] == [4, 4, 4]
    # 合计口径自动按本组保留记录计算, 不再需要确认
    assert report["can_export"] and not report["issues"]
    result = split_export(payload)
    financial = load_workbook(next(f["path"] for f in result["files"] if f["label"] == "财务部"))
    ws = financial.active
    assert ws.max_column == 3 and ws.max_row == 7
    assert ws["C2"].value == "1月"
    assert [ws.cell(r, 3).value for r in range(3, 7)] == [18, 17, 30, 27]
    assert ws["A4"].value is None  # attribution is inherited without rewriting blank cells
    assert ws["C7"].value == "=SUM(C3:C6)"
    financial.close()


def feedback_case(case, tmp_path):
    folder = ROOT / case
    base = spec(folder / "原始.xlsx", 1, 2)
    return {**base, "feedback": [spec(folder / "反馈.xlsx", 1, 2)], "key_fields": [1],
            "update_fields": list(range(2, base["selection"]["field_end"] + 1)),
            "compare_fields": list(range(2, base["selection"]["field_end"] + 1)),
            "result_form": "direct", "out": str(tmp_path / f"{case}.xlsx")}


def test_c32_changed_name_unit_and_month_are_matched_by_selected_id(tmp_path):
    payload = feedback_case("C32", tmp_path)
    result = reconcile_export(payload)
    assert result["exported"] and len(result["changes"]) == 4
    actual = load_workbook(payload["out"])
    feedback = load_workbook(payload["feedback"][0]["path"])
    assert values(actual.active, 1) == values(feedback.active, 1)
    assert all(red(actual.active[ref]) for ref in ("B2", "D3", "C4", "F5"))
    actual.close()
    feedback.close()


def test_c33_addition_requires_confirmation_and_missing_is_kept(tmp_path):
    payload = feedback_case("C33", tmp_path)
    before = reconcile_preview(payload)
    assert before["can_export"]          # 新记录默认自动追加
    assert [r["key"] for r in before["missing"]] == [["B003"]]
    result = reconcile_export(payload)
    assert result["exported"]
    actual = load_workbook(payload["out"])
    assert [actual.active.cell(r, 1).value for r in range(2, 8)] == [
        "B001", "B002", "B003", "B004", "B005", "NEW01"]
    assert actual.active["F4"].value == "原有备注"
    actual.close()


def test_c35_cancellation_text_is_preserved_and_missing_does_not_delete(tmp_path):
    payload = feedback_case("C35", tmp_path)
    result = reconcile_export(payload)
    assert result["exported"] and [r["key"] for r in result["missing"]] == [["B004"]]
    actual = load_workbook(payload["out"])
    assert actual.active.max_row == 6
    assert actual.active["F4"].value == "该任务已取消" and red(actual.active["F4"])
    assert actual.active["A5"].value == "B004"
    actual.close()


def test_c36_duplicate_uuid_and_empty_id_cannot_silently_match(tmp_path):
    payload = feedback_case("C36", tmp_path)
    payload["key_fields"] = [5]
    result = reconcile_preview(payload)
    assert not result["can_export"]
    assert result["issues"]
    assert not Path(payload["out"]).exists()


def test_c37_duplicate_import_does_not_duplicate_records(tmp_path):
    payload = feedback_case("C37", tmp_path)
    payload["feedback"] *= 2
    result = reconcile_export(payload)
    assert result["exported"] and result["changes"] == []
    actual = load_workbook(payload["out"])
    assert actual.active.max_row == 4
    actual.close()


def test_w6_interval_filter_uses_every_covered_month_without_duplicate_annual_rows(tmp_path):
    # The product's selected interval rule is distinct from v3's start-month-only gold file.
    folder = ROOT / "W6"
    payload = {**spec(folder / "W6_原始_跨月.xlsx", 1, 2), "group_fields": [],
               "filters": [{"field": 4, "end_field": 5, "op": "interval", "value": "2026-03"}]}
    before = split_preview(payload)
    assert before["can_export"], before["issues"]
    assert before["groups"][0]["records"] == [2, 3, 4]
    update = {**spec(folder / "W6_原始_跨月.xlsx", 1, 2),
              "feedback": [spec(folder / "W6_反馈_月份调整.xlsx", 1, 2)],
              "key_fields": [1], "update_fields": [4, 5], "compare_fields": [4, 5],
              "result_form": "direct", "out": str(tmp_path / "全年.xlsx"),
              "filtered_output": {"filters": payload["filters"]}}
    result = workflow_export(update)
    assert result["exported"], result["issues"]
    annual = load_workbook(result["out"])
    monthly = load_workbook(result["filtered_output_path"])
    assert annual.active.max_row == 5
    assert sum(annual.active.cell(r, 6).value for r in range(2, 6)) == 17
    assert [monthly.active.cell(r, 1).value for r in range(2, monthly.active.max_row + 1)] == ["C001", "C002"]
    assert red(annual.active["D4"]) and red(annual.active["E4"])
    annual.close()
    monthly.close()


def test_rt42_split_feedback_roundtrip_preserves_every_cell(tmp_path):
    original = ROOT / "RT42" / "RT42_原始.xlsx"
    base = spec(original, 1, 2)
    split = split_export({**base, "group_fields": [{"kind": "field", "index": 3}],
                          "out_dir": str(tmp_path / "拆分")})
    payload = {**base, "feedback": [spec(f["path"], 1, 2) for f in split["files"]],
               "key_fields": [1], "update_fields": [2, 3, 4, 5], "compare_fields": [2, 3, 4, 5],
               "result_form": "direct", "out": str(tmp_path / "往返.xlsx")}
    result = reconcile_export(payload)
    assert result["exported"] and result["changes"] == [] and result["missing"] == []
    source, actual = load_workbook(original), load_workbook(payload["out"])
    assert values(actual.active, 1) == values(source.active, 1)
    for row in source.active:
        for cell in row:
            assert actual.active[cell.coordinate]._style == cell._style
    source.close()
    actual.close()
