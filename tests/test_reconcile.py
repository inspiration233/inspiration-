"""Business invariants for reconciliation, using independent temporary workbooks."""
from copy import copy, deepcopy
import json
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import warnings

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from engine.tools.excel_merge.merge import Cancelled
from engine.tools.excel_merge.reconcile import reconcile_export, reconcile_preview


def _add_extension(path, part, feature=None, encoding="utf-8"):
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    xml = ET.fromstring(parts[part])
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    container = xml.find("{" + main + "}extLst")
    if container is None:
        container = ET.SubElement(xml, "{" + main + "}extLst")
    if feature is not None:
        extension = ET.SubElement(container, "{" + main + "}ext", uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}")
        if feature:
            ET.SubElement(extension, "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}" + feature)
    parts[part] = ET.tostring(xml, encoding=encoding)
    with ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def book(path, records, headers=("编号", "单位", "计划", "完成", "状态"), axis="rows",
         native=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "计划"
    table = [list(headers)] + [list(r) for r in records]
    if axis == "columns":
        table = list(map(list, zip(*table)))
    for row in table:
        ws.append(row)
    if native:
        ws.freeze_panes = "C2"
        ws.column_dimensions["C"].width = 27
        ws.row_dimensions[2].height = 35
        ws["C2"].font = Font(name="宋体", size=13, bold=True, italic=True, color="FF112233")
        ws["C2"].number_format = "0.00"
        ws["C2"].fill = PatternFill("solid", fgColor="FFFFCC")
        ws.page_setup.orientation = "landscape"
        ws.page_margins.left = 0.4
        ws.print_title_rows = "1:1"
        ws.print_area = f"A1:E{len(records)+1}"
        validation = DataValidation(type="list", formula1='"进行中,已完成,已取消"')
        validation.add(f"E2:E{len(records)+1}")
        ws.add_data_validation(validation)
        ws.conditional_formatting.add(
            f"D2:D{len(records)+1}", CellIsRule(operator="greaterThan", formula=["5"],
                                              fill=PatternFill("solid", fgColor="C6EFCE")))
        other = wb.create_sheet("说明")
        other["A1"] = "保留此页"
    wb.save(path)
    wb.close()
    return path


def selection(records, fields=5, axis="rows"):
    return {"axis": axis, "header_start": 1, "header_end": 1,
            "data_start": 2, "data_end": records + 1, "field_start": 1, "field_end": fields}


def payload(tmp_path, current_records, feedback_records, axis="rows", native=False):
    current = book(tmp_path / "current.xlsx", current_records, axis=axis, native=native)
    feedback = book(tmp_path / "feedback.xlsx", feedback_records, axis=axis)
    return {"path": str(current), "sheet": "计划", "selection": selection(len(current_records), axis=axis),
            "feedback": [{"path": str(feedback), "sheet": "计划",
                          "selection": selection(len(feedback_records), axis=axis)}],
            "key_fields": [1], "update_fields": [3, 4, 5], "compare_fields": [3, 5],
            "out": str(tmp_path / "result.xlsx")}


def resolve(p, preview, kind, value):
    p.setdefault("resolutions", {})
    for issue in preview["issues"]:
        if issue["kind"] == kind:
            p["resolutions"][issue["id"]] = value


@pytest.mark.parametrize("part,feature,encoding", [("xl/styles.xml", "slicerStyles", "utf-8"),
                                         ("xl/worksheets/sheet1.xml", "conditionalFormattings", "utf-8"),
                                         ("xl/worksheets/sheet1.xml", "futureCellData", "utf-8"),
                                         ("xl/workbook.xml", "futureWorkbookFeature", "utf-8"),
                                         ("xl/workbook.xml", "futureWorkbookFeature", "utf-16")])
def test_extensions_report_notice_without_blocking_updates(tmp_path, part, feature, encoding):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")],
                [("T1", "甲", 12, 3, "进行中")], native=True)
    current = load_workbook(p["path"])
    current["计划"]["F1"] = "公式保留"
    current["计划"]["F2"] = "=C2*2"
    current.save(p["path"]); current.close()
    _add_extension(p["path"], part, feature, encoding)
    original = Path(p["path"]).read_bytes()
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("error", UserWarning)
        preview = reconcile_preview(p)
        result = reconcile_export(p)
    assert not caught
    assert preview["can_export"] and result["exported"], result
    assert any("扩展" in notice for notice in preview["warnings"])
    assert result["warnings"] == preview["warnings"]
    actual = load_workbook(p["out"])
    ws = actual["计划"]
    assert ws["C2"].value == 12 and ws["D2"].value == 3
    assert ws["F2"].value == "=C2*2"
    assert ws["C2"].font.name == "宋体" and ws["C2"].font.bold and ws["C2"].font.italic
    assert ws["C2"].font.color.rgb == "FFFF0000" and ws["C2"].number_format == "0.00"
    assert ws["C2"].fill.fgColor.rgb == "00FFFFCC"
    assert ws.page_setup.orientation == "landscape" and ws.page_margins.left == 0.4
    assert len(ws.data_validations.dataValidation) == 1 and len(ws.conditional_formatting) == 1
    assert Path(p["path"]).read_bytes() == original
    log = json.loads(Path(result["changes_path"]).read_text(encoding="utf-8"))
    assert log["warnings"] == result["warnings"]
    actual.close()


def test_empty_extensions_pass_and_extension_removal_warning_does_not_block(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 12, 3, "进行中")])
    _add_extension(p["path"], "xl/worksheets/sheet1.xml")
    preview = reconcile_preview(p)
    assert preview["can_export"] and not preview["warnings"], preview
    _add_extension(p["path"], "xl/worksheets/sheet1.xml", "")
    preview = reconcile_preview(p)
    assert preview["can_export"], preview
    assert any("extension is not supported and will be removed" in x for x in preview["warnings"])
    assert reconcile_export(p)["exported"]
    actual = load_workbook(p["out"])
    assert actual["计划"]["C2"].value == 12
    actual.close()


def test_printer_binary_is_notice_and_keeps_standard_page_setup(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")],
                [("T1", "甲", 12, 3, "进行中")], native=True)
    with ZipFile(p["path"], "a") as archive:
        archive.writestr("xl/printerSettings/printerSettings1.bin", b"printer-specific-settings")
    result = reconcile_export(p)
    assert result["exported"] and any("打印机" in x for x in result["warnings"]), result
    actual = load_workbook(p["out"])
    assert actual["计划"].page_setup.orientation == "landscape"
    assert actual["计划"].page_margins.left == 0.4
    with ZipFile(p["out"]) as archive:
        assert not any("printerSettings" in x for x in archive.namelist())
    actual.close()


def test_style_permission_does_not_allow_unrelated_read_warning(tmp_path, monkeypatch):
    from engine.tools.excel_merge import reconcile
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 12, 3, "进行中")])
    _add_extension(p["path"], "xl/styles.xml", "slicerStyles")
    original_load = reconcile.load_workbook
    def warned_load(*args, **kwargs):
        warnings.warn("Unreadable formula dependency", UserWarning)
        return original_load(*args, **kwargs)
    monkeypatch.setattr(reconcile, "load_workbook", warned_load)
    preview = reconcile_preview(p)
    assert not preview["can_export"]
    assert any("Unreadable formula dependency" in issue["message"] for issue in preview["issues"])


def test_updates_missing_and_native_preservation(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中"), ("T2", "乙", 20, 2, "进行中")],
                [("T1", "甲", 12, 8, "已取消")], native=True)
    before = Path(p["path"]).read_bytes()
    preview = reconcile_preview(p)
    assert preview["can_export"], preview
    assert preview["missing"] == [{"key": ["T2"], "position": 3, "status": "not_received",
                                   "action": "keep_current", "message": "本次未反馈，保留当前内容"}]
    assert {c["field"] for c in preview["changes"]} == {3, 4, 5}
    assert next(c for c in preview["changes"] if c["field"] == 5)["kind"] == "cancel"
    result = reconcile_export(p)
    assert result["exported"], result
    assert Path(p["path"]).read_bytes() == before
    wb = load_workbook(p["out"])
    ws = wb["计划"]
    assert [ws.cell(2, c).value for c in range(1, 6)] == ["T1", "甲", 12, 8, "已取消"]
    assert ws["C3"].value == 20 and ws["D3"].value == 2
    assert ws["C2"].font.color.rgb == "FFFF0000"
    assert ws["C2"].font.name == "宋体" and ws["C2"].font.bold and ws["C2"].font.italic
    assert ws["D2"].font.color.type != "rgb" or ws["D2"].font.color.rgb != "FFFF0000"
    assert ws["C2"].number_format == "0.00" and ws.column_dimensions["C"].width == 27
    assert ws.row_dimensions[2].height == 35 and ws.freeze_panes == "C2"
    assert ws.page_setup.orientation == "landscape" and ws.page_margins.left == 0.4
    assert ws.print_title_rows == "$1:$1"
    assert str(ws.data_validations.dataValidation[0].sqref) == "E2:E3"
    assert len(ws.conditional_formatting) == 1 and wb["说明"]["A1"].value == "保留此页"
    wb.close()
    audit = json.loads(Path(result["changes_path"]).read_text(encoding="utf-8"))
    assert audit["changes"] == result["changes"]


def test_default_requires_update_fields_and_no_compare_means_no_red(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 8, "进行中")])
    p.pop("update_fields")
    assert not reconcile_preview(p)["can_export"]
    p["update_fields"] = [3, 4]
    p.pop("compare_fields")
    result = reconcile_export(p)
    assert result["exported"] and not any(c["marked"] for c in result["changes"])
    wb = load_workbook(p["out"])
    assert wb.active["C2"].font.color.type == "theme"
    wb.close()


def test_column_axis_and_composite_keys(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中"), ("T1", "乙", 20, 0, "进行中")],
                [("T1", "乙", 25, 5, "已完成")], axis="columns")
    p["key_fields"] = [1, 2]
    result = reconcile_export(p)
    assert result["exported"], result
    wb = load_workbook(p["out"])
    ws = wb.active
    assert ws["B3"].value == 10 and ws["C3"].value == 25 and ws["C4"].value == 5
    assert ws["C3"].font.color.rgb == "FFFF0000"
    wb.close()


@pytest.mark.parametrize("role", ["current", "feedback"])
@pytest.mark.parametrize("bad", [None, "duplicate"])
def test_empty_and_duplicate_identity_never_pick_first_or_last(tmp_path, role, bad):
    current = [("T1", "甲", 10, 0, "进行中"), ("T2", "乙", 20, 0, "进行中")]
    feedback = [("T1", "甲", 11, 1, "进行中"), ("T2", "乙", 21, 1, "进行中")]
    records = current if role == "current" else feedback
    records[1] = (None if bad is None else "T1", *records[1][1:])
    p = payload(tmp_path, current, feedback)
    preview = reconcile_preview(p)
    assert not preview["can_export"]
    assert any(i["kind"] == ("empty_key" if bad is None else "duplicate_key") for i in preview["issues"])
    assert not reconcile_export(p)["exported"]
    assert not Path(p["out"]).exists()


def test_conflicting_feedback_needs_explicit_source_or_keep(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    second = book(tmp_path / "other.xlsx", [("T1", "甲", 12, 2, "进行中")])
    p["feedback"].append({"path": str(second), "selection": selection(1)})
    preview = reconcile_preview(p)
    issue = next(i for i in preview["issues"] if i["kind"] == "conflict")
    assert not preview["can_export"] and len(issue["choices"]) == 3
    resolve(p, preview, "conflict", "keep")
    kept = reconcile_preview(p)
    assert kept["can_export"] and kept["changes"] == []
    choice = next(c["value"] for c in issue["choices"] if "other.xlsx" in c["label"])
    resolve(p, preview, "conflict", choice)
    result = reconcile_export(p)
    assert result["exported"]
    wb = load_workbook(p["out"])
    assert wb.active["C2"].value == 12
    wb.close()


def test_duplicate_import_is_idempotent(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    single = reconcile_preview(p)
    p["feedback"].append(deepcopy(p["feedback"][0]))
    repeated = reconcile_preview(p)
    assert repeated["can_export"] and repeated["changes"] == single["changes"]
    assert repeated["totals"]["duplicate_imports"] == 1


def test_addition_requires_confirmation_and_extends_rules_and_template(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")],
                [("T1", "甲", 10, 0, "进行中"), ("T2", "乙", 15, 1, "进行中")], native=True)
    preview = reconcile_preview(p)
    assert not preview["can_export"]
    resolve(p, preview, "addition", "append")
    result = reconcile_export(p)
    assert result["exported"], result
    wb = load_workbook(p["out"])
    ws = wb["计划"]
    assert ws["A3"].value == "T2" and ws["B3"].value == "乙" and ws["C3"].value == 15
    assert ws["C3"].font.bold and ws["C3"].number_format == "0.00"
    assert str(ws.data_validations.dataValidation[0].sqref) == "E2:E3"
    assert str(next(iter(ws.conditional_formatting)).sqref) == "D2:D3"
    assert str(ws.print_area) == "'计划'!$A$1:$E$3"
    wb.close()
    assert any(c["kind"] == "new" for c in result["changes"])


def test_addition_with_footer_can_only_skip(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中"), ("页脚", None, None, None, None)],
                [("T2", "乙", 15, 1, "进行中")])
    p["selection"]["data_end"] = 2
    preview = reconcile_preview(p)
    issue = next(i for i in preview["issues"] if i["kind"] == "addition")
    assert [c["value"] for c in issue["choices"]] == ["skip"]
    resolve(p, preview, "addition", "skip")
    assert reconcile_export(p)["exported"]


def test_reordered_columns_and_explicit_mapping(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    feedback = Path(p["feedback"][0]["path"])
    book(feedback, [(8, "T1", "已完成", 14, "甲")], ("完成", "编号", "状态", "计划", "单位"))
    preview = reconcile_preview(p)
    assert preview["can_export"]
    assert next(c for c in preview["changes"] if c["field"] == 3)["after"] == 14
    book(feedback, [(8, "T1", "已完成", 14, "甲")], ("完成", "编号", "状态", "改名计划", "单位"))
    assert not reconcile_preview(p)["can_export"]
    p["feedback"][0]["field_map"] = {"3": 4}
    assert reconcile_export(p)["exported"]


def test_duplicate_header_requires_manual_mapping(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    book(Path(p["feedback"][0]["path"]), [("T1", "甲", 11, 1, "进行中")],
         ("编号", "单位", "计划", "计划", "状态"))
    p["update_fields"] = [3]
    preview = reconcile_preview(p)
    assert not preview["can_export"]
    p["feedback"][0]["field_map"] = {"3": 3}
    assert reconcile_preview(p)["can_export"]


def test_baseline_and_relocated_formula_semantics(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, "=C2*2", "进行中"), ("T2", "乙", 20, "=C3*2", "进行中")],
                [("T2", "乙", 25, "=C2*2", "进行中")])
    baseline = book(tmp_path / "baseline.xlsx",
                    [("T2", "乙", 18, "=C2*2", "进行中"), ("T1", "甲", 10, "=C3*2", "进行中")])
    p["baseline"] = {"path": str(baseline), "selection": selection(2)}
    p["compare_fields"] = [3, 4]
    preview = reconcile_preview(p)
    assert preview["can_export"], preview
    assert len(preview["changes"]) == 1
    assert preview["changes"][0]["before"] == 18 and preview["changes"][0]["after"] == 25
    assert reconcile_export(p)["exported"]
    wb = load_workbook(p["out"])
    assert wb.active["D3"].value == "=C3*2"
    wb.close()


def test_existing_red_is_not_evidence_of_change(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 10, 0, "进行中")])
    wb = load_workbook(p["path"])
    wb.active["C2"].font = Font(color="FFFF0000", bold=True)
    wb.save(p["path"])
    wb.close()
    assert reconcile_preview(p)["changes"] == []


@pytest.mark.parametrize("object_name", ["xl/embeddings/object1.bin", "xl/drawings/drawing1.xml",
                                      "xl/vbaProject.bin", "xl/slicers/slicer1.xml"])
def test_unsupported_objects_are_rejected_without_silent_loss(tmp_path, object_name):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    with ZipFile(p["path"], "a") as archive:
        archive.writestr(object_name, b"unsupported")
    result = reconcile_export(p)
    assert not result["exported"] and not result["can_export"]
    assert not Path(p["out"]).exists()


def test_no_overwrite_non_xlsx_and_cancellation(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    p["out"] = p["path"]
    assert not reconcile_export(p)["exported"]
    p["out"] = str(tmp_path / "new.xlsx")
    with pytest.raises(Cancelled):
        reconcile_export(p, cancel=lambda: True)
    assert not Path(p["out"]).exists()
    p["path"] = str(tmp_path / "unsupported.xls")
    assert not reconcile_preview(p)["can_export"]


def test_multilevel_header_paths_distinguish_same_leaf_titles(tmp_path):
    current = tmp_path / "current.xlsx"
    feedback = tmp_path / "feedback.xlsx"
    for path, planned, actual in ((current, 10, 1), (feedback, 12, 5)):
        wb = Workbook()
        ws = wb.active
        ws.append(["任务", "计划", "进展"])
        ws.append(["编号", "数量", "数量"])
        ws.append(["T1", planned, actual])
        wb.save(path)
        wb.close()
    p = {"path": str(current), "selection": {"axis": "rows", "header_start": 1, "header_end": 2,
         "data_start": 3, "data_end": 3, "field_start": 1, "field_end": 3},
         "feedback": [{"path": str(feedback)}], "key_fields": [1], "update_fields": [2, 3],
         "compare_fields": [2], "out": str(tmp_path / "result.xlsx")}
    preview = reconcile_preview(p)
    assert preview["can_export"], preview
    assert preview["fields"][1]["path"] == ["计划", "数量"]
    assert reconcile_export(p)["exported"]


def test_feedback_selection_defaults_to_own_extent(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中"), ("T2", "乙", 20, 0, "进行中")],
                [("T2", "乙", 22, 1, "进行中")])
    p["feedback"][0].pop("selection")
    assert reconcile_preview(p)["can_export"]
    assert reconcile_export(p)["exported"]


@pytest.mark.parametrize("axis", ["rows", "columns"])
def test_new_tasks_preserve_context_and_relocate_formula_for_both_axes(tmp_path, axis):
    formula_old = "=C2*2" if axis == "rows" else "=B3*2"
    formula_new = "=C3*2" if axis == "rows" else "=C3*2"
    p = payload(tmp_path, [("T1", "甲", 10, formula_old, "进行中")],
                [("T1", "甲", 10, formula_old, "进行中"), ("T2", "乙", 20, formula_new, "进行中")],
                axis=axis)
    preview = reconcile_preview(p)
    resolve(p, preview, "addition", "append")
    result = reconcile_export(p)
    assert result["exported"], result
    wb = load_workbook(p["out"])
    ws = wb.active
    assert ws.cell(3, 2).value == "乙" if axis == "rows" else ws.cell(2, 3).value == "乙"
    assert ws.cell(3, 4).value == "=C3*2" if axis == "rows" else ws.cell(4, 3).value == "=C3*2"
    wb.close()


def test_merged_unit_can_participate_in_composite_identity(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中"), ("T2", "甲", 20, 0, "进行中")],
                [("T2", "甲", 22, 1, "进行中")])
    wb = load_workbook(p["path"])
    wb.active.merge_cells("B2:B3")
    wb.save(p["path"])
    wb.close()
    p["key_fields"] = [1, 2]
    result = reconcile_export(p)
    assert result["exported"], result
    wb = load_workbook(p["out"])
    assert wb.active["C3"].value == 22 and "B2:B3" in wb.active.merged_cells
    wb.close()


def test_formula_in_permuted_fields_blocks_instead_of_summing_wrong_column(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, "=C2*2", "进行中")],
                [("T1", "甲", 10, "=C2*2", "进行中")])
    book(Path(p["feedback"][0]["path"]), [("T1", "甲", "=D2*2", 12, "进行中")],
         ("编号", "单位", "完成", "计划", "状态"))
    result = reconcile_export(p)
    assert not result["exported"] and not result["can_export"]
    assert any("公式" in i["message"] for i in result["issues"])


def test_explicit_null_mapping_disables_automatic_matching(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 11, 1, "进行中")])
    p["feedback"][0]["field_map"] = {"3": None}
    preview = reconcile_preview(p)
    assert not preview["can_export"]
    assert any(i["kind"] == "field_mapping" and i.get("target_field") == 3 for i in preview["issues"])


@pytest.mark.parametrize("axis", ["rows", "columns"])
def test_second_import_reverted_to_explicit_baseline_restores_only_font_color(tmp_path, axis):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")],
                [("T1", "甲", 12, 5, "进行中")], axis=axis)
    planned = "C2" if axis == "rows" else "B3"
    completed = "D2" if axis == "rows" else "B4"
    baseline_path = p["path"]
    wb = load_workbook(baseline_path)
    wb.active[planned].font = Font(name="宋体", size=13, bold=True, italic=True, color="FF112233")
    wb.active[completed].font = Font(color="FF775577")
    wb.save(baseline_path)
    wb.close()
    p["compare_fields"] = [3]
    p["baseline"] = {"path": baseline_path, "selection": selection(1, axis=axis)}
    first = reconcile_export(p)
    assert first["exported"], first
    wb = load_workbook(first["out"])
    assert wb.active[planned].font.color.rgb == "FFFF0000"
    # A user changes other font attributes between monthly imports.
    font = copy(wb.active[planned].font)
    font.sz = 16
    font.strike = True
    wb.active[planned].font = font
    wb.save(first["out"])
    wb.close()

    returned = book(tmp_path / "returned.xlsx", [("T1", "甲", 10, 8, "进行中")], axis=axis)
    p["path"] = first["out"]
    p["feedback"] = [{"path": str(returned), "selection": selection(1, axis=axis)}]
    p["out"] = str(tmp_path / "second.xlsx")
    preview = reconcile_preview(p)
    assert preview["can_export"]
    assert not any(change["field"] == 3 for change in preview["changes"])
    second = reconcile_export(p)
    assert second["exported"], second
    wb = load_workbook(second["out"])
    cell = wb.active[planned]
    assert cell.value == 10 and cell.font.color.rgb == "FF112233"
    assert cell.font.name == "宋体" and cell.font.sz == 16
    assert cell.font.bold and cell.font.italic and cell.font.strike
    assert wb.active[completed].value == 8 and wb.active[completed].font.color.rgb == "FF775577"
    wb.close()


def test_explicit_baseline_original_red_remains_red_without_content_change(tmp_path):
    p = payload(tmp_path, [("T1", "甲", 10, 0, "进行中")], [("T1", "甲", 10, 0, "进行中")])
    wb = load_workbook(p["path"])
    wb.active["C2"].font = Font(name="宋体", color="FFFF0000", bold=True)
    wb.save(p["path"])
    wb.close()
    p["baseline"] = {"path": p["path"], "selection": selection(1)}
    preview = reconcile_preview(p)
    assert preview["can_export"] and preview["changes"] == []
    assert reconcile_export(p)["exported"]
    wb = load_workbook(p["out"])
    assert wb.active["C2"].font.color.rgb == "FFFF0000" and wb.active["C2"].font.bold
    wb.close()
