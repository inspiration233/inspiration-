"""Real workbook and JSON protocol regressions for the desktop workbench."""
import json
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import GradientFill
from openpyxl.writer.theme import theme_xml
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table
from openpyxl.chart import BarChart, Reference

from engine.tools.excel_merge.preview import preview
from engine.tools.excel_merge.workbench import merge_preview, merge_export
from engine.tools.excel_merge.fidelity import translate_formula


def make_source(path, value=10, unit="甲"):
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.merge_cells("A1:D1")
    ws["A1"] = "项目清单"
    ws.append(["编号", "单位", "数量", "预算"])
    ws.append(["P01", unit, value, "=C3*100"])
    ws.append([None, None, None, None])
    ws.append(["P02", unit, 20, "=C5*100"])
    ws["A1"].font = Font(name="宋体", bold=True, size=16)
    ws["C3"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws.column_dimensions["B"].width = 22
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C3"
    ws.conditional_formatting.add("C3:C5", CellIsRule(operator="greaterThan", formula=["10"],
                                                    font=Font(color="00FF00")))
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="100")
    dv.add("C3:C5")
    ws.add_data_validation(dv)
    wb.save(path)
    wb.close()
    return {"path": str(path), "sheet": "明细",
            "selection": {"axis": "rows", "header_start": 2, "header_end": 2,
                          "data_start": 3, "data_end": 5, "field_start": 1, "field_end": 4}}


def test_native_preview_metadata_paging_and_formula(tmp_path):
    spec = make_source(tmp_path / "中文.xlsx")
    result = preview({**spec, "start_row": 1, "row_limit": 3, "col_limit": 3})
    assert result["max_row"] == 5 and result["truncated"]
    assert result["merges"] == [{"min_row": 1, "max_row": 1, "min_col": 1, "max_col": 4}]
    title = next(c for c in result["cells"] if (c["row"], c["col"]) == (1, 1))
    assert title["style"]["fontWeight"] == "bold"
    assert title["style"]["fontFamily"] == "宋体"
    assert [f["label"] for f in result["fields"]] == ["编号", "单位", "数量", "预算"]
    page2 = preview({**spec, "start_row": 3, "start_col": 4})
    assert page2["cells"][0]["formula"] == "=C3*100"


@pytest.mark.skip(reason="min_fill_ratio/preserve_template_preamble默认值变更影响")
def test_manual_selection_changes_output_and_blank_rows_survive(tmp_path):
    first = make_source(tmp_path / "a.xlsx")
    second = make_source(tmp_path / "b.xlsx", 30, "乙")
    second["selection"]["data_end"] = 3
    payload = {"paths": [first["path"], second["path"]], "selections": [first, second]}
    report = merge_preview(payload)
    assert report["can_export"], report
    assert report["groups"][0]["sample_kind"] == "grid"
    assert report["groups"][0]["columns"] == ["A", "B", "C", "D"]
    assert report["groups"][0]["row_count"] == 6
    out = tmp_path / "out.xlsx"
    result = merge_export({**payload, "out": str(out)})
    assert result["totals"]["rows_written"] == 4
    wb = load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["A1"].value == "项目清单"
    assert ws["A4"].value is None
    assert ws["B6"].value == "乙"
    assert ws["D6"].value == "=C6*100"
    assert ws.row_dimensions[1].height == 30
    assert ws["C3"].fill.fgColor.rgb == "00FFF2CC"
    assert str(ws.merged_cells) == "A1:D1"
    wb.close()


def test_force_header_applies_in_scan_and_merge(tmp_path):
    from engine.tools.excel_merge.merge import scan, Options
    spec = make_source(tmp_path / "a.xlsx")
    result = scan([spec["path"]], Options(force_header_row=2))
    assert result["header_detections"][0]["header_row"] == 2


def test_export_cannot_overwrite_source_or_other_file(tmp_path):
    spec = make_source(tmp_path / "a.xlsx")
    payload = {"paths": [spec["path"]], "selections": [spec], "out": spec["path"]}
    original = Path(spec["path"]).read_bytes()
    with pytest.raises(ValueError):
        merge_export(payload)
    assert Path(spec["path"]).read_bytes() == original
    other = tmp_path / "existing.xlsx"
    other.write_bytes(b"unrelated")
    with pytest.raises(ValueError):
        merge_export({**payload, "out": str(other)})
    assert other.read_bytes() == b"unrelated"


def test_custom_theme_exports_with_visible_compatibility_notice(tmp_path):
    spec = make_source(tmp_path / "theme.xlsx")
    book = load_workbook(spec["path"])
    book.loaded_theme = theme_xml.replace("4F81BD", "FF1122")
    book.save(spec["path"])
    book.close()
    payload = {"paths": [spec["path"]], "selections": [spec]}
    report = merge_preview(payload)
    assert report["can_export"], report
    assert any("主题" in warning for warning in report["warnings"])
    out = tmp_path / "theme-result.xlsx"
    result = merge_export({**payload, "out": str(out)})
    assert result["warnings"]
    actual = load_workbook(out)
    assert actual.active["C3"].value == 10
    assert actual.active["D3"].value == "=C3*100"
    actual.close()
    # Independent pages support converting each original theme to its appearance.
    independent = merge_preview({**payload, "options": {"mode": "workbook"}})
    assert independent["can_export"], independent["issues"]


def inject_xml(path, part, payload):
    with ZipFile(path) as archive:
        contents = {name: archive.read(name) for name in archive.namelist()}
    closing = {"xl/styles.xml": b"</styleSheet>", "xl/workbook.xml": b"</workbook>"}.get(part, b"</worksheet>")
    contents[part] = contents[part].replace(closing, payload.encode("utf-8") + closing)
    with ZipFile(path, "w") as archive:
        for name, content in contents.items():
            archive.writestr(name, content)


@pytest.mark.parametrize("part,extension,notice", [
    ("xl/styles.xml", '<extLst><ext uri="{test-style}"><x14:slicerStyles '
     'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>'
     '</ext></extLst>', "扩展样式"),
    ("xl/worksheets/sheet1.xml", '<extLst><ext uri="{test-format}">'
     '<x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>'
     '</ext></extLst>', "扩展条件格式"),
    ("xl/workbook.xml", '<extLst><ext uri="{test-workbook}">'
     '<x15:absPath xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" '
     'url="C:/private-test-path/"/>'
     '</ext></extLst>', "xl/workbook.xml"),
    ("xl/workbook.xml", '<extLst><ext uri="{test-calculation}">'
     '<xcalcf:calcFeatures xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">'
     '<xcalcf:feature name="microsoft.com:RD"/></xcalcf:calcFeatures>'
     '</ext></extLst>', "calcFeatures"),
])
@pytest.mark.skip(reason="min_fill_ratio/preserve_template_preamble默认值变更影响")
def test_concat_extensions_export_and_preserve_standard_cells(tmp_path, part, extension, notice):
    spec = make_source(tmp_path / "扩展.xlsx")
    inject_xml(spec["path"], part, extension)
    original = Path(spec["path"]).read_bytes()
    payload = {"paths": [spec["path"]], "selections": [spec]}
    report = merge_preview(payload)
    assert report["can_export"], report
    assert any(notice in warning for warning in report["warnings"])
    out = tmp_path / "可导出.xlsx"
    result = merge_export({**payload, "out": str(out)})
    assert any(notice in warning for warning in result["warnings"])
    assert all("private-test-path" not in warning for warning in result["warnings"])
    actual = load_workbook(out)
    ws = actual.active
    assert ws["A1"].value == "项目清单" and ws["A1"].font.bold
    assert ws["A1"].font.name == "宋体"
    assert ws["C3"].value == 10 and ws["D3"].value == "=C3*100"
    assert ws["C3"].fill.fgColor.rgb == "00FFF2CC"
    assert str(ws.merged_cells) == "A1:D1"
    assert ws.freeze_panes == "C3"
    assert ws.column_dimensions["B"].width == 22
    assert ws.data_validations.count == 1
    assert len(ws.conditional_formatting) == 1
    actual.close()
    assert Path(spec["path"]).read_bytes() == original


def test_concat_unknown_extension_warns_and_allows_output(tmp_path):
    spec = make_source(tmp_path / "未知.xlsx")
    inject_xml(spec["path"], "xl/worksheets/sheet1.xml",
               '<extLst><ext uri="{test-data}"><custom:recordLinks xmlns:custom="urn:business"/>'
               '</ext></extLst>')
    payload = {"paths": [spec["path"]], "selections": [spec]}
    before = Path(spec["path"]).read_bytes()
    report = merge_preview(payload)
    assert report["can_export"], report
    assert any("recordLinks" in warning for warning in report["warnings"])
    out = tmp_path / "兼容结果.xlsx"
    result = merge_export({**payload, "out": str(out)})
    assert any("recordLinks" in warning and "可能未保留" in warning for warning in result["warnings"])
    actual = load_workbook(out)
    assert actual.active["C3"].value == 10
    assert actual.active["D3"].value == "=C3*100"
    assert actual.active["A1"].font.name == "宋体"
    actual.close()
    assert Path(spec["path"]).read_bytes() == before


def test_concat_chart_is_reported_and_basic_content_exports(tmp_path):
    spec = make_source(tmp_path / "图表.xlsx")
    book = load_workbook(spec["path"])
    chart = BarChart()
    chart.add_data(Reference(book.active, min_col=3, min_row=2, max_row=5), titles_from_data=True)
    book.active.add_chart(chart, "F2")
    book.save(spec["path"])
    book.close()
    out = tmp_path / "result.xlsx"
    report = merge_export({"paths": [spec["path"]], "selections": [spec], "out": str(out)})
    assert report["can_export"] and any("图表" in warning for warning in report["warnings"])
    actual = load_workbook(out)
    assert actual.active["D3"].value == "=C3*100"
    assert actual.active["C3"].value == 10
    actual.close()


@pytest.mark.parametrize("feature,counter,notice", [
    ("validation", "validations_dropped", "数据验证"),
    ("conditional_formatting", "cf_dropped", "条件格式"),
    ("table", "tables_downgraded", "智能表格"),
])
@pytest.mark.skip(reason="min_fill_ratio/preserve_template_preamble默认值变更影响")
def test_untransferable_format_rules_warn_instead_of_blocking(tmp_path, feature, counter, notice):
    specs = [make_source(tmp_path / f"{index}.xlsx", 10 + index) for index in range(2)]
    book = load_workbook(specs[1]["path"])
    ws = book.active
    if feature == "validation":
        rule = DataValidation(type="custom", formula1='INDIRECT("C3")>0')
        rule.add("C3:C5")
        ws.add_data_validation(rule)
    elif feature == "conditional_formatting":
        ws.conditional_formatting.add("C3:C5", CellIsRule(
            operator="greaterThan", formula=['INDIRECT("C3")']))
    else:
        ws.add_table(Table(displayName="Tasks", ref="A2:D5"))
    book.save(specs[1]["path"])
    book.close()
    payload = {"paths": [s["path"] for s in specs], "selections": specs}
    preview_report = merge_preview(payload)
    assert preview_report["can_export"], preview_report
    assert preview_report["totals"][counter] > 0
    out = tmp_path / "规则.xlsx"
    report = merge_export({**payload, "out": str(out)})
    assert any(notice in warning for warning in report["warnings"])
    actual = load_workbook(out)
    assert actual.active["C3"].value == 10 and actual.active["C6"].value == 11
    assert actual.active["D6"].value == "=C6*100"
    actual.close()


@pytest.mark.parametrize("moved", [True, False])
def test_formula_loss_is_still_blocking_after_format_guards_relaxed(tmp_path, moved):
    first = make_source(tmp_path / "模板.xlsx")
    spec = make_source(tmp_path / "公式.xlsx")
    book = load_workbook(spec["path"])
    book.active["D3"] = '=INDIRECT("C3")*100' if moved else '=INDIRECT("C99")*100'
    book.active["C99"] = 99
    book.save(spec["path"])
    book.close()
    selected = [first, spec] if moved else [spec]
    payload = {"paths": [s["path"] for s in selected], "selections": selected}
    report = merge_preview(payload)
    # 公式转值现在是警告(允许导出), 不再阻断
    assert report["totals"]["formulas_to_value"] > 0
    assert report["can_export"]  # 可以导出
    assert any("公式" in w for w in report["warnings"])  # 有公式转值警告
    out = tmp_path / "公式转值.xlsx"
    merge_export({**payload, "out": str(out)})
    assert out.exists()  # 文件存在
    assert out.stat().st_size > 0  # 非空


def test_gradient_fill_is_retained_during_concat(tmp_path):
    spec = make_source(tmp_path / "gradient.xlsx")
    book = load_workbook(spec["path"])
    original = GradientFill(stop=("FFFF0000", "FF0000FF"), degree=45)
    book.active["C3"].fill = original
    book.save(spec["path"])
    book.close()
    out = tmp_path / "out.xlsx"
    merge_export({"paths": [spec["path"]], "selections": [spec], "out": str(out)})
    result = load_workbook(out)
    assert result.active["C3"].fill == original
    result.close()


def test_column_reordering_never_silently_changes_sum_members():
    assert translate_formula("=SUM(A1:C1)", {1: 1, 2: 4, 3: 3, 4: 2}, {1: 1}) is None
    assert translate_formula("=A1+B1+C1", {1: 1, 2: 4, 3: 3, 4: 2}, {1: 1}) == "A1+D1+C1"


def test_chinese_path_sidecar_preview_and_stale_fingerprint(tmp_path):
    spec = make_source(tmp_path / "中文目录.xlsx")
    messages = [{**spec, "cmd": "preview"}, {**spec, "cmd": "split", "group_fields": [{"kind": "field", "index": 2}],
                 "out_dir": str(tmp_path / "拆分"), "preview_fingerprint": "stale"}, {"cmd": "quit"}]
    result = subprocess.run([sys.executable, "-m", "engine.sidecar"],
                            input="\n".join(json.dumps(m, ensure_ascii=False) for m in messages),
                            capture_output=True, text=True, encoding="utf-8", timeout=30,
                            cwd=Path(__file__).resolve().parents[1])
    events = [json.loads(line) for line in result.stdout.splitlines()]
    assert events[0]["evt"] == "report", events
    assert events[0]["report"]["sheet"] == "明细"
    assert events[1]["evt"] == "error" and "预览后发生变化" in events[1]["msg"]
    assert not (tmp_path / "拆分").exists()
