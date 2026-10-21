"""Desktop adapters: consistent preview/export with visible format compromises."""
from dataclasses import fields
from pathlib import Path
import os
import tempfile
import warnings as python_warnings
from zipfile import ZipFile
from xml.etree import ElementTree

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.writer.theme import theme_xml

from .merge import Options, Cancelled, merge, scan
from .preview import json_value
from .compatibility import inspect_extensions
from .package_check import validate_xlsx


def _theme_signature(blob):
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    root = ElementTree.fromstring(blob)
    # A newly created concatenation workbook has openpyxl's default theme.
    # Compare actual theme values, ignoring the descriptive scheme names.
    return tuple(
        (node.tag, tuple(sorted(node.attrib.items())), (node.text or "").strip())
        for section in ("clrScheme", "fontScheme")
        for branch in root.findall(f".//a:{section}", ns)
        for node in branch.iter() if node is not branch
    )


def options_for(payload):
    raw = dict(payload.get("options") or {})
    defaults = {"style": "fidelity", "remove_summary": False, "strip_serial_rows": False,
                "renumber_serial": False, "source_column": False, "preserve_blank_rows": True,
                "preserve_template_preamble": True, "metadata_columns": False, "exact_headers": False}
    defaults.update(raw)
    defaults["selections"] = payload.get("selections") or []
    # An explicit positional correspondence is stronger than the legacy
    # similarity heuristic, which could otherwise split renamed fields apart.
    if any(item.get("field_map") for item in defaults["selections"]):
        defaults["auto_group"] = False
    if payload.get("template"):
        defaults["template_path"] = payload["template"]["path"]
        defaults["template_sheet"] = payload["template"]["sheet"]
    if defaults.get("mode") == "auto":
        names = {}
        if defaults["selections"]:
            for item in defaults["selections"]:
                names.setdefault(item["path"], []).append(item["sheet"])
            count = len(defaults["selections"])
        else:
            report = scan(payload.get("paths", []), Options(include_hidden=defaults.get("include_hidden", False)))
            count = len(report["header_detections"])
            for item in report["files"]:
                names[item["path"]] = [s["sheet"] for s in item["sheets"] if not s.get("skipped")]
        values = list(names.values())
        defaults["mode"] = ("sheets" if len(values) == 1 and count > 1 else
                            "group" if len(values) > 1 and len(values[0]) > 1 and
                            all(set(v) == set(values[0]) for v in values) else "single")
    allowed = {f.name for f in fields(Options)}
    return Options(**{k: v for k, v in defaults.items() if k in allowed})


def _preflight(payload, options):
    issues, warnings = [], []
    paths = payload.get("paths") or []
    if not paths:
        return [{"id": "no-files", "message": "请先添加文件"}], warnings
    if options.style == "fidelity" and options.mode != "workbook":
        for raw in paths:
            path = Path(raw)
            if path.suffix.lower() != ".xlsx":
                warnings.append(f"{path.name}：原格式无法完整携带 Excel 原生属性；内容可合并，需要保留版式请先另存为 .xlsx。")
                continue
            with ZipFile(path) as archive:
                names = archive.namelist()
                if ("xl/theme/theme1.xml" in names and
                        _theme_signature(archive.read("xl/theme/theme1.xml")) != _theme_signature(theme_xml)):
                    warnings.append(f"{path.name} 使用自定义颜色或字体主题，拼接后部分颜色或字体可能变化；已继续处理。如需更接近各页外观，可选“装入同一文件”。")
                visual_prefixes = ("xl/drawings/", "xl/charts/", "xl/printerSettings/")
                if any(n.startswith(visual_prefixes) for n in names):
                    warnings.append(f"{path.name}：拼接结果可能不保留图表、图形或打印机专属设置；已继续处理单元格内容和常规格式。")
                unsupported = [n for n in names if n.startswith(("xl/pivotTables/", "xl/pivotCache/",
                               "xl/slicer", "xl/activeX/", "xl/externalLinks/", "xl/embeddings/",
                               "xl/queryTables/", "xl/richData/", "xl/model/"))
                               or n in ("xl/connections.xml", "xl/metadata.xml") or n.endswith("vbaProject.bin")]
                if unsupported:
                    issues.append({"id": f"objects:{path}", "message": f"{path.name} 包含数据透视、宏或外部数据依赖，处理后可能改变数据含义，暂不能可靠拼接。"})
                for name in names:
                    if not (name.startswith("xl/") and name.endswith(".xml")) or name.startswith(visual_prefixes):
                        continue
                    try:
                        warnings.extend(f"{path.name}：{notice}" for notice in inspect_extensions(name, archive.read(name)))
                    except ValueError as error:
                        issues.append({"id": f"extensions:{path}:{name}", "message": f"{path.name}：{error}"})
    return issues, warnings


def _generate(payload, out, cancel=None, progress=None):
    # Readers can report a discarded format feature as a Python warning. Keep
    # that information in both the preview and the export completion report.
    with python_warnings.catch_warnings(record=True) as caught:
        python_warnings.simplefilter("always", UserWarning)
        report = _generate_impl(payload, out, cancel, progress)
    report["warnings"] = list(dict.fromkeys([
        *report.get("warnings", []),
        *(f"读取兼容提示：{item.message}" for item in caught),
    ]))
    return report


def _generate_impl(payload, out, cancel=None, progress=None):
    options = options_for(payload)
    issues, warnings = _preflight(payload, options)
    if issues:
        return {"issues": issues, "warnings": warnings, "groups": [], "can_export": False, "totals": {}}
    if cancel and cancel():
        raise Cancelled()
    try:
        if options.mode == "workbook" and all(Path(p).suffix.lower() == ".xlsx" for p in payload.get("paths", [])):
            from .native_book import merge_workbook
            report = merge_workbook(payload, out, cancel=cancel, progress=progress)
        else:
            report = merge(payload["paths"], out, options, cancel=cancel, progress=progress)
    except ValueError as error:
        return {"issues": [{"id": "merge-selection", "message": str(error)}],
                "warnings": warnings, "groups": [], "can_export": False, "totals": {}}
    if report.get("cancelled") or (cancel and cancel()):
        raise Cancelled()
    issues.extend({"id": item.get("id", f"native:{i}"), **item}
                  for i, item in enumerate(report.get("issues", [])))
    warnings += report.get("warnings", [])
    if options.mode != "workbook":
        for item in report.get("header_detections", []):
            if item.get("orientation") == "horizontal":
                warnings.append(f"{item['display']}：横向记录将按所选字段拼接为纵向结果；需保持横向原布局请使用各表独立成页。")
    totals = report.get("totals", {})
    for key, title in (("formulas_to_value", "公式无法保留"),
                       ("data_merges_dropped", "数据合并单元格无法保留")):
        if totals.get(key):
            issues.append({"id": key, "message": f"{title}：{totals[key]} 项。可能影响计算或数据完整性，请调整字段对应/选区或使用各表独立成页。"})
    for key, title in (("validations_dropped", "下拉或数据验证"), ("cf_dropped", "条件格式"),
                       ("tables_downgraded", "智能表格功能")):
        if totals.get(key):
            warnings.append(f"{title}有 {totals[key]} 项未能完整保留；已继续处理单元格内容，请核对对应显示或输入规则。")
    for item in report.get("files", []):
        if item.get("status") == "failed":
            issues.append({"id": "file:" + item["path"], "message": f"{item['path']}：{item.get('reason', '读取失败')}"})
    if not Path(out).is_file() and not issues:
        issues.append({"id": "empty", "message": "没有可合并的数据"})
    groups = []
    valid_package = Path(out).is_file()
    if valid_package:
        try:
            validate_xlsx(out)
        except ValueError as error:
            valid_package = False
            issues.append({"id": "output-package", "message": f"生成文件的结构检查未通过：{error}。请反馈此提示，当前结果未发布。"})
    if valid_package:
        wb = load_workbook(out, data_only=False, read_only=True)
        try:
            originals = report.get("groups") or []
            for ws in wb:
                if ws.sheet_state != "visible":
                    continue
                group = next((g for g in originals if g.get("sheet", g.get("label")) == ws.title), {})
                groups.append({"label": ws.title, "sheet": ws.title, "shape": group.get("shape"),
                               "count": group.get("count"), "row_count": ws.max_row,
                               "sample_kind": "grid", "sample_start_row": 1,
                               "columns": [get_column_letter(i) for i in range(1, min(ws.max_column, 30) + 1)],
                               "sources": group.get("sources", []),
                               "sample": [[json_value(c.value) for c in row] for row in ws.iter_rows(
                                   min_row=1, max_row=min(ws.max_row, 20), max_col=min(ws.max_column, 30))]})
        finally:
            wb.close()
    report.update(groups=groups, issues=issues, warnings=list(dict.fromkeys(warnings)),
                  can_export=not issues, totals=totals)
    report.pop("output", None)
    report.pop("out", None)
    return report


def merge_preview(payload, cancel=None, progress=None):
    with tempfile.TemporaryDirectory(prefix="excel-merge-preview-") as temp:
        return _generate(payload, str(Path(temp) / "preview.xlsx"), cancel, progress)


def merge_export(payload, cancel=None, progress=None):
    out = Path(payload["out"]).resolve()
    inputs = {Path(p).resolve() for p in payload.get("paths", [])}
    if out in inputs:
        raise ValueError("输出不能覆盖输入文件")
    if out.suffix.lower() != ".xlsx":
        raise ValueError("合并结果请保存为 .xlsx")
    if out.exists():
        raise ValueError("输出文件已存在，请选择新文件名")
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".excel-merge-", dir=out.parent) as temp:
        staged = Path(temp) / out.name
        report = _generate(payload, str(staged), cancel, progress)
        if not report["can_export"]:
            error = ValueError("合并仍有未解决的问题：" + "；".join(i["message"] for i in report["issues"][:3]))
            error.preview = report
            raise error
        if cancel and cancel():
            raise Cancelled()
        # Exclusive reservation prevents overwriting a file created after preview.
        with out.open("xb"):
            pass
        try:
            os.replace(staged, out)
        except BaseException:
            out.unlink(missing_ok=True)
            raise
        report["output"] = str(out)
        return report
