"""Optional two-output workflow, composed from the same standalone operations."""
import json
import os
from contextlib import contextmanager
from pathlib import Path
import shutil
import tempfile

from .merge import Cancelled
from .native_split import split_preview, split_export
from .reconcile import reconcile_preview, reconcile_export


@contextmanager
def _temporary_directory(prefix, parent=None):
    parent = Path(parent or tempfile.gettempdir()).resolve()
    directory = Path(tempfile.mkdtemp(prefix=prefix, dir=parent)).resolve()
    if directory.parent != parent or not directory.name.startswith(prefix):
        raise ValueError("临时目录超出了指定父目录")
    try:
        yield directory
    finally:
        resolved = directory.resolve()
        if resolved != directory or resolved.parent != parent or not resolved.name.startswith(prefix):
            raise ValueError("临时目录路径发生变化，拒绝递归清理")
        shutil.rmtree(resolved)


def _optional_filtered_preview(filtered):
    """An optional empty result does not prevent exporting the updated whole."""
    if filtered.get("groups"):
        return filtered
    result = dict(filtered)
    result["issues"] = [issue for issue in filtered.get("issues", [])
                        if issue.get("code") != "no_records"]
    result["can_export"] = not any(issue.get("blocking", True) for issue in result["issues"])
    result["warnings"] = list(filtered.get("warnings", [])) + [
        "没有符合条件的记录；本次只导出更新后的完整表。"]
    return result


def _filtered_payload(payload, updated, report):
    selected = dict(payload["selection"])
    selected["data_end"] += report.get("totals", {}).get("additions", 0)
    options = dict(payload["filtered_output"])
    # File identity belongs to this transaction, never to a saved filtering scheme.
    for key in ("path", "sheet", "selection", "group_fields", "out_dir"):
        options.pop(key, None)
    return {**options, "path": str(updated), "sheet": payload["sheet"],
            "selection": selected, "group_fields": []}


def _merge_report(main, filtered):
    main["filtered_preview"] = filtered
    issues = []
    for issue in filtered.get("issues", []):
        entry = dict(issue)
        entry["id"] = "filtered:" + entry["id"]
        entry["message"] = "筛选结果：" + entry["message"]
        choices = []
        for choice in entry.get("choices", []):
            item = dict(choice)
            if item.get("patch"):
                item["patch"] = {"filtered_output": item["patch"]}
            choices.append(item)
        entry["choices"] = choices
        issues.append(entry)
    main["issues"] = main.get("issues", []) + issues
    main["warnings"] = main.get("warnings", []) + ["筛选结果：" + w for w in filtered.get("warnings", [])]
    main["can_export"] = bool(main["can_export"] and filtered["can_export"])
    return main


def workflow_preview(payload, cancel=None, progress=None):
    report = reconcile_preview(payload)
    if not payload.get("filtered_output") or not report["can_export"]:
        return report
    with _temporary_directory(prefix="excel-feedback-preview-") as temp:
        updated = Path(temp) / "updated.xlsx"
        # 筛选结果从"直接替换"形态计算(对比原值行对筛选无意义且会扰乱结构识别)
        materialized = reconcile_export({**payload, "out": str(updated), "result_form": "direct"}, cancel, progress)
        if not materialized.get("exported"):
            return materialized
        filtered = _optional_filtered_preview(split_preview(_filtered_payload(payload, updated, report)))
        return _merge_report(report, filtered)


def workflow_export(payload, cancel=None, progress=None):
    if not payload.get("filtered_output"):
        return reconcile_export(payload, cancel, progress)
    out = Path(payload["out"]).resolve()
    if out.suffix.lower() != ".xlsx":
        raise ValueError("输出文件必须使用 .xlsx 扩展名")
    inputs = [payload["path"]] + [f["path"] for f in payload.get("feedback", [])]
    if payload.get("baseline"):
        inputs.append(payload["baseline"]["path"])
    if out in {Path(p).resolve() for p in inputs}:
        raise ValueError("输出不能覆盖输入文件")
    destinations = [out, out.with_suffix(".changes.json"), out.with_name(out.stem + "_筛选结果.xlsx")]
    if any(p.exists() for p in destinations):
        raise ValueError("结果文件或变更清单已存在，请选择新名称")
    if not out.parent.is_dir():
        raise ValueError("输出目录不存在")
    created = []
    with _temporary_directory(prefix=".excel-feedback-", parent=out.parent) as temp:
        staged = Path(temp) / "updated.xlsx"
        report = reconcile_export({**payload, "out": str(staged)}, cancel, progress)
        if not report.get("exported"):
            return report
        split_source = staged
        if (payload.get("result_form") or "compare") != "direct":
            # 主输出保留了原值行; 筛选结果从干净的直接替换形态计算
            staged_direct = Path(temp) / "updated-direct.xlsx"
            direct = reconcile_export({**payload, "out": str(staged_direct), "result_form": "direct"},
                                      cancel, progress)
            if not direct.get("exported"):
                return direct
            split_source = staged_direct
        split_payload = _filtered_payload(payload, split_source, report)
        split_report = _optional_filtered_preview(split_preview(split_payload))
        result = _merge_report(report, split_report)
        if not result["can_export"]:
            result["exported"] = False
            return result
        artifacts = [(staged, destinations[0])]
        if split_report.get("groups"):
            split_dir = Path(temp) / "filtered"
            split_dir.mkdir()
            exported = split_export({**split_payload, "out_dir": str(split_dir)}, cancel, progress)
            files = [p for p in split_dir.rglob("*.xlsx")]
            if len(files) != 1:
                raise ValueError("筛选输出数量异常，未发布任何结果")
            artifacts.append((files[0], destinations[2]))
            result["filtered_output_path"] = str(destinations[2])
        result.update(out=str(out), output=str(out), changes_path=str(destinations[1]), exported=True)
        # Remove temporary paths from the delivered audit report.
        result.pop("filtered_preview", None)
        audit = Path(temp) / "audit.json"
        audit.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        artifacts.append((audit, destinations[1]))
        try:
            for source, destination in artifacts:
                if cancel and cancel():
                    raise Cancelled()
                os.link(source, destination)
                created.append(destination)
            if cancel and cancel():
                raise Cancelled()
        except BaseException:
            for p in created:
                p.unlink(missing_ok=True)
            raise
        result["outputs"] = [str(p) for p in created]
        return result
