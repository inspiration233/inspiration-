"""Report XLSX extension compatibility without blocking readable workbooks."""
from xml.etree import ElementTree as ET


_OFFICE_SPREADSHEET = "http://schemas.microsoft.com/office/spreadsheetml/"
# 已知无害的扩展: 丢失不影响数据与外观(计算引擎缓存/文件路径记录),
# WPS/Excel 保存的文件几乎都带, 提示它们等于"正常文件永远有警告"
_HARMLESS_FEATURES = {"calcFeatures", "absPath"}
_FORMATTING_FEATURES = {
    "conditionalFormattings": "扩展条件格式",
    "conditionalFormatting": "扩展条件格式",
    "dataValidations": "扩展下拉或数据验证",
    "sparklineGroups": "迷你图",
    "dxfs": "扩展差异样式",
    "tableStyles": "扩展表样式",
}


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _namespace(tag):
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


def inspect_extensions(name, raw):
    """Return notices for known and unknown extension payloads.

    An unknown extension is not evidence of corrupted data. The user chooses
    best-effort export by default, so report potential omission without
    requiring a feature allowlist. XML parse errors remain real errors.
    Report element names only: attributes/text may contain private paths.
    """
    root = ET.fromstring(raw)
    containers = [node for node in root.iter() if _local(node.tag) == "extLst"]
    populated = [node for node in containers if len(node)]
    if not populated:
        return []
    if name.lower() == "xl/styles.xml":
        return ["检测到扩展样式：已继续处理常规字体、边框和数字格式；"
                "附加样式信息可能未保留，请核对导出后的外观。"]
    notices = []
    unknown = []
    for container in populated:
        for extension in container:
            if _local(extension.tag) != "ext":
                unknown.append(_local(extension.tag))
                continue
            for feature in extension:
                feature_name = _local(feature.tag)
                if feature_name in _HARMLESS_FEATURES:
                    continue
                if (not name.lower().startswith("xl/worksheets/")
                        or not _namespace(feature.tag).startswith(_OFFICE_SPREADSHEET)
                        or feature_name not in _FORMATTING_FEATURES):
                    unknown.append(feature_name)
                    continue
                title = _FORMATTING_FEATURES[feature_name]
                message = f"{title}的附加部分可能未保留；单元格内容继续处理，请核对对应显示或输入规则。"
                if message not in notices:
                    notices.append(message)
    if unknown:
        features = list(dict.fromkeys(unknown))
        detail = "（" + "、".join(features[:6]) + ("等" if len(features) > 6 else "") + "）"
        notices.append(f"{name} 含扩展信息{detail}：已允许继续导出；"
                       "扩展附加信息可能未保留，请核对相关功能和结果。")
    return notices
