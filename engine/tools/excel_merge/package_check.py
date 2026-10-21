"""Check generated XLSX references that permissive readers can overlook.

This checks package integrity, not a feature allowlist or full Excel conformance.
Unknown extensions remain allowed. No workbook cell contents enter diagnostics.
"""
import posixpath
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET
from zipfile import ZipFile, BadZipFile


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_RID = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def _index(value, count, description):
    try:
        valid = 0 <= int(value) < count
    except (TypeError, ValueError):
        valid = False
    if not valid:
        raise ValueError(f"{description} 引用了不存在的编号")


def _owner(name):
    if name == "_rels/.rels":
        return ""
    return posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])


def _target(owner, target):
    part = unquote(urlsplit(target).path)
    return posixpath.normpath(part.lstrip("/") if part.startswith("/")
                             else posixpath.join(posixpath.dirname(owner), part))


def _styles(archive):
    root = ET.fromstring(archive.read("xl/styles.xml"))
    collections = {node.tag.rsplit("}", 1)[-1]: list(node) for node in root}
    number_formats = {int(fmt.get("numFmtId")) for fmt in collections.get("numFmts", [])}
    for section in ("cellStyleXfs", "cellXfs"):
        for xf in collections.get(section, []):
            for attr, collection in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                if attr in xf.attrib:
                    _index(xf.get(attr), len(collections.get(collection, [])), f"xl/styles.xml 的 {attr}")
            if section == "cellXfs" and "xfId" in xf.attrib:
                _index(xf.get("xfId"), len(collections.get("cellStyleXfs", [])), "xl/styles.xml 的 xfId")
            number_format = int(xf.get("numFmtId", 0))
            if number_format >= 164 and number_format not in number_formats:
                raise ValueError("xl/styles.xml 引用了不存在的自定义数字格式")
    for named in collections.get("cellStyles", []):
        _index(named.get("xfId"), len(collections.get("cellStyleXfs", [])), "xl/styles.xml 的命名样式")
    return len(collections.get("cellXfs", [])), len(collections.get("dxfs", []))


def validate_xlsx(path):
    """Raise ValueError for broken ZIP/XML/relationships/style references only."""
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            members = set(names)
            if len(names) != len(members):
                raise ValueError("工作簿压缩包含重复文件条目")
            if archive.testzip() is not None:
                raise ValueError("工作簿压缩包校验失败")
            required = {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/styles.xml"}
            if not required <= members:
                raise ValueError("工作簿缺少必要的文件结构")
            manifest = ET.fromstring(archive.read("[Content_Types].xml"))
            for entry in manifest.findall(f"{{{_CT}}}Override"):
                if unquote(entry.get("PartName", "")).lstrip("/") not in members:
                    raise ValueError("工作簿内容清单引用了缺失的文件")
            relations = {}
            for name in names:
                if not name.endswith(".rels"):
                    continue
                owner = _owner(name)
                if owner and owner not in members:
                    raise ValueError(f"{name} 缺少所属文件")
                ids = set()
                for rel in ET.fromstring(archive.read(name)).findall(f"{{{_RELS}}}Relationship"):
                    rid = rel.get("Id")
                    if not rid or rid in ids:
                        raise ValueError(f"{name} 的关系编号缺失或重复")
                    ids.add(rid)
                    if rel.get("TargetMode") != "External":
                        if _target(owner, rel.get("Target", "")) not in members:
                            raise ValueError(f"{name} 引用了缺失的内部文件")
                relations[owner] = ids
            cell_styles, differential_styles = _styles(archive)
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            book_views = len(workbook.findall(f"{{{_MAIN}}}bookViews/{{{_MAIN}}}workbookView"))
            for name in names:
                if not name.endswith(".xml") or name == "xl/styles.xml":
                    continue
                # Stream worksheet XML: validation should not keep another full
                # worksheet tree alongside the engine's native workbook objects.
                with archive.open(name) as stream:
                    for event, node in ET.iterparse(stream, events=("start", "end")):
                        if event == "end":
                            node.clear()
                            continue
                        for attr, value in node.attrib.items():
                            if attr.startswith("{" + _RID + "}") and value not in relations.get(name, set()):
                                raise ValueError(f"{name} 含没有对应关系的文件引用")
                            if attr.lower().endswith("dxfid") and node.tag.startswith("{" + _MAIN + "}"):
                                _index(value, differential_styles, f"{name} 的条件/表格样式")
                        if node.tag in (f"{{{_MAIN}}}c", f"{{{_MAIN}}}row") and "s" in node.attrib:
                            _index(node.get("s"), cell_styles, f"{name} 的单元格/行样式")
                        if node.tag == f"{{{_MAIN}}}col" and "style" in node.attrib:
                            _index(node.get("style"), cell_styles, f"{name} 的列样式")
                        if node.tag == f"{{{_MAIN}}}sheetView":
                            _index(node.get("workbookViewId", 0), book_views, f"{name} 的工作簿视图")
    except (BadZipFile, ET.ParseError, KeyError) as error:
        raise ValueError("生成的工作簿压缩包或 XML 结构不完整") from error
