"""整表为空(样式骨架无数据)时的报错必须指明"没有数据", 而不是误导模板设置。"""
import os
import tempfile

import pytest
from openpyxl import Workbook

from engine.tools.excel_merge.merge import Options, merge


def test_all_empty_sources_message_points_to_data_not_template(tmp_path):
    wb = Workbook(); ws = wb.active; ws.title = "填报表"
    ws.merge_cells("A1:B2")                       # 只有样式骨架, 无任何值
    ws.column_dimensions["A"].width = 20
    path = tmp_path / "空骨架.xlsx"; wb.save(path)
    out = tmp_path / "out.xlsx"
    with pytest.raises(ValueError) as info:
        merge([str(path)], str(out), Options(template_path=str(path),
                                             template_sheet="填报表"))
    message = str(info.value)
    assert "没有可合并的数据" in message, message
    assert "版式参考表" not in message


def test_shell_file_suppresses_compat_warnings(tmp_path):
    # 空壳文件的兼容警告(主题/图表/calcFeatures)全部抑制——内容不进输出,提示纯属误导
    import sys
    from openpyxl import Workbook
    sys.path.insert(0, '.')
    from engine.tools.excel_merge.workbench import merge_preview
    wb = Workbook(); ws = wb.active; ws.title = "填报表"
    ws.merge_cells("A1:B2"); ws.column_dimensions["A"].width = 20
    path = tmp_path / "空壳.xlsx"; wb.save(path)
    # 注入扩展样式与 calcFeatures 触发原警告路径
    from zipfile import ZipFile, ZIP_DEFLATED
    from io import BytesIO
    import xml.etree.ElementTree as ET
    buf = BytesIO(); wb.save(buf)
    with ZipFile(BytesIO(buf.getvalue())) as src, ZipFile(path, "w", ZIP_DEFLATED) as dst:
        for name in src.namelist():
            data = src.read(name)
            if name == "xl/workbook.xml":
                data = data.replace(b"</workbook>",
                    b'<extLst><ext uri="{C66500BD-6782-4325-8C5A-71D048A78866}"/></extLst></workbook>')
            dst.writestr(name, data)
    rep = merge_preview({"paths": [str(path)], "template": {"path": str(path), "sheet": "填报表"}})
    assert not rep["can_export"]
    messages = [str(i.get("message")) for i in rep["issues"]]
    assert any("没有可合并的数据" in m for m in messages)
    joined = " ".join(str(w) for w in rep["warnings"])
    assert "calcFeatures" not in joined and "扩展样式" not in joined and "主题" not in joined
