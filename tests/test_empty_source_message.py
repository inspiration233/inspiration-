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
