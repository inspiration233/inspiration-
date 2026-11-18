# -*- coding: utf-8 -*-
"""识别 v2(全局候选评分)的行为测试: 组合策略下 v1 为基线,
冻结/筛选(用户钦定)时 v2 接管; v1 低置信时 v2 附议提示。"""
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from engine.tools.excel_merge import reader, recognize2  # noqa: E402
from engine.tools.excel_merge.reader import _XlsxReader  # noqa: E402


def _book(tmp_path, rows, freeze=None, autofilter=None, name="t.xlsx"):
    wb = Workbook(); ws = wb.active; ws.title = "S"
    for row in rows:
        ws.append(list(row))
    if freeze:
        ws.freeze_panes = freeze
    if autofilter:
        ws.auto_filter.ref = autofilter
    path = tmp_path / name
    wb.save(path); wb.close()
    return path


def test_freeze_overrides_header_guess(tmp_path, monkeypatch):
    # 全文本脏表里 v1 判不准; 冻结窗格=用户钦定表头位置, v2 接管
    monkeypatch.setattr(reader, "RECOG_ENGINE", "v2")
    rows = [["单位", "项目", "状态"],
            ["甲公司", "道路施工", "进行中"], ["乙公司", "桥涵检测", "已完成"],
            ["丙公司", "绿化养护", "进行中"], ["合计", "3项", ""]]
    path = _book(tmp_path, rows, freeze="A2")     # 冻结首行 → 表头=第1行
    rd = _XlsxReader(str(path))
    st = rd.build_structure("S")
    # 冻结保护下结果正确; 接管警告仅在冻结判断与 v1 分歧时出现(v1 一致则无痕)
    assert st.header_row == 1 and st.height == 1


def test_low_confidence_second_opinion(tmp_path, monkeypatch):
    # v1 低置信且 v2 有不同意见: 不抢答, 提示"另一种判断"供人工核对
    monkeypatch.setattr(reader, "RECOG_ENGINE", "v2")
    rows = [["单位", "项目"], ["甲", "道路"], ["乙", "桥涵"], ["丙", "绿化"]]
    path = _book(tmp_path, rows)
    rd = _XlsxReader(str(path))
    st = rd.build_structure("S")
    # 无冻结/筛选且结构干净 → v1 结果原样, 无接管警告
    assert st.header_row == 1
    assert not any("冻结/筛选" in w for w in st.warnings)


def test_v2full_switch_runs(tmp_path, monkeypatch):
    # v2full 开关: 全量走 v2 全局评分(供后续打磨), 路径不炸且结果合法
    monkeypatch.setattr(reader, "RECOG_ENGINE", "v2full")
    rows = [["标题", None, None], ["序号", "名称", "金额"], [1, "a", 10], [2, "b", 20]]
    path = _book(tmp_path, rows)
    rd = _XlsxReader(str(path))
    st = rd.build_structure("S")
    assert st.header_row in (1, 2) and 1 <= st.height <= 2


def test_guess_scenarios_direct():
    # recognize2 单元: 标题排除/合并双行/宽标签/高度先验
    from engine.tools.excel_merge.recognize import MergeRange
    g = recognize2.detect_structure(
        [["全年计划", None, None], ["序号", "名称", "金额"], [1, "a", 10], [2, "b", 20], [3, "c", 30]])
    assert (g.header_row, g.height) == (2, 1)
    g2 = recognize2.detect_structure(
        [["基本信息", None, None], ["序号", "名称", "单位"], [1, "a", "甲"], [2, "b", "乙"], [3, "c", "甲"]],
        [MergeRange(1, 1, 3, 1)])
    # "基本信息"(A1:C1 整行大合并)按标题样排除 → 子行独立成表头, 语义不损
    assert (g2.header_row, g2.height) == (2, 1)
    g3 = recognize2.detect_structure(
        [["填报说明如下", None, None], ["一、按月填报", None, None], [None, None, None],
         ["序号", "名称", "金额"], [1, "a", 10], [2, "b", 20]])
    assert (g3.header_row, g3.height) == (4, 1)


def test_style_facts_quick_and_deep(tmp_path):
    rows = [["单位", "数量"], ["甲", 10], ["乙", 20]]
    path = _book(tmp_path, rows, freeze="A2", autofilter="A1:B3")
    rd = _XlsxReader(str(path))
    quick = rd._style_facts("S", 3, quick_only=True)
    assert quick.freeze_row == 2 and quick.filter_row == 1
    assert quick.bold_rows == ()               # quick 不做逐格扫描
