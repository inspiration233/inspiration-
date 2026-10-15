# -*- coding: utf-8 -*-
"""识别层单元测试：规范化 / 汇总行 / 探测 / 展平键。"""
from engine.tools.excel_merge import recognize as rc
from engine.tools.excel_merge import reader as rd

from conftest import scenario


# --------------------------------------------------------------- 规范化

def test_normalize_name():
    assert rc.normalize_name(" 姓名 ") == "姓名"
    assert rc.normalize_name("商品　信息") == "商品 信息"      # 全角空格折叠
    assert rc.normalize_name("金额（元）") == "金额(元)"        # 全角括号转半角
    assert rc.normalize_name("ID") == "id"
    assert rc.normalize_name("ID") == rc.normalize_name(" id ")
    assert rc.normalize_name(None) == ""


# --------------------------------------------------------------- 汇总行剔除

def test_summary_row_variants():
    kw = rc.DEFAULT_SUMMARY_KEYWORDS
    assert rc.is_summary_row(["合计", None, 1])
    assert rc.is_summary_row(["总计", None, 1], kw)
    assert rc.is_summary_row(["华东区小计", "", "", 31, 1130.5])       # 第1列带前缀
    assert rc.is_summary_row(["", "小计", "", 2, 1898.0])              # 第2列
    assert rc.is_summary_row(["本页小计", None])
    assert not rc.is_summary_row(["2026-10-01", "合计装订册(套)", 2])   # 干扰项在第3列
    assert not rc.is_summary_row(["DD1024", "客户自取消"])
    assert not rc.is_summary_row([None, None, None])


# --------------------------------------------------------------- 表头探测（基于测试集）

def _structure(path, sheet=None):
    tables, reader = rd.list_tables(path)
    try:
        srcs = [t for t, _ in tables if t is not None]
        if sheet:
            srcs = [s for s in srcs if s.structure.sheet_name == sheet]
        return srcs[0].structure
    finally:
        reader.close()


def test_detect_header_row3_with_title():
    st = _structure(scenario("场景E_表头智能识别") / "报表_甲.xlsx")
    assert st.header_row == 3
    assert st.height == 1
    assert st.title == "甲部门2026年10月销售报表"
    assert st.keys == ["日期", "客户", "产品", "数量", "单价", "金额"]


def test_detect_csv_leading_blank_lines():
    st = _structure(scenario("场景E_表头智能识别") / "报表_丙.csv")
    assert st.header_row == 3                     # 前导空行×2
    assert "区域" in st.keys


def test_detect_header_not_in_first_row():
    st = _structure(scenario("边缘情况") / "表头不在第一行.xlsx")
    assert st.header_row == 2
    assert st.title == "2026年第三季度客户回款统计表"


def test_double_header_height_and_keys():
    st = _structure(scenario("场景E_表头智能识别") / "双行表头_1月.xlsx")
    assert st.height == 2
    assert st.keys == ["订单号", "商品信息/品名", "商品信息/数量", "金额/单价", "金额/总额"]


def test_double_header_variant_vertical_merge():
    st = _structure(scenario("场景E_表头智能识别") / "双行表头_变体.xlsx")
    assert st.keys[:2] == ["订单号", "商品信息/品名"]
    assert "备注" in st.keys                       # 跨行合并 → 单列名不加后缀


def test_flat_header_with_horizontal_merge_no_child():
    st = _structure(scenario("场景E_表头智能识别") / "单行表头带合并.xlsx")
    assert st.height == 1                          # 行2是数据, 不扩展
    assert st.keys[:2] == ["商品(1)", "商品(2)"]   # 末行横向合并无子级 → 同名加序号


def test_out_of_bounds_vertical_merge():
    st = _structure(scenario("场景E_表头智能识别") / "纵向合并越界.xlsx")
    assert st.height == 1                          # 内容信号优先, 不被越界合并拉高
    assert any("延伸进数据区" in w for w in st.warnings)


def test_flat_and_multirow_align_same_keys():
    from engine.tools.excel_merge.align import align_sources
    d = scenario("场景E_表头智能识别")
    t1, r1 = rd.list_tables(d / "双行表头_1月.xlsx")
    t3, r3 = rd.list_tables(d / "双行表头_3月.xlsx")
    try:
        s1 = [t for t, _ in t1 if t][0]
        s3 = [t for t, _ in t3 if t][0]
        out_keys, maps = align_sources([s1, s3], "loose")
        # 平铺的 品名/数量/单价/总额 通过末段回退对上复合键
        assert out_keys == ["订单号", "商品信息/品名", "商品信息/数量",
                            "金额/单价", "金额/总额"]
        assert maps[1] == [0, 1, 2, 3, 4]      # 3月五列分别对上五列
    finally:
        r1.close()
        r3.close()


def test_duplicate_and_unnamed_columns():
    st = _structure(scenario("边缘情况") / "重名列与无名列.xlsx")
    assert st.keys.count("金额") == 0
    assert "金额(1)" in st.keys and "金额(2)" in st.keys
    assert "未命名1" in st.keys


# --------------------------------------------------------------- 编码

def test_csv_encoding_detection(tmp_path):
    p1 = tmp_path / "a.csv"
    p1.write_text("客户,金额\n张伟,100", encoding="gbk")
    p2 = tmp_path / "b.csv"
    p2.write_text("客户,金额\n王芳,200", encoding="utf-8-sig")
    assert rd.detect_csv_encoding(p1) in ("gbk", "gb18030")
    assert rd.detect_csv_encoding(p2) == "utf-8-sig"
