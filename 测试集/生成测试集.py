# -*- coding: utf-8 -*-
"""
生成「Excel 合并小工具」的测试集。

用法:
    python 生成测试集.py [--rows N]    # N 为大规模文件的行数, 默认 100000

设计原则: 数据刻意做得"脏", 贴近真实办公场景——
  - 列顺序不同 / 列名带首尾空格 / 多余列 / 缺失列
  - 空行、空 Sheet、只有表头没有数据的文件、尾部"合计"行
  - 真日期与文本日期混杂、数字存成文本、日期序列号
  - GBK / UTF-8-BOM 编码的 CSV、Excel 97-2003 老格式 .xls
  - 双行表头、标题行占据第一行、合并单元格、损坏文件
  - 一个大目录, 用于性能与内存测试

注意: 不要手工"清洗"这些文件, 它们的脏点是测试点本身。
重新生成本测试集: 删除 测试集/场景*、测试集/边缘情况、测试集/大规模 后重跑本脚本。
"""
import argparse
import csv
import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent
random.seed(2026)  # 固定种子, 保证每次生成的数据一致

# ---------------------------------------------------------------- 数据池
NAMES = ["张伟", "王芳", "李娜", "刘强", "陈静", "杨洋", "赵敏", "刘洋", "周杰", "吴倩", "徐磊", "孙丽", "马超", "朱婷"]
PRODUCTS = [
    ("A4打印纸(箱)", 89.5), ("签字笔(盒)", 25.0), ("订书机", 18.8), ("文件夹(包)", 32.0),
    ("白板笔(盒)", 30.5), ("激光打印机", 1299.0), ("碎纸机", 599.0), ("会议记录本", 12.0),
    ("硒鼓", 388.0), ("白板(块)", 149.0),
]
CUSTOMERS = [
    "杭州明启科技有限公司", "上海岚图贸易有限公司", "深圳市恒信电子有限公司",
    "北京中科联合实业有限公司", "广州云图文化传媒有限公司", "成都锦绣家园物业管理有限公司",
    "武汉江城建材经营部", "南京新视野广告工作室",
]
HEADER_A = ["日期", "客户姓名", "产品", "数量", "单价", "销售额", "备注"]

THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def rand_date(start=(2026, 9, 20), end=(2026, 10, 12)):
    d = dt.date(2026, 10, 12) - dt.date(*start)
    return dt.date(*start) + dt.timedelta(days=random.randint(0, d.days))


def sale_row(name_pool=None):
    """造一条真实感销售记录; 刻意让部分数量/备注留空、部分文本带尾空格。"""
    product, price = random.choice(PRODUCTS)
    qty = random.choice([1, 2, 2, 3, 5, 8, 10, 12])
    remark = random.choice([None, None, None, "加急", "月结", "老客户折扣", "含运费 "])
    name = random.choice(name_pool or NAMES)
    return [rand_date(), name, product, qty, price, round(qty * price, 2), remark]


def style_header(ws, row=1):
    for c in ws[row]:
        c.font = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN


def save(wb, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"  ✓ {path.relative_to(ROOT)}")


def write_csv(path: Path, header, rows, encoding):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding=encoding) as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"  ✓ {path.relative_to(ROOT)}  (编码: {encoding})")


# ================================================================ 场景A
def scenario_a():
    print("场景A: 多文件纵向合并成一张总表")
    d = ROOT / "场景A_多文件纵向合并"

    # 部门A: 标准模板文件, 带基础样式(用于验证"保留原样式"), 中间夹一个空行, 尾部有"合计"行
    wb = Workbook(); ws = wb.active; ws.title = "销售明细"
    ws.append(HEADER_A)
    for _ in range(14):
        ws.append(sale_row())
    ws.append([None] * 7)                                   # 中间空行
    for _ in range(3):
        ws.append(sale_row())
    ws.append(["合计", None, None, "=SUM(...)", None, 8765.43, None])  # 尾部合计行
    style_header(ws)
    ws["G3"].fill = PatternFill("solid", fgColor="FFC7CE")   # 一个标红单元格
    ws["G3"].font = Font(color="9C0006")                     # (超期未回款标红这类操作)
    for col, w in zip("ABCDEFG", [11, 9, 14, 6, 8, 10, 10]):
        ws.column_dimensions[col].width = w
    save(wb, d / "部门A_销售明细.xlsx")

    # 部门B: 列顺序完全不同 + 列名带首尾空格 —— 宽容模式必须按列名(去空格)对齐
    wb = Workbook(); ws = wb.active; ws.title = "销售明细"
    ws.append(["数量", " 单价", "销售额", "产品 ", "客户姓名 ", "日期", "备注"])  # 刻意乱序+空格
    for _ in range(12):
        date, name, product, qty, price, amount, remark = sale_row()
        ws.append([qty, price, amount, product, name, date, remark])
    style_header(ws)
    save(wb, d / "部门B_销售明细.xlsx")

    # 部门C: 多一列"折扣"、缺"备注"列, 数据 8 行 —— 宽容模式下多出的列照收、缺的列留空
    wb = Workbook(); ws = wb.active; ws.title = "销售明细"
    ws.append(["日期", "客户姓名", "产品", "数量", "单价", "销售额", "折扣"])
    for _ in range(8):
        date, name, product, qty, price, amount, _ = sale_row()
        ws.append([date, name, product, qty, price, amount, random.choice([0.95, 1.0, 0.9, None])])
    style_header(ws)
    save(wb, d / "部门C_销售明细.xlsx")

    # 部门D: 改了两个列名("客户姓名"→"客户名称", "销售额"→"销售金额")
    #        —— 严格模式应报错并指明此文件; 宽容模式下会多出两条新列(语义识别待议的靶子)
    wb = Workbook(); ws = wb.active; ws.title = "销售明细"
    ws.append(["日期", "客户名称", "产品", "数量", "单价", "销售金额", "备注"])
    for _ in range(10):
        date, name, product, qty, price, amount, remark = sale_row()
        ws.append([date, name, product, qty, price, amount, remark])
    style_header(ws)
    save(wb, d / "部门D_销售明细.xlsx")

    # 部门E: 只有表头没有任何数据 —— 应跳过并在报告中说明, 而不是报错中断
    wb = Workbook(); ws = wb.active; ws.title = "销售明细"
    ws.append(HEADER_A)
    style_header(ws)
    save(wb, d / "部门E_销售明细.xlsx")

    # 华东区: GBK 编码 CSV, 其中一行与部门A首行完全相同(将来做去重功能的验证点)
    rows = [[r[0].strftime("%Y-%m-%d")] + r[1:] for r in (sale_row() for _ in range(10))]
    dup = ["2026-09-28", "张伟", "激光打印机", 1, 1299.0, 1299.0, "华东区补录"]
    rows.insert(3, dup)
    write_csv(d / "华东区_销售明细.csv", HEADER_A, rows, encoding="gbk")


# ================================================================ 场景B
def scenario_b():
    print("场景B: 多文件合成一个工作簿")
    d = ROOT / "场景B_多文件合成工作簿"

    def month_file(path, sheets):
        wb = Workbook()
        wb.remove(wb.active)
        for sname, n in sheets:
            ws = wb.create_sheet(sname)
            ws.append(["日期", "商品", "数量", "金额"])
            for _ in range(n):
                product, price = random.choice(PRODUCTS)
                ws.append([rand_date(), product, random.choice([1, 2, 3, 5, 10]), round(price * random.choice([1, 2, 3]), 2)])
            style_header(ws)
        save(wb, path)

    month_file(d / "1月.xlsx", [("销售", 10), ("库存", 8)])
    month_file(d / "2月.xlsx", [("销售", 10), ("库存", 8)])      # 与1月Sheet重名 → 结果中应为 销售(2)/库存(2)
    # 3月: Excel 97-2003 老格式 .xls; "销售"再次重名 → 销售(3)
    import xlwt
    wb = xlwt.Workbook(encoding="utf-8")
    for sname, n in [("销售", 9), ("退货", 5)]:
        ws = wb.add_sheet(sname)
        for j, h in enumerate(["日期", "商品", "数量", "金额"]):
            ws.write(0, j, h)
        for i in range(n):
            product, price = random.choice(PRODUCTS)
            row = [rand_date().strftime("%Y/%m/%d"), product, random.choice([1, 2, 4, 6]), round(price * 2, 2)]
            for j, v in enumerate(row):
                ws.write(i + 1, j, v)
    path = d / "3月.xls"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"  ✓ {path.relative_to(ROOT)}  (老格式 xls)")


# ================================================================ 场景C
def scenario_c():
    print("场景C: 单文件内多Sheet汇总成一张表")
    d = ROOT / "场景C_单文件多Sheet汇总"
    wb = Workbook()
    wb.remove(wb.active)
    header = ["日期", "单号", "商品", "数量", "金额", "经手人"]
    for day in ["周一", "周二", "周三", "周四"]:
        ws = wb.create_sheet(day)
        ws.append(header)
        for i in range(8):
            product, price = random.choice(PRODUCTS)
            ws.append([rand_date(), f"NO{random.randint(1000, 9999)}", product,
                       random.choice([1, 2, 3, 5]), round(price * random.choice([1, 2]), 2),
                       random.choice(NAMES)])
        if day == "周四":
            ws.append(["合计", None, None, 27, 4321.0, None])   # 尾部合计行: 真实店铺日结常见
        style_header(ws)
    wb.create_sheet("空白页")                                    # 全空Sheet: 应跳过并报告
    ws = wb.create_sheet("作废单")
    ws.append(["说明", "作废单据登记, 本页结构与其它页不同"])
    ws.append(["NO1024", "客户自取消"])
    save(wb, d / "门店日结.xlsx")


# ================================================================ 场景D
def scenario_d():
    print("场景D: 多文件同名Sheet分组汇总")
    d = ROOT / "场景D_同名Sheet分组汇总"
    header = ["日期", "商品", "数量", "金额"]

    def branch(path, sheets):
        wb = Workbook()
        wb.remove(wb.active)
        for sname, n in sheets:
            ws = wb.create_sheet(sname)
            ws.append(header)
            for _ in range(n):
                product, price = random.choice(PRODUCTS)
                ws.append([rand_date(), product, random.choice([1, 2, 3, 4, 6]), round(price * random.choice([1, 2, 3]), 2)])
            style_header(ws)
        save(wb, path)

    branch(d / "分部1.xlsx", [("销售", 9), ("库存", 7)])
    branch(d / "分部2.xlsx", [("销售", 11), ("库存", 6), ("退货", 5)])  # "退货"仅此分部有 → 独立成页
    branch(d / "分部3.xlsx", [("销售", 8), ("库存", 9)])
    # CSV没有Sheet概念 → 按约定视为单Sheet文件, Sheet名取文件名"广州分部", 独立成一组
    rows = [[rand_date().strftime("%Y-%m-%d"), random.choice(PRODUCTS)[0],
             random.choice([1, 2, 5]), round(random.uniform(50, 3000), 2)] for _ in range(9)]
    write_csv(d / "广州分部.csv", header, rows, encoding="utf-8-sig")


# ================================================================ 场景E
def scenario_e():
    print("场景E: 表头智能识别(表头行探测/多行表头/合计行剔除)")
    d = ROOT / "场景E_表头智能识别"
    header = ["日期", "客户", "产品", "数量", "单价", "金额"]

    # 甲: 标题行(合并A1:F1) + 制表信息行(合并A2:F2) + 真表头在第3行 + 尾部合计行
    #    数据中埋一个"合计装订册(套)"干扰项(在第3列, 不应被合计行剔除逻辑误杀)
    wb = Workbook(); ws = wb.active; ws.title = "销售"
    ws.append(["甲部门2026年10月销售报表"])
    ws.merge_cells("A1:F1")
    ws["A1"].font = Font(bold=True, size=14); ws["A1"].alignment = Alignment(horizontal="center")
    ws.append(["制表人：王芳    制表日期：2026-10-11    单位：元"])
    ws.merge_cells("A2:F2")
    ws.append(header)
    for _ in range(10):
        date, name, product, qty, price, amount, _ = sale_row()
        ws.append([date, name, product, qty, price, amount])
    ws.append([rand_date(), random.choice(CUSTOMERS), "合计装订册(套)", 2, 45.0, 90.0])  # 干扰项
    for _ in range(3):
        date, name, product, qty, price, amount, _ = sale_row()
        ws.append([date, name, product, qty, price, amount])
    ws.append(["合计", None, None, 48, None, 6123.45])
    style_header(ws, row=3)
    save(wb, d / "报表_甲.xlsx")

    # 乙: 表头在第1行(基准) + 尾部"本页小计"/"总计" + 尾部空行 —— 与甲字段名相同, 供跨文件共识探测
    wb = Workbook(); ws = wb.active; ws.title = "销售"
    ws.append(header)
    for _ in range(12):
        date, name, product, qty, price, amount, _ = sale_row()
        ws.append([date, name, product, qty, price, amount])
    ws.append([None] * 6)
    ws.append(["本页小计", None, None, 35, None, 4820.10])
    ws.append([None] * 6)
    ws.append(["总计", None, None, 35, None, 4820.10])
    style_header(ws)
    save(wb, d / "报表_乙.xlsx")

    # 丙: CSV带前导空行 + 按区域分组、组间夹"小计"(两种写法) + 尾部"合计"
    rows = [
        ["2026-10-01", "华东区", "杭州明启科技有限公司", "A4打印纸(箱)", 5, 447.5],
        ["2026-10-02", "华东区", "上海岚图贸易有限公司", "签字笔(盒)", 20, 500.0],
        ["2026-10-03", "华东区", "上海岚图贸易有限公司", "白板笔(盒)", 6, 183.0],
        ["华东区小计", "", "", "", 31, 1130.5],            # 关键词在第1列(带前缀)
        ["2026-10-05", "华北区", "北京中科联合实业有限公司", "激光打印机", 1, 1299.0],
        ["2026-10-06", "华北区", "北京中科联合实业有限公司", "碎纸机", 1, 599.0],
        ["", "小计", "", "", 2, 1898.0],                   # 关键词在第2列
        ["2026-10-08", "华南区", "深圳市恒信电子有限公司", "硒鼓", 3, 1164.0],
        ["华南区小计", "", "", "", 3, 1164.0],
        ["合计", "", "", "", 36, 4192.5],
    ]
    # 手工写文件: 前导空行必须在表头之前(某些系统导出的特征), write_csv 做不到
    path = d / "报表_丙.csv"
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([]); w.writerow([])                   # 前导空行
        w.writerow(["日期", "区域", "客户", "产品", "数量", "金额"])
        w.writerows(rows)
    print(f"  ✓ {path.relative_to(ROOT)}  (编码: utf-8, 带前导空行)")

    # 双行表头跨文件合并: 两个文件结构相同, 验证展平后按复合列名("商品信息/品名")对齐
    def double_header_file(path, n):
        wb = Workbook(); ws = wb.active; ws.title = "费用"
        ws.append(["订单号", "商品信息", "商品信息", "金额", "金额"])
        ws.append(["", "品名", "数量", "单价", "总额"])
        ws.merge_cells("B1:C1"); ws.merge_cells("D1:E1")
        for _ in range(n):
            product, price = random.choice(PRODUCTS)
            qty = random.choice([1, 2, 3])
            ws.append([f"DD{random.randint(10000, 99999)}", product, qty, price, round(qty * price, 2)])
        for r in (1, 2):
            for c in ws[r]:
                c.font = Font(bold=True)
                c.border = THIN
        save(wb, path)

    double_header_file(d / "双行表头_1月.xlsx", 9)
    double_header_file(d / "双行表头_2月.xlsx", 11)

    # 双行表头_3月: 平铺单行表头、字段与双行表头文件相同 —— 验证单行/多行表头互相对齐
    # 预期: 模板取第一个文件时, 输出仍保留两行表头结构, 3月的数据按内部对齐键(品名=商品信息/品名)填入
    wb = Workbook(); ws = wb.active; ws.title = "费用"
    ws.append(["订单号", "品名", "数量", "单价", "总额"])
    for _ in range(10):
        product, price = random.choice(PRODUCTS)
        qty = random.choice([1, 2, 3])
        ws.append([f"DD{random.randint(10000, 99999)}", product, qty, price, round(qty * price, 2)])
    style_header(ws)
    save(wb, d / "双行表头_3月.xlsx")

    # 双行表头变体: 在跨列合并基础上叠加"跨行合并"——订单号合并A1:A2、备注合并F1:F2
    # 预期: 订单号/备注为单列名(不加后缀), 与双行表头_1月/2月可按复合列名对齐(多一列备注)
    wb = Workbook(); ws = wb.active; ws.title = "费用"
    ws.append(["订单号", "商品信息", "商品信息", "金额", "金额", "备注"])
    ws.append(["", "品名", "数量", "单价", "总额", ""])
    ws.merge_cells("A1:A2"); ws.merge_cells("F1:F2")
    ws.merge_cells("B1:C1"); ws.merge_cells("D1:E1")
    for _ in range(8):
        product, price = random.choice(PRODUCTS)
        qty = random.choice([1, 2, 4])
        ws.append([f"DD{random.randint(10000, 99999)}", product, qty, price, round(qty * price, 2),
                   random.choice([None, "含安装", "加急"])])
    for r in (1, 2):
        for c in ws[r]:
            c.font = Font(bold=True)
            c.border = THIN
    save(wb, d / "双行表头_变体.xlsx")

    # 纵向合并越界: "订单号"合并A1:A2, 盖住了第一条数据行 —— 现实中拖拽合并范围失手的常见产物
    # 预期: 表头区域仍判定为第1行(内容信号优先于合并范围); 被覆盖的数据格按空值透传; 报告提示
    wb = Workbook(); ws = wb.active; ws.title = "费用"
    ws.append(["订单号", "品名", "数量", "单价", "金额"])
    ws.append([None, "白板笔(盒)", 4, 30.5, 122.0])     # 第一条数据的订单号被合并覆盖, 录不进去
    for _ in range(8):
        product, price = random.choice(PRODUCTS)
        qty = random.choice([1, 2, 3])
        ws.append([f"DD{random.randint(10000, 99999)}", product, qty, price, round(qty * price, 2)])
    ws.merge_cells("A1:A2")                              # 先写数据再合并, 模拟越界
    style_header(ws)
    save(wb, d / "纵向合并越界.xlsx")

    # 单行表头带横向合并、无子级行: "商品"横跨A、B两列, 但数据区两列各有内容(品名/规格)
    # 预期: 展平为 商品(1)、商品(2) 两列, 数据一个不丢
    wb = Workbook(); ws = wb.active; ws.title = "报价"
    ws.append(["商品", "商品", "数量", "单价", "金额"])
    ws.merge_cells("A1:B1")
    for _ in range(8):
        product, price = random.choice(PRODUCTS)
        qty = random.choice([1, 2, 3, 5])
        ws.append([product, f"X{random.randint(100, 999)}", qty, price, round(qty * price, 2)])
    style_header(ws)
    save(wb, d / "单行表头带合并.xlsx")


# ================================================================ 边缘情况
def edge_cases():
    print("边缘情况")
    d = ROOT / "边缘情况"

    # 1) 标题行占据第一行, 真正表头在第2行(合并了A1:F1) —— 现实中极常见
    wb = Workbook(); ws = wb.active; ws.title = "回款"
    ws.append(["2026年第三季度客户回款统计表"])
    ws.merge_cells("A1:F1")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.append(["到账日期", "客户", "合同号", "应收金额", "实收金额", "状态"])
    for _ in range(10):
        recv = round(random.uniform(5000, 200000), 2)
        ws.append([rand_date((2026, 7, 1), ), random.choice(CUSTOMERS),
                   f"HT-{random.randint(2026, 9999)}", recv, round(recv * random.uniform(0.8, 1.0), 2),
                   random.choice(["已到账", "部分到账", "逾期"])])
    style_header(ws, row=2)
    save(wb, d / "表头不在第一行.xlsx")

    # 2) 双行表头 + 合并单元格 —— 对任何合并工具都是硬骨头, V1预期行为见README
    wb = Workbook(); ws = wb.active; ws.title = "费用"
    ws.append(["订单号", "商品信息", "商品信息", "金额", "金额"])
    ws.append(["", "品名", "数量", "单价", "总价"])
    ws.merge_cells("B1:C1"); ws.merge_cells("D1:E1")
    for _ in range(9):
        product, price = random.choice(PRODUCTS)
        qty = random.choice([1, 2, 3])
        ws.append([f"DD{random.randint(100, 999)}", product, qty, price, round(qty * price, 2)])
    for row in (1, 2):
        for c in ws[row]:
            c.font = Font(bold=True)
    save(wb, d / "双行表头.xlsx")

    # 3) 样式丰富: 验证"保留原样式"模式(填充色/字体/数字格式/列宽/冻结首行/边框)
    wb = Workbook(); ws = wb.active; ws.title = "回款明细"
    ws.append(["到账日期", "客户", "合同号", "应收", "实收", "差额", "状态"])
    for _ in range(10):
        recv = round(random.uniform(8000, 160000), 2)
        paid = round(recv * random.uniform(0.7, 1.0), 2)
        ws.append([rand_date((2026, 8, 1), ), random.choice(CUSTOMERS),
                   f"HT-{random.randint(1000, 9999)}", recv, paid, round(recv - paid, 2),
                   "已结清" if paid >= recv else "未结清"])
    style_header(ws)
    for r in range(2, 12):
        ws.cell(r, 4).number_format = "#,##0.00"
        ws.cell(r, 5).number_format = "#,##0.00"
        if ws.cell(r, 7).value == "未结清":
            ws.cell(r, 7).fill = PatternFill("solid", fgColor="FFF2CC")
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFG", [12, 26, 10, 12, 12, 10, 9]):
        ws.column_dimensions[col].width = w
    save(wb, d / "样式丰富.xlsx")

    # 4) 日期格式混乱: 真日期/文本日期/中文日期/小数点日期/Excel日期序列号混杂
    wb = Workbook(); ws = wb.active; ws.title = "流水"
    ws.append(["日期", "事项", "金额"])
    mixed = [
        [dt.date(2026, 10, 1), "银行转账", 12000.0],
        ["2026/10/2", "现金收款", 800.0],
        ["2026-10-3", "支付宝收款", 2340.5],
        ["10月4日", "银行转账", 5600.0],
        ["2026.10.5", "微信收款", 990.0],
        [46552, "系统导入(日期序列号)", 300.0],       # 46552 = 2027-07-10 的序列号, 坏导出的产物
        [None, "未注明日期", 150.0],
    ]
    for row in mixed:
        ws.append(row)
    style_header(ws)
    save(wb, d / "日期格式混乱.xlsx")

    # 5) 损坏文件: 伪装成xlsx的文本 —— 工具必须优雅报错, 不能崩溃
    path = d / "损坏文件.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("这不是一个真正的Excel文件, 用于测试错误处理。\n" * 3, encoding="utf-8")
    print(f"  ✓ {path.relative_to(ROOT)}  (故意损坏)")

    # 6) 重名列/无名列/单元格内换行/长数字 —— 特殊列规则(V1.3)验证
    wb = Workbook(); ws = wb.active; ws.title = "明细"
    ws.append(["单号", "客户", "金额", "金额", None, "身份证号", "地址"])
    #      两个"金额"重名列 | E列无表头但有数据 | 身份证(文本,安全) | 地址含Alt+Enter换行
    ws.append(["DD1001", "张伟", 1000.0, 200.0, "手工补录", "330106199001011234",
               "上海市浦东新区\n世纪大道100号"])
    ws.append([20261015000120016, "王芳", 800.0, 300.0, None, "110101198505052345",
               "北京市海淀区知春路1号"])   # 单号是16位数字(源头已损, 透传验证)
    for i in range(6):
        ws.append([f"DD{1002 + i}", random.choice(NAMES), round(random.uniform(100, 3000), 2),
                   round(random.uniform(100, 3000), 2), random.choice([None, "手工补录"]),
                   f"33010619{random.randint(70, 99)}0101{random.randint(1000, 9999)}",
                   f"杭州市西湖区文{random.randint(1, 99)}号"])
    style_header(ws)
    save(wb, d / "重名列与无名列.xlsx")

    # 8) 横向表(转置): 字段在首列A, 记录按列排列(B/C/D) —— 方向自动识别靶子
    wb = Workbook(); ws = wb.active; ws.title = "横向"
    ws.append(["客户", "上海岚图贸易有限公司", "北京中科联合实业有限公司", "杭州明启科技有限公司"])
    ws.append(["联系人", "王芳", "李娜", "张伟"])
    ws.append(["签订日期", dt.date(2026, 8, 12), dt.date(2026, 9, 3), dt.date(2026, 10, 1)])
    ws.append(["金额", 128000.0, 96500.5, 233000.0])
    ws.append(["回款状态", "已结清", "部分回款", "未回款"])
    ws.append(["备注", None, "分两期", None])
    for c in "ABCD":
        ws.column_dimensions[c].width = 22
    save(wb, d / "横向表.xlsx")

    # 9) 序号空行: 第3/5行只有序号没有内容 —— 剔除后序号重排 1..N 的靶子
    wb = Workbook(); ws = wb.active; ws.title = "台账"
    ws.append(["序号", "事项", "金额", "经手人"])
    data = [("复印纸采购", 350.0, "王芳"), None, ("打印机维修", 1200.0, "李娜"),
            None, None, ("团建费用", 8000.0, "张伟"), ("快递费", 156.5, "王芳")]
    for i, item in enumerate(data, start=1):
        if item is None:
            ws.append([i, None, None, None])          # 仅序号的空行
        else:
            ws.append([i, item[0], item[1], item[2]])
    ws.append(["合计", None, 9706.5, None])           # 尾部合计(同时验证剔除)
    style_header(ws)
    save(wb, d / "序号空行.xlsx")

    # 7) 隐藏Sheet: 可见"销售"页 + 隐藏"配置"页(模板下拉选项, 真实模板常有) —— 默认跳过隐藏页并报告
    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet("销售")
    ws.append(["日期", "商品", "数量", "金额"])
    for _ in range(8):
        product, price = random.choice(PRODUCTS)
        ws.append([rand_date(), product, random.choice([1, 2, 3]), round(price * random.choice([1, 2]), 2)])
    style_header(ws)
    hid = wb.create_sheet("配置")
    hid["A1"] = "下拉选项配置页, 请勿删除"
    hid["A2"] = "待审批"
    hid["A3"] = "已审批"
    hid.sheet_state = "hidden"
    save(wb, d / "隐藏Sheet.xlsx")


# ================================================================ 大规模
def large_scale(rows: int):
    print(f"大规模性能测试集 (每个文件 {rows} 行)")
    d = ROOT / "大规模"
    header = ["序号", "日期", "客户", "业务员", "产品", "数量", "单价", "金额", "状态"]

    def gen_row(i):
        product, price = random.choice(PRODUCTS)
        return [i, (dt.date(2026, 1, 1) + dt.timedelta(days=i % 300)).strftime("%Y-%m-%d"),
                random.choice(CUSTOMERS), random.choice(NAMES), product,
                i % 17 + 1, price, round((i % 17 + 1) * price, 2),
                random.choice(["已发货", "已签收", "退货中", "已回款"])]

    for n in (1, 2):
        wb = Workbook(write_only=True)   # 流式写出, 生成大数据文件本身也不吃内存
        ws = wb.create_sheet(f"数据{n}")
        ws.append(header)
        for i in range(rows):
            ws.append(gen_row(i))
        save(wb, d / f"海量数据_{n}.xlsx")

    write_csv(d / "海量数据_3.csv", header,
              [gen_row(i) for i in range(rows)], encoding="utf-8-sig")


# ---------------------------------------------------------------- main
def main():
    parser = argparse.ArgumentParser(description="生成Excel合并工具测试集")
    parser.add_argument("--rows", type=int, default=100000, help="大规模文件行数 (默认 100000)")
    parser.add_argument("--skip-large", action="store_true", help="跳过大规模文件生成")
    args = parser.parse_args()

    print(f"测试集根目录: {ROOT}\n")
    scenario_a()
    scenario_b()
    scenario_c()
    scenario_d()
    scenario_e()
    edge_cases()
    if not args.skip_large:
        large_scale(args.rows)
    print("\n全部生成完毕。各文件的测试要点见 README.md。")


if __name__ == "__main__":
    main()
