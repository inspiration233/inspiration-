# -*- coding: utf-8 -*-
"""模糊测试集生成器 —— 刻意"不懂引擎规则"地随机造表。

设计原则(与 生成测试集.py 相反): 参数全部由随机数决定, 不经过引擎设计者的思路。
用途: 供 不变量测试(tests/test_fuzz_invariants.py) 与人工抽查看怪案例。

    python 生成模糊集.py --seeds 0-49 --out 模糊集     # 生成一批到目录
"""
import argparse
import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

NAMES = ["张伟", "王芳", "李娜", "刘强", "陈静", "杨洋", "赵敏", "周杰",
         "吴倩", "徐磊", "孙丽", "马超", "朱婷", "胡军", "郭涛", "林红"]
ITEMS = ["打印纸", "签字笔", "订书机", "文件夹", "白板笔", "打印机", "碎纸机",
         "硒鼓", "白板", "会议本", "扫描仪", "装订机"]
UNITS = ["总部", "一分部", "二分部", "华东区", "华南区", "北区"]
STATES = ["已审批", "待审批", "已驳回", "进行中", "已完成"]


def _cell_value(rng, col_kind):
    r = rng.random()
    if col_kind == "num":
        if r < 0.1:
            return None
        return round(rng.uniform(1, 99999), rng.choice([0, 1, 2]))
    if col_kind == "date":
        if r < 0.1:
            return None
        return dt.date(2026, rng.randint(1, 12), rng.randint(1, 28))
    if r < 0.12:
        return None
    v = rng.choice(NAMES + ITEMS + UNITS + STATES)
    if r < 0.3:
        v = f" {v} "                      # 带空格
    if r < 0.18:
        v = str(rng.randint(1, 99999))    # 数字文本
    return v


def gen_case(seed: int, path: Path):
    """随机构造一张结构合法但参数随机的表.

    随机的是: 行列数/表头位置/双行表头/方向/列类型/脏数据(空行/序号空行/合计行/空格/
    数字文本/公式/合并/隐藏页), 不随机的是: 表必须是"前置区+表头+数据"的合法分区.
    """
    rng = random.Random(seed)
    wb = Workbook()
    ws = wb.active
    ws.title = rng.choice(["数据", "Sheet1", "明细", "报表", "a"])

    ncols = rng.randint(3, 7)
    header_row = rng.choice([1, 1, 1, 2, 2, 3, 4])
    double_header = rng.random() < 0.25 and ncols >= 4 and header_row == 1
    transposed = rng.random() < 0.18 and not double_header and ncols >= 5
    nrows = rng.randint(4, 20)

    col_kinds = [rng.choice(["num", "text", "text", "date"]) for _ in range(ncols)]
    headers = []
    for i in range(ncols):
        h = rng.choice(["序号", "姓名", "经手人", "品名", "数量", "金额",
                        "日期", "单位", "状态", "备注", "编号", f"列{i+1}"])
        if rng.random() < 0.15:
            h = f" {h}" if rng.random() < 0.5 else f"{h} "
        headers.append(h)
    if len(set(h.strip() for h in headers)) != len(headers):
        headers = [f"{h.strip()}{i+1}" for i, h in enumerate(headers)]

    # ---- 组装 grid: [前置区(空行)] + [表头行(1-2行)] + [数据行] ----
    grid = [[None] * ncols for _ in range(header_row - 1)]
    if double_header:
        grid.append([headers[c] if rng.random() < 0.6 else None for c in range(ncols)])
        grid.append([f"子{c+1}" if rng.random() < 0.7 else headers[c]
                     for c in range(ncols)])
    else:
        grid.append(list(headers))
    hdr_end = len(grid)                             # grid 中表头最后一行的下标(0-based)

    for _ in range(nrows):
        row = []
        for c in range(ncols):
            if col_kinds[c] == "num":
                row.append(round(rng.uniform(1, 99999), rng.choice([0, 1, 2])))
            elif col_kinds[c] == "date":
                row.append(dt.date(2026, rng.randint(1, 12), rng.randint(1, 28)))
            else:
                v = rng.choice(NAMES + ITEMS + UNITS + STATES)
                r = rng.random()
                if r < 0.15:
                    v = f" {v} "
                elif r < 0.25:
                    v = str(rng.randint(1, 99999))
                row.append(v if r >= 0.05 else None)
        p = rng.random()
        if p < 0.08:
            row = [None] * ncols                    # 全空行
        elif p < 0.14:
            row = [rng.randint(1, 999)] + [None] * (ncols - 1)  # 仅序号空行
        elif p < 0.2:
            row = [rng.choice(["合计", "小计", "总计"])] +                   [round(rng.uniform(1, 999), 1) if rng.random() < 0.5 else None
                   for _ in range(ncols - 1)]
        row = [v if not (isinstance(v, str) and v.strip() == "") else None for v in row]
        grid.append(row)

    # 公式注入(引用同行数字列)
    for i in range(hdr_end, len(grid)):
        if rng.random() > 0.15 or ncols < 2:
            continue
        c = rng.randrange(1, ncols)
        if isinstance(grid[i][c - 1], (int, float)):
            grid[i][c] = f"={chr(65 + c - 1)}{i + 1}*2"

    if transposed:
        # 造横向表: 字段沿列排(原grid的列变行), 记录沿行(原grid的行变列)
        grid = [[grid[r][c] if r < len(grid) and c < len(grid[r]) else None
                 for r in range(len(grid))] for c in range(ncols)]
        ncols = max(len(r) for r in grid)
        header_row = 1

    for row in grid:
        ws.append(row)
    sheet_hdr_row = header_row                      # 写入 sheet 后的表头行号(1-based)

    # ---- 前置区覆盖(仅非转置且真有前置区) ----
    if header_row > 1:
        ws.cell(row=1, column=1, value=f"随机标题{seed}")
        if ncols > 1:
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=ncols)
        if header_row > 2 and rng.random() < 0.6:
            ws.cell(row=2, column=1, value=f"说明：随机种子{seed}生成")
            if ncols > 1:
                ws.merge_cells(start_row=2, start_column=1,
                               end_row=2, end_column=ncols)

    # 双行表头的跨列合并
    if double_header:
        c0 = rng.randrange(0, max(1, ncols - 2))
        ws.merge_cells(start_row=sheet_hdr_row, start_column=c0 + 1,
                       end_row=sheet_hdr_row, end_column=min(c0 + 2, ncols))

    # 数据区纵向合并(随机)
    if not transposed and nrows > 6 and rng.random() < 0.35:
        c = rng.randrange(0, ncols)
        r0 = sheet_hdr_row + (2 if double_header else 1) + 1
        if r0 + 2 <= ws.max_row:
            ws.merge_cells(start_row=r0, start_column=c + 1,
                           end_row=min(r0 + 2, ws.max_row), end_column=c + 1)

    if rng.random() < 0.3:
        ws.freeze_panes = f"A{sheet_hdr_row + (2 if double_header else 1)}"
    if rng.random() < 0.25 and ncols >= 3:
        dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
        ws.add_data_validation(dv)
        last = rng.randrange(2, ncols + 1)
        dv.add(f"{chr(64 + last)}2:{chr(64 + last)}{max(3, ws.max_row)}")
    if rng.random() < 0.15:
        hid = wb.create_sheet("配置页")
        hid["A1"] = "x"
        hid.sheet_state = "hidden"
    if rng.random() < 0.2:
        extra = wb.create_sheet("备注页")
        extra["A1"] = f"seed={seed}"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--seeds", default="0-49", help="如 0-49 或 7")
    ap.add_argument("--out", default="模糊集")
    args = ap.parse_args()
    if "-" in args.seeds:
        lo, hi = map(int, args.seeds.split("-"))
        seeds = range(lo, hi + 1)
    else:
        seeds = [int(args.seeds)]
    out = Path(args.out)
    for s in seeds:
        p = out / f"模糊_{s:04d}.xlsx"
        gen_case(s, p)
        print(f"✓ {p}")


if __name__ == "__main__":
    main()
