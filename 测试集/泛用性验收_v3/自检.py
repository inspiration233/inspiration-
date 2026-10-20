# -*- coding: utf-8 -*-
"""泛用性验收测试集自检 v2 — 按 case.json 发现全部案例，必检项有失败判断"""
import json, sys
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
R = {"pass": 0, "fail": 0, "skip": 0, "warn": 0}
DETAIL = {"pass": [], "fail": [], "skip": [], "warn": []}

def _r(kind, msg):
    R[kind] += 1
    DETAIL[kind].append(msg)
    icon = {"pass": "✓", "fail": "✗", "skip": "○", "warn": "⚠"}[kind]
    print(f"  {icon} {msg}")

def discover_cases():
    """按 case.json 发现全部案例，不依赖目录名前缀"""
    cases = []
    for cj in sorted(ROOT.rglob("case.json")):
        d = cj.parent
        cases.append((d.name, d, json.loads(cj.read_text(encoding="utf-8"))))
    return cases

def check_file_readable(d, name):
    f = d / name
    if not f.exists():
        _r("fail", f"{d.name}/{name}: 文件不存在")
        return None
    try:
        wb = load_workbook(f)
        _r("pass", f"{d.name}/{name}: 可读({len(wb.sheetnames)}Sheet)")
        return wb
    except Exception as e:
        _r("fail", f"{d.name}/{name}: 无法打开 — {e}")
        return None

def check_data(d, cid, cj):
    """数据正确性检查（按 case.json 中的真值）"""
    gt = cj.get("ground_truth", {})
    if not gt:
        _r("skip", f"{cid}: 无 ground_truth，跳过数据检查")
        return

    # 找原始文件
    orig_files = [f for f in d.glob("*.xlsx") if "原始" in f.name or "original" in f.name.lower()]
    if not orig_files:
        orig_files = [f for f in d.glob("*.xlsx") if not any(
            k in f.name for k in ["反馈", "拆分", "预期", "汇总", "更新"])]
    if not orig_files:
        _r("skip", f"{cid}: 无原始文件")
        return
    wb = load_workbook(orig_files[0])
    ws = wb.active

    # 检查总任务数
    expected_tasks = gt.get("total_tasks")
    if expected_tasks:
        data_rows = sum(1 for r in range(1, ws.max_row + 1)
                       if ws.cell(r, 1).value and not ws.cell(r, 1).font.bold)
        # 减去前置区行
        hdr_row = cj.get("structure", {}).get("header_row", 1)
        if cid == 'W5':
            actual = sum(sum(1 for r in range(2, wb[sn].max_row + 1) if wb[sn].cell(r, 1).value)
                        for sn in wb.sheetnames)
        else:
            actual = sum(1 for r in range(hdr_row + 1, ws.max_row + 1) if ws.cell(r, 1).value)
        if actual == expected_tasks:
            _r("pass", f"{cid}: 任务数={actual}")
        else:
            _r("fail", f"{cid}: 任务数{actual}≠期望{expected_tasks}")

    # 检查单位拆分
    units = gt.get("units", [])
    splits = list(d.glob("*拆分*.xlsx"))
    if units and splits:
        split_units = set()
        for f in splits:
            wbs = load_workbook(f)
            for sn in wbs.sheetnames:
                split_units.add(sn)
        for u in units:
            if u in split_units:
                _r("pass", f"{cid}: 拆分含'{u}'")
            else:
                _r("fail", f"{cid}: 拆分缺少'{u}'")
        # 检查拆分文件内数据单位一致性
        for f in splits:
            wbs = load_workbook(f)
            wss = wbs.active
            unit_name = wss.title
            # W2横向月份表A列是分组单位, C列是月份数据, 不检查C列单位一致
            if cid == 'W2':
                continue
            rows_ok = all(wss.cell(r, 3).value in (unit_name, None)
                         for r in range(2, wss.max_row + 1) if wss.cell(r, 1).value)
            if rows_ok:
                _r("pass", f"{cid}/{f.stem}: 单位一致")
            else:
                _r("fail", f"{cid}/{f.stem}: 混入其他单位数据")

def check_fidelity(d, cid, cj):
    """保真检查（按 case.json 中的 checks）"""
    checks = cj.get("checks", {})
    if not checks:
        return
    orig = [f for f in d.glob("*.xlsx") if "原始" in f.name or (d / f"{cid.split('_')[0]}_原始.xlsx").exists]
    if not orig:
        orig = list(d.glob("*.xlsx"))
    if not orig:
        _r("skip", f"{cid}: 无文件可检查保真")
        return
    wb = load_workbook(orig[0])

    for check_name, expected in checks.items():
        _r("warn", f"{cid}/{check_name}: 需人工在Excel中验证 — {expected}")

def check_workflow(d, cid, cj):
    """完整流程检查"""
    needed = {"original": False, "split": False, "feedback": False,
              "monthly": False, "annual": False}
    for f in d.glob("*.xlsx"):
        n = f.name.lower()
        if "原始" in f.name or "original" in n: needed["original"] = True
        if "拆分" in f.name or "split" in n: needed["split"] = True
        if "反馈" in f.name or "feedback" in n: needed["feedback"] = True
        if "汇总" in f.name or "monthly" in n: needed["monthly"] = True
        if "更新" in f.name or "annual" in n or "update" in n: needed["annual"] = True

    if cj.get("type") == "完整流程":
        for k, v in needed.items():
            if v:
                _r("pass", f"{cid}: {k}文件存在")
            else:
                _r("fail", f"{cid}: 缺少{k}文件")

def check_ambiguity(d, cid, cj):
    """歧义案例检查"""
    amb = cj.get("ambiguity", [])
    if amb:
        for a in amb:
            _r("warn", f"{cid}: 歧义待确认 — {a[:50]}...")
        # 歧义案例不算 pass 也不算 fail
        if not cj.get("expected_changes") and not cj.get("expected"):
            _r("skip", f"{cid}: 歧义案例，不计入通过")

def check_w1_details(d):
    """W1 具体内容检查"""
    cid = "W1"
    # 原始表
    wb = check_file_readable(d, "W1_原始计划表.xlsx")
    if not wb: return
    ws = wb.active
    if ws.cell(4, 1).value == "任务编号":
        _r("pass", f"{cid}: 表头在第4行")
    else:
        _r("fail", f"{cid}: 表头不在第4行({ws.cell(4,1).value})")
    if ws.cell(5, 6).value == "=E5*100":
        _r("pass", f"{cid}: 预算公式=E5*100")
    else:
        _r("fail", f"{cid}: 预算公式错误({ws.cell(5,6).value})")
    dvs = ws.data_validations.dataValidation
    if any('"进行中,已完成,已取消"' in str(dv.formula1) for dv in dvs):
        _r("pass", f"{cid}: 状态下拉")
    else:
        _r("fail", f"{cid}: 缺少状态下拉")
    if ws.freeze_panes:
        _r("pass", f"{cid}: 冻结窗格={ws.freeze_panes}")
    else:
        _r("fail", f"{cid}: 无冻结窗格")

    # 反馈-财务部：T002数量=25
    fb = check_file_readable(d, "W1_反馈_财务部.xlsx")
    if fb:
        qty = fb.active.cell(2, 5).value
        if qty == 25:
            _r("pass", f"{cid}: 财务反馈T002数量=25")
        else:
            _r("fail", f"{cid}: 财务反馈T002数量={qty}(应25)")

    # 反馈-工程部：T999公式
    fb2 = check_file_readable(d, "W1_反馈_工程技术部.xlsx")
    if fb2:
        ws2 = fb2.active
        # 找T999所在行
        t999_row = None
        for r in range(2, ws2.max_row + 1):
            if ws2.cell(r, 1).value == "T999":
                t999_row = r
                break
        if t999_row:
            formula = ws2.cell(t999_row, 6).value
            expected = f"=E{t999_row}*100"
            if formula == expected:
                _r("pass", f"{cid}: T999公式={formula}")
            else:
                _r("fail", f"{cid}: T999公式={formula}(应{expected})")
        else:
            _r("fail", f"{cid}: 反馈中找不到T999")

    # 全年更新：T002数量=25, 预算应为25*100=2500或公式
    exp = check_file_readable(d, "W1_预期_全年更新.xlsx")
    if exp:
        ws3 = exp.active
        for r in range(2, ws3.max_row + 1):
            if ws3.cell(r, 1).value == "T002":
                qty = ws3.cell(r, 5).value
                budget = ws3.cell(r, 6).value
                if qty == 25:
                    _r("pass", f"{cid}: 更新版T002数量=25")
                else:
                    _r("fail", f"{cid}: 更新版T002数量={qty}(应25)")
                if budget == 2500 or (isinstance(budget, str) and "=E" in budget and "100" in budget):
                    _r("pass", f"{cid}: 更新版T002预算={budget}(正确)")
                else:
                    _r("fail", f"{cid}: 更新版T002预算={budget}(应2500或=E*100公式)")
                break

def check_w3_headers(d):
    """W3 双行表头结构检查"""
    for f in d.glob("W3_拆分_*.xlsx"):
        wb = load_workbook(f)
        ws = wb.active
        # 检查是否有合并单元格（双行表头应有）
        merges = list(ws.merged_cells.ranges)
        if len(merges) >= 2:
            _r("pass", f"W3/{f.stem}: 保留合并表头({len(merges)}个)")
        else:
            _r("fail", f"W3/{f.stem}: 双行表头结构丢失(仅{len(merges)}个合并)")

def check_w5_truth(d):
    """W5 真值核对"""
    wb = load_workbook(d / "W5_原始_Sheet月份.xlsx")
    actual = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        count = sum(1 for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value)
        actual[sn] = count
    cj = json.loads((d / "case.json").read_text(encoding="utf-8"))
    claimed = cj.get("ground_truth", {}).get("tasks_per_sheet", [])
    for i, (sn, cnt) in enumerate(actual.items()):
        if i < len(claimed):
            if cnt == claimed[i]:
                _r("pass", f"W5/{sn}: {cnt}条任务")
            else:
                _r("fail", f"W5/{sn}: 实际{cnt}条, JSON写的{claimed[i]}条")
        else:
            _r("warn", f"W5/{sn}: JSON未覆盖此Sheet")

def check_f23(d):
    """F23 验证规则与数据范围一致性"""
    wb = load_workbook(d / "F23_条件格式验证.xlsx")
    ws = wb.active
    # 检查完成率数据范围
    rates = [ws.cell(r, 2).value for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value is not None]
    if rates:
        min_v, max_v = min(rates), max(rates)
        if min_v < 1:
            _r("warn", f"F23: 完成率范围{min_v}~{max_v}, 整数验证1-100不匹配")
    dvs = ws.data_validations.dataValidation
    for dv in dvs:
        if dv.type == "whole":
            _r("pass", f"F23: 存在整数验证(type={dv.type}, range={dv.sqref})")

def check_d41_default(d):
    """D41 默认预期应为保留空行"""
    cj = json.loads((d / "case.json").read_text(encoding="utf-8"))
    exp = cj.get("expected", "")
    if "保留" in exp:
        _r("pass", "D41: 默认预期=保留空行")
    else:
        _r("fail", "D41: 默认预期应为保留（原生复刻要求）")

def main():
    print("=" * 60)
    print("泛用性验收测试集自检 v2（按 case.json 发现全部案例）")
    print("=" * 60)

    cases = discover_cases()
    print(f"\n发现 {len(cases)} 个案例\n")

    for cid, d, cj in cases:
        print(f"[{cid}] type={cj.get('type','?')}")

        # 文件可读性
        xlsx_files = sorted(d.glob("*.xlsx"))
        if not xlsx_files:
            _r("fail", f"{cid}: 无Excel文件")
            continue
        for f in xlsx_files:
            check_file_readable(d, f.name)

        # 分类检查
        check_data(d, cid, cj)
        check_fidelity(d, cid, cj)
        check_workflow(d, cid, cj)
        check_ambiguity(d, cid, cj)

    # 重点案例深度检查
    print("\n--- 重点案例深度检查 ---\n")
    w1 = ROOT / "W1"
    if w1.exists(): check_w1_details(w1)
    w3 = ROOT / "W3"
    if w3.exists(): check_w3_headers(w3)
    w5 = ROOT / "W5"
    if w5.exists(): check_w5_truth(w5)
    f23 = ROOT / "F23"
    if f23.exists(): check_f23(f23)
    d41 = ROOT / "D41"
    if d41.exists(): check_d41_default(d41)

    # 汇总
    print(f"\n{'=' * 60}")
    print(f"通过: {R['pass']} | 失败: {R['fail']} | 跳过: {R['skip']} | 警告: {R['warn']}")
    if R['fail'] > 0:
        print(f"\n失败项:")
        for f in DETAIL['fail']:
            print(f"  ✗ {f}")
    print(f"\n未验证项（需人工在Excel/WPS中打开确认）:")
    for w in DETAIL['warn']:
        print(f"  ⚠ {w}")

    if R['fail'] > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
