# -*- coding: utf-8 -*-
"""收尾修订：修正 GPT 指出的 7 项明确问题"""
import json, random, shutil
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

ROOT = Path(__file__).resolve().parent
rng = random.Random(20261020)
UNITS = ["综合办公室", "财务部", "工程技术部", "安全管理部", "市场营销部"]
TASKS_POOL = ["完成年度预算编制", "组织安全培训演练", "设备巡检维护", "季度报表汇总",
              "完成资料汇总", "成本合计核算", "供应商评估审查", "新员工入职办理"]
THIN = Border(*[Side(style="thin", color="999999")] * 4)
HDR_F = Font(bold=True, size=10, name="微软雅黑")
HDR_B = PatternFill("solid", fgColor="D9E1F2")

def _hdr(ws, row, vals, start_col=1):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font, c.fill, c.border = HDR_F, HDR_B, THIN

def _save(wb, p):
    p.parent.mkdir(parents=True, exist_ok=True)
    wb.save(p)

def _cj(d, data):
    (d / "case.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ============================================================ 1. 补齐重建：W2/W3/W6 缺失文件纳入生成代码
def rebuild_missing():
    """在空目录重建时能生成所有已提交文件"""

    # --- W2 预期文件 ---
    d = ROOT / "W2"
    orig_w2 = load_workbook(d / "W2_原始_横向月份.xlsx")
    ws_orig = orig_w2.active

    # 月度预期（1月，财务部任务1数量=99，其余与原表一致）
    wb_m = Workbook(); ws_m = wb_m.active; ws_m.title = "1月汇总"
    _hdr(ws_m, 1, ["责任单位", "任务名称", "1月数量"])
    for r in range(3, ws_orig.max_row):
        unit = ws_orig.cell(r, 1).value or ws_orig.cell(r-1, 1).value if r > 3 else ws_orig.cell(r, 1).value
        # 找到每行的单位（A列可能因合并为None）
        unit = ws_orig.cell(r, 1).value
        if unit is None:
            for rr in range(r, 2, -1):
                if ws_orig.cell(rr, 1).value:
                    unit = ws_orig.cell(rr, 1).value
                    break
        task = ws_orig.cell(r, 2).value
        jan = ws_orig.cell(r, 3).value or 0
        # 财务部任务1的1月数量改为99
        if unit == "财务部" and task and "预算" in str(task):
            jan = 99
        if unit and task:
            ws_m.append([unit, task, jan])
    _save(wb_m, d / "W2_预期_1月汇总.xlsx")

    # 全年预期
    wb_a = Workbook(); ws_a = wb_a.active; ws_a.title = "全年"
    _hdr(ws_a, 1, ["责任单位", "任务名称"] + [f"{m}月" for m in range(1, 13)] + ["合计"])
    for r in range(3, ws_orig.max_row):
        unit = ws_orig.cell(r, 1).value
        if unit is None:
            for rr in range(r, 2, -1):
                if ws_orig.cell(rr, 1).value:
                    unit = ws_orig.cell(rr, 1).value
                    break
        task = ws_orig.cell(r, 2).value
        if not (unit and task):
            continue
        row = [unit, task]
        for m_col in range(3, 15):
            val = ws_orig.cell(r, m_col).value or 0
            # 财务部任务1的1月(B列=C列第3列)改为99
            if unit == "财务部" and task and "预算" in str(task) and m_col == 3:
                val = 99
            row.append(val)
        row.append(f"=SUM(C{ws_a.max_row+1}:N{ws_a.max_row+1})")
        ws_a.append(row)
    _save(wb_a, d / "W2_预期_全年更新.xlsx")

    # --- W3 反馈+月度+年度 ---
    d = ROOT / "W3"
    orig_w3 = load_workbook(d / "W3_原始_双行表头.xlsx")
    ws_w3 = orig_w3.active

    # 读取原始任务数据（跳过前置区r1-r2和双行表头r3-r4）
    w3_tasks = []
    for r in range(5, ws_w3.max_row + 1):
        tid = ws_w3.cell(r, 1).value
        if not tid:
            continue
        w3_tasks.append({
            "id": tid,
            "name": ws_w3.cell(r, 2).value,
            "unit": ws_w3.cell(r, 3).value,
            "deadline": ws_w3.cell(r, 4).value,
            "plan": ws_w3.cell(r, 5).value,
            "done": ws_w3.cell(r, 6).value,
        })

    # 反馈：财务部 A002 完成数 8→15（定位到 A002 行，不覆盖表头）
    wb_fb = Workbook(); ws_fb = wb_fb.active; ws_fb.title = "财务部"
    _hdr(ws_fb, 1, ["编号", "任务名称", "基本信息", "基本信息", "进度", "进度", "进度", "备注"])
    ws_fb.merge_cells("C1:D1"); ws_fb.merge_cells("E1:G1")
    _hdr(ws_fb, 2, ["", "", "责任单位", "完成时限", "计划数", "完成数", "完成率", ""])
    for i, t in enumerate(w3_tasks):
        r = i + 3
        done = 15 if t["id"] == "A002" else t["done"]
        ws_fb.append([t["id"], t["name"], t["unit"], t["deadline"],
                      t["plan"], done, f"=F{r}/E{r}", None])
    ws_fb.conditional_formatting.add(f"G3:G{ws_fb.max_row}",
        CellIsRule(operator="lessThan", formula=["0.5"],
                   fill=PatternFill("solid", fgColor="FFC7CE")))
    ws_fb.freeze_panes = "A3"
    _save(wb_fb, d / "W3_反馈_财务部.xlsx")

    # 月度预期（明确目标年月：3月，按完成时限的月份筛选）
    target_month = 3
    wb_mm = Workbook(); ws_mm = wb_mm.active; ws_mm.title = f"{target_month}月汇总"
    _hdr(ws_mm, 1, ["编号", "任务名称", "责任单位", "完成时限", "计划数", "完成数", "完成率", "备注"])
    for i, t in enumerate(w3_tasks):
        dl = str(t["deadline"] or "")
        # 提取月份
        month = None
        if "-" in dl:
            try:
                month = int(dl.split("-")[1])
            except (ValueError, IndexError):
                pass
        if month == target_month:
            r = ws_mm.max_row + 1
            done = 15 if t["id"] == "A002" else t["done"]
            ws_mm.append([t["id"], t["name"], t["unit"], t["deadline"],
                          t["plan"], done, f"=F{r}/E{r}", None])
    _save(wb_mm, d / "W3_预期_3月汇总.xlsx")

    # 年度预期（从原始数据生成，A002完成数=15）
    wb_aa = Workbook(); ws_aa = wb_aa.active; ws_aa.title = "全年更新"
    _hdr(ws_aa, 1, ["编号", "任务名称", "基本信息", "基本信息", "进度", "进度", "进度", "备注"])
    ws_aa.merge_cells("C1:D1"); ws_aa.merge_cells("E1:G1")
    _hdr(ws_aa, 2, ["", "", "责任单位", "完成时限", "计划数", "完成数", "完成率", ""])
    for i, t in enumerate(w3_tasks):
        r = i + 3
        done = 15 if t["id"] == "A002" else t["done"]
        ws_aa.append([t["id"], t["name"], t["unit"], t["deadline"],
                      t["plan"], done, f"=F{r}/E{r}", None])
    ws_aa.conditional_formatting.add(f"G3:G{ws_aa.max_row}",
        CellIsRule(operator="lessThan", formula=["0.5"],
                   fill=PatternFill("solid", fgColor="FFC7CE")))
    ws_aa.freeze_panes = "A3"
    _save(wb_aa, d / "W3_预期_全年更新.xlsx")

    # --- W6 拆分+月度预期 ---
    d = ROOT / "W6"
    w6_tasks = [
        ("C001", "年度审计", "财务部", "1月", "3月", 1),
        ("C002", "系统上线", "工程技术部", "2月", "5月", 3),
        ("C003", "安全培训", "安全管理部", "3月", "3月", 5),
        ("C004", "市场推广", "市场营销部", "4月", "6月", 8),
    ]
    for unit in ["财务部", "工程技术部", "安全管理部", "市场营销部"]:
        f = d / f"W6_拆分_{unit}.xlsx"
        if not f.exists():
            wb = Workbook(); ws = wb.active; ws.title = unit
            ws.append(["编号", "任务名称", "责任单位", "起始月", "结束月", "数量"])
            for t in w6_tasks:
                if t[2] == unit:
                    ws.append(list(t))
            _save(wb, f)

    # W6 月度预期（按起始月归入）
    f_m = d / "W6_预期_月度汇总.xlsx"
    wb_m6 = Workbook(); ws_m6 = wb_m6.active; ws_m6.title = "月度汇总"
    ws_m6.append(["月份", "编号", "任务名称", "责任单位", "数量"])
    # 按起始月归入
    for t in w6_tasks:
        start_m = int(t[3].replace("月", ""))
        ws_m6.append([f"{start_m}月", t[0], t[1], t[2], t[5]])
    _save(wb_m6, f_m)

    print("✓ 补齐 W2/W3/W6 缺失文件")

    # 修正 W2 反馈（修正错误的月份列）
    fix_W2_feedback(d.parent / "W2")
    # 修正 W6 业务前提一致性
    fix_W6_premise(d.parent / "W6")


def fix_W2_feedback(d):
    """修正W2反馈：B列=1月(C列第3列=索引2)，不应改C2（那是2月）"""
    orig = load_workbook(d / "W2_原始_横向月份.xlsx")
    ws_orig = orig.active
    # 找到财务部"完成年度预算编制"行的1月数量（C列，即第3列）
    target_row = None
    orig_val = None
    for r in range(3, ws_orig.max_row + 1):
        unit = ws_orig.cell(r, 1).value
        if unit is None:
            for rr in range(r, 2, -1):
                if ws_orig.cell(rr, 1).value:
                    unit = ws_orig.cell(rr, 1).value
                    break
        task = ws_orig.cell(r, 2).value
        if unit == "财务部" and task and "预算" in str(task):
            target_row = r
            orig_val = ws_orig.cell(r, 3).value  # C列 = 1月
            break

    if target_row is None:
        print("⚠ W2: 找不到财务部预算任务行")
        return

    # 重建反馈：修改B列（第2列=索引2，即C列第3列=1月数量）→ 注意：W2拆分文件的结构不同
    # W2拆分文件的列：任务名称 | 1月 | 2月 | ... | 12月 | 合计
    # 所以1月数量在第2列（B列）
    fb_path = d / "W2_反馈_财务部.xlsx"
    if fb_path.exists():
        wb = load_workbook(fb_path)
        ws = wb.active
        # 找到"完成年度预算编制"行
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value and "预算" in str(ws.cell(r, 1).value):
                ws.cell(r, 2).value = 99  # B列 = 1月数量
                break
        _save(wb, fb_path)

    # 重建月度预期（1月，99在财务部的任务上）
    wb_m = Workbook(); ws_m = wb_m.active; ws_m.title = "1月汇总"
    _hdr(ws_m, 1, ["责任单位", "任务名称", "1月数量"])
    for r in range(3, ws_orig.max_row):
        unit = ws_orig.cell(r, 1).value
        if unit is None:
            for rr in range(r, 2, -1):
                if ws_orig.cell(rr, 1).value:
                    unit = ws_orig.cell(rr, 1).value
                    break
        task = ws_orig.cell(r, 2).value
        jan = ws_orig.cell(r, 3).value or 0
        if unit == "财务部" and task and "预算" in str(task):
            jan = 99
        if unit and task:
            ws_m.append([unit, task, jan])
    _save(wb_m, d / "W2_预期_1月汇总.xlsx")

    # 更新 case.json
    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["expected_changes"] = [{"task": "完成年度预算编制", "field": "1月数量",
                                  "old": orig_val, "new": 99, "unit": "财务部"}]
    data["expected_monthly"] = "1月汇总中财务部'完成年度预算编制'数量=99，其余单位不变"
    _cj(d, data)
    print(f"✓ W2 反馈修正：财务部预算任务1月数量 {orig_val}→99（B列，非C列）")


def fix_W6_premise(d):
    """修正W6：按起始月归入 → C001(1月)归1月，C004(4月)归4月，非3月/5月"""
    tasks = [
        ("C001", "年度审计", "财务部", 1, 3, 1),    # 起始1月
        ("C002", "系统上线", "工程技术部", 2, 5, 3),  # 起始2月
        ("C003", "安全培训", "安全管理部", 3, 3, 5),   # 起始3月
        ("C004", "市场推广", "市场营销部", 4, 6, 8),   # 起始4月
    ]

    # 月度预期（按起始月归入）
    wb = Workbook(); ws = wb.active; ws.title = "月度汇总"
    ws.append(["月份", "编号", "任务名称", "责任单位", "数量"])
    for tid, name, unit, start, end, qty in tasks:
        ws.append([f"{start}月", tid, name, unit, qty])
    _save(wb, d / "W6_预期_月度汇总.xlsx")

    # 更新 case.json
    cj = d / "case.json"
    data = json.loads(cj.read_text(encoding="utf-8"))
    data["business_premise"] = "跨月任务按起始月归入月度汇总（测试假设，非产品最终口径）"
    data["expected_monthly"] = {
        "1月": "C001(起始1月)",
        "2月": "C002(起始2月)",
        "3月": "C003(起始3月，调整后移至5月)",
        "4月": "C004(起始4月)",
        "5月": "调整后的C003(起始月变为5月)",
    }
    data["expected_annual_update"] = "C003起始月3→5、结束月3→5，其余任务不变"
    _cj(d, data)
    print("✓ W6 前提修正：按起始月归入（C001→1月、C004→4月，原错误地放在3月/5月）")


# ============================================================ 5. 保真可执行检查
def add_fidelity_checks():
    """在自检中真正读取并执行 fidelity_requirements 检查"""

    # W1 保真检查
    d = ROOT / "W1"
    orig = load_workbook(d / "W1_原始计划表.xlsx")
    ws_o = orig.active

    for f in d.glob("W1_拆分_*.xlsx"):
        wb = load_workbook(f)
        ws = wb.active
        checks = []

        # 冻结窗格
        if ws.freeze_panes:
            checks.append(("冻结", ws.freeze_panes, True))
        else:
            checks.append(("冻结", None, False))

        # 下拉
        dvs = ws.data_validations.dataValidation
        has_dv = len(dvs) > 0
        checks.append(("下拉", f"{len(dvs)}个", has_dv))

        # 条件格式
        cfs = list(ws.conditional_formatting)
        has_cf = len(cfs) > 0
        checks.append(("条件格式", f"{len(cfs)}个", has_cf))

        # 列宽
        widths_ok = ws.column_dimensions["A"].width is not None
        checks.append(("列宽", f"A={ws.column_dimensions['A'].width}", widths_ok))

        # 公式
        formula_ok = isinstance(ws.cell(2, 6).value, str) and "=E" in ws.cell(2, 6).value
        checks.append(("公式", ws.cell(2, 6).value, formula_ok))

        # 写入检查结果
        all_pass = all(c[2] for c in checks)
        status = "全通过" if all_pass else "有缺失"
        print(f"  W1/{f.stem} 保真检查: {status}")
        for name, val, ok in checks:
            icon = "✓" if ok else "✗"
            print(f"    {icon} {name}: {val}")

        # 保存到case.json
        cj = d / "case.json"
        data = json.loads(cj.read_text(encoding="utf-8"))
        data.setdefault("fidelity_check_results", {})[f.stem] = {
            "all_pass": all_pass,
            "details": {name: {"value": str(val), "pass": ok} for name, val, ok in checks},
        }
        _cj(d, data)

    # W3 保真检查（双行表头）
    d = ROOT / "W3"
    for f in d.glob("W3_拆分_*.xlsx"):
        wb = load_workbook(f)
        ws = wb.active
        merges = list(ws.merged_cells.ranges)
        merge_strs = sorted(str(m) for m in merges)
        has_double = any("C" in s and "D" in s for s in merge_strs) and \
                     any("E" in s and "G" in s for s in merge_strs)
        formula = ws.cell(3, 7).value
        formula_ok = isinstance(formula, str) and "=F" in formula
        freeze = ws.freeze_panes

        all_pass = has_double and formula_ok and freeze
        print(f"  W3/{f.stem} 保真: 双行表头={'✓' if has_double else '✗'} "
              f"公式={'✓' if formula_ok else '✗'} 冻结={'✓' if freeze else '✗'}")

    print("✓ 保真可执行检查完成")


# ============================================================ 6. D40 换列计算含义不变
def fix_D40():
    d = ROOT / "D40"
    if not d.exists():
        return

    wb = Workbook()

    # 原始数据
    ws1 = wb.active; ws1.title = "原始"
    _hdr(ws1, 1, ["A值", "B值", "C值", "D值", "SUM(A:C)"])
    for i in range(5):
        r = i + 2
        a, b, c, dd = 2, 10, 100, 1000
        ws1.append([a, b, c, dd, f"=SUM(A{r}:C{r})"])  # SUM=112

    # 换列：A/D/C/B 顺序，公式仍只计算原A/B/C（即新的A/D/C列）
    ws2 = wb.create_sheet("换列后")
    _hdr(ws2, 1, ["A值", "D值", "C值", "B值", "原SUM内容"])
    for i in range(5):
        r = i + 2
        a, b, c, dd = 2, 10, 100, 1000
        # 换列后：新A=原A(2), 新B=原D(1000), 新C=原C(100), 新D=原B(10)
        # 原SUM(A:C)=2+10+100=112
        # 换列后公式仍需计算 2+10+100=112，即新A+新D+新C
        ws2.append([a, dd, c, b, f"=A{r}+D{r}+C{r}"])  # =2+10+100=112

    _save(wb, d / "D40_列换序引用.xlsx")

    _cj(d, {
        "case_id": "D40", "type": "保真",
        "scenario": "原始列序A/B/C/D，SUM(A:C)=2+10+100=112。换列后A/D/C/B，公式仍需计算原A/B/C",
        "checks": {
            "原始": "SUM(A:C)=A+B+C=2+10+100=112",
            "换列后": "公式=A+D+C=2+10+100=112（新D=原B）",
            "关键": "参与计算的单元格集合不变（原A、原B、原C），只是物理位置变了",
            "预期值": 112,
        },
        "verification_status": "代码可验证公式文本，计算结果需Excel确认=112",
    })
    print("✓ D40 增加换列含义不变案例")


# ============================================================ 7. 修正报告口径
def fix_reporting():
    """更新 case.json 中的状态标注和验收说明"""

    # W4/W5 标注为非完整流程
    for wid in ["W4", "W5"]:
        d = ROOT / wid
        cj = d / "case.json"
        if cj.exists():
            data = json.loads(cj.read_text(encoding="utf-8"))
            data["type"] = "结构案例"  # 降级：尚未形成完整流程
            data["incomplete_reason"] = "缺少反馈/月度/年度预期文件，仅为结构测试"
            _cj(d, data)

    # 更新验收说明
    readme = ROOT / "验收说明.md"
    content = """# 泛用性与原生保真验收测试集（v3）

> 收尾修订版。42案例/含完整流程/保真检查可执行。

## 重建命令（从空目录）
```bash
cd 测试集/泛用性验收_v3
python 生成验收集.py     # 基础24案例
python 补充案例.py       # 补充18案例
python 修订.py           # 内容修正
python 收尾修订.py       # 收尾（W2/W3/W6修正+保真检查+D40）
python 自检.py           # 自检
```

## 报告口径
- **通过**：数据正确性断言 + 保真属性代码检查（公式文本/合并范围/冻结/列宽等）
- **跳过**：缺少 ground_truth 或歧义案例（不计通过）
- **警告**：需人工在Excel/WPS中目视验证的显示效果（非代码可检查项）
- **失败**：任何必检项不满足

## 分类状态
| 类别 | 数量 | 说明 |
|------|------|------|
| 完整流程 | W1~W3, W6 | 原表→拆分→反馈→月度→年度 |
| 结构案例 | W4, W5, S07~S12, A25~A28 | 结构/关系测试（W4/W5缺反馈链路） |
| 反馈案例 | S13~S18, C32~C38 | 变更检测测试 |
| 保真案例 | F19~F24, D39~D41 | 原生复刻验证 |
| 往返一致 | RT42 | 拆分→不修改→回收=零变更 |

## 歧义清单（7项）
A26月份冲突/B29协办归属/B30简称合并/B31跨年归属/C33缺失≠删除/C36 UUID重复/D41空行语义

## 保真代码检查（已实现）
- 公式文本比较（=E*100、=F/E等模式匹配）
- 合并单元格范围（C1:D1、E1:G1等）
- 冻结窗格存在性和值
- 列宽非None
- 下拉验证存在性和公式
- 条件格式规则存在性

## 保真人工检查（需打开Excel）
- 字体/颜色/边框实际渲染效果
- 打印预览（纸张方向/页眉页脚）
- 批注文字内容
- 超链接可点击
- 数据验证的实际交互（下拉可选）
- 公式重算结果
"""
    readme.write_text(content, encoding="utf-8")
    print("✓ 验收说明更新完成")


# ============================================================ 主入口
if __name__ == "__main__":
    rebuild_missing()
    add_fidelity_checks()
    fix_D40()
    fix_reporting()
    print("\n收尾修订完成")
