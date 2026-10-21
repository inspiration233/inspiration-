# -*- coding: utf-8 -*-
"""补充案例：填补初版31项覆盖缺口"""
import json, random, shutil, datetime as dt
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent
rng = random.Random(20261021)
UNITS = ["综合办公室", "财务部", "工程技术部", "安全管理部", "市场营销部"]
TASKS = ["完成年度预算", "组织安全培训", "设备巡检维护", "季度报表汇总",
         "完成资料汇总", "成本合计核算", "供应商评估", "新员工入职"]
THIN = Border(*[Side(style="thin", color="999999")] * 4)
HDR_F = Font(bold=True, size=10)
HDR_B = PatternFill("solid", fgColor="D9E1F2")

def _hdr(ws, row, vals):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=i+1, value=v)
        c.font, c.fill, c.border = HDR_F, HDR_B, THIN

def _save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

def _cj(d, data):
    (d / "case.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def _mk_orig(d, tasks, cols=("编号","任务名称","责任单位","月份","数量","备注")):
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, list(cols))
    for t in tasks:
        ws.append(t)
    _save(wb, d / "原始.xlsx")
    return wb

# ============================================================ A 结构补充

def A25(d):
    """季度+全年月份"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["任务", "第一季度", "第二季度", "第三季度", "第四季度", "全年"])
    for i, t in enumerate(TASKS[:4]):
        r = i + 2
        ws.append([t, rng.randint(1,20), rng.randint(1,20), rng.randint(1,20),
                    rng.randint(1,20), f"=SUM(B{r}:E{r})"])
    _save(wb, d / "A25_季度全年.xlsx")
    _cj(d, {"case_id":"A25","type":"结构","scenario":"季度列(第一~第四季度)+全年合计公式",
            "ambiguity":["季度应展开为月还是保持季度粒度？"]})

def A26(d):
    """月份在标题但与单元格冲突"""
    wb = Workbook(); ws = wb.active
    ws.append(["2026年3月工作计划"]); ws.merge_cells("A1:D1")
    _hdr(ws, 2, ["任务","单位","月份","数量"])
    # 标题说3月, 但数据里有1月和3月 → 冲突
    ws.append(["任务A","财务部","3月",10])
    ws.append(["任务B","工程部","1月",5])   # ← 标题说3月但这行是1月
    ws.append(["任务C","安全部","3月",8])
    _save(wb, d / "A26_月份冲突.xlsx")
    _cj(d, {"case_id":"A26","type":"结构","scenario":"标题写'3月'但数据含1月任务",
            "ambiguity":["以标题为准还是以单元格为准？1月的任务B应归入哪个月？"]})

def A27(d):
    """首条任务全是文字（无数字列特征）"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务描述","负责科室","完成标准","备注"])
    ws.append(["RW001","编制年度安全工作要点并印发执行","综合科","印发正式文件","涉及多个部门会签"])
    ws.append(["RW002","组织安全生产月启动仪式","安监科","活动方案+照片",""])
    ws.append(["RW003","修订完善应急预案体系","应急科","预案通过评审",""])
    _save(wb, d / "A27_全文字任务.xlsx")
    _cj(d, {"case_id":"A27","type":"结构","scenario":"首条任务全是文字，无数字特征列",
            "note":"检验表头探测在纯文本数据下的表现"})

def A28(d):
    """空表头列"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务",None,"数量"])  # 第3列无表头
    ws.append(["E001","任务A","手工标注",10])
    ws.append(["E002","任务B",None,20])
    _save(wb, d / "A28_空表头.xlsx")
    _cj(d, {"case_id":"A28","type":"结构","scenario":"第3列表头为空但有数据",
            "expected":"该列保留为'未命名'列，数据不丢失"})

# ============================================================ B 关系补充

def B29(d):
    """牵头与协办"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["任务","牵头单位","协办单位","数量"])
    ws.append(["联合检查","安全管理部","工程技术部,综合办公室",20])
    ws.append(["年度审计","财务部","综合办公室",1])
    ws.append(["系统升级","工程技术部","财务管理部,安全管理部",3])
    _save(wb, d / "B29_牵头协办.xlsx")
    _cj(d, {"case_id":"B29","type":"关系","scenario":"牵头/协办分离，协办可多个单位",
            "ambiguity":["拆分时协办单位是否也收到该任务？数量归属牵头还是均分？"]})

def B30(d):
    """单位简称vs全称"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["任务","单位","数量"])
    ws.append(["任务A","综合办",10])      # 简称
    ws.append(["任务B","综合办公室",20])   # 全称
    ws.append(["任务C","财务",5])         # 简称
    ws.append(["任务D","财务部",8])       # 全称
    _save(wb, d / "B30_单位简称.xlsx")
    _cj(d, {"case_id":"B30","type":"关系","scenario":"'综合办'和'综合办公室'是同一单位的不同写法",
            "ambiguity":["是否应合并为同一单位？需人工确认简称映射关系"]})

def B31(d):
    """跨年份任务"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["任务","单位","起始年月","结束年月","数量"])
    ws.append(["五年规划编制","综合办","2025-01","2026-12",1])
    ws.append(["年度审计","财务部","2026-01","2026-03",1])
    ws.append(["系统建设","工程部","2025-06","2027-06",5])
    _save(wb, d / "B31_跨年份.xlsx")
    _cj(d, {"case_id":"B31","type":"关系","scenario":"任务跨年份（2025-2027），非单一年度",
            "ambiguity":["跨年任务归入哪一年的年度表？起始年？结束年？每年都算？"]})

# ============================================================ C 反馈补充（15个）

def _base_tasks():
    return [("B001","任务A","综合办公室","1月",10,None),
            ("B002","任务B","财务部","2月",20,None),
            ("B003","任务C","工程技术部","3月",30,"原有备注"),
            ("B004","任务D","安全管理部","4月",40,None),
            ("B005","任务E","市场营销部","5月",50,None)]

def C32(d):
    """改任务名称+改期限+改单位+改备注"""
    orig = _base_tasks()
    _mk_orig(d, orig, ("编号","任务名称","责任单位","月份","数量","备注"))
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务名称","责任单位","月份","数量","备注"])
    for t in orig:
        tid, name, unit, month, qty, note = t
        if tid == "B001": name = "任务A（已改名）"
        if tid == "B002": month = "6月"  # 改期限
        if tid == "B003": unit = "财务部"  # 改单位
        if tid == "B004": note = "新增备注信息"
        ws.append([tid, name, unit, month, qty, note])
    _save(wb, d / "反馈.xlsx")
    _cj(d, {"case_id":"C32","type":"反馈",
            "scenario":"4种字段修改：B001改名称/B002改月份/B003改单位/B004加备注",
            "expected_changes": [
                {"task":"B001","field":"任务名称","old":"任务A","new":"任务A（已改名）"},
                {"task":"B002","field":"月份","old":"2月","new":"6月"},
                {"task":"B003","field":"责任单位","old":"工程技术部","new":"财务部"},
                {"task":"B004","field":"备注","old":None,"new":"新增备注信息"},
            ], "invariant":"B005无变化"})

def C33(d):
    """插入行+删除行"""
    orig = _base_tasks()
    _mk_orig(d, orig)
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务名称","责任单位","月份","数量","备注"])
    # 删除B003, 在B002和B004之间插入新行
    ws.append(["B001","任务A","综合办公室","1月",10,None])
    ws.append(["B002","任务B","财务部","2月",20,None])
    ws.append(["NEW01","插入的新任务","安全管理部","3月",15,None])  # 插入
    ws.append(["B004","任务D","安全管理部","4月",40,None])
    ws.append(["B005","任务E","市场营销部","5月",50,None])
    # B003 被删除
    _save(wb, d / "反馈.xlsx")
    _cj(d, {"case_id":"C33","type":"反馈",
            "scenario":"删除B003, 插入NEW01",
            "expected_changes": [
                {"task":"B003","field":"整行","old":"存在","new":"缺失"},
                {"task":"NEW01","field":"新增","old":None,"new":"插入的新任务"},
            ],
            "ambiguity":["B003缺失是被删除还是未填写？不应自动等同删除"]})

def C34(d):
    """插入列+删除列+改列名"""
    orig = _base_tasks()
    _mk_orig(d, orig)
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务名称","责任单位","月份","数量","优先级","备注"])  # 加了优先级, 删了原备注改为此名
    for t in orig:
        ws.append([t[0], t[1], t[2], t[3], t[4], rng.choice(["高","中","低"]), t[5]])
    _save(wb, d / "反馈.xlsx")
    _cj(d, {"case_id":"C34","type":"反馈",
            "scenario":"新增'优先级'列, 原备注列改为新位置",
            "expected_changes": [{"field":"列结构","old":"6列","new":"7列(+优先级)"}],
            "ambiguity":["新增列如何映射回原表？加在末尾还是按列名对齐？"]})

def C35(d):
    """明确取消+仅缺失"""
    orig = _base_tasks()
    _mk_orig(d, orig)
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务名称","责任单位","月份","数量","备注"])
    ws.append(["B001","任务A","综合办公室","1月",10,None])
    ws.append(["B002","任务B","财务部","2月",20,None])
    ws.append(["B003","任务C","工程技术部","3月",30,"该任务已取消"])  # 明确取消
    # B004 完全缺失（未标注取消）
    ws.append(["B005","任务E","市场营销部","5月",50,None])
    _save(wb, d / "反馈.xlsx")
    _cj(d, {"case_id":"C35","type":"反馈",
            "scenario":"B003备注明确'已取消'; B004整行缺失无说明",
            "expected_changes": [
                {"task":"B003","type":"明确取消","evidence":"备注含'取消'"},
                {"task":"B004","type":"仅缺失","evidence":"无行"},
            ],
            "ambiguity":["B004缺失≠取消。应标记为'待确认'，不应自动从全年表中删除"]})

def C36(d):
    """隐藏标识保留/丢失/重复"""
    orig = [("H001","任务A","财务部",10,"UUID-ABC123"),
            ("H002","任务B","工程部",20,"UUID-DEF456"),
            ("H003","任务C","安全部",30,"UUID-GHI789")]
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务名称","责任单位","数量","隐藏UUID"])
    ws.column_dimensions["E"].hidden = True  # 隐藏UUID列
    for t in orig:
        ws.append(list(t))
    _save(wb, d / "原始.xlsx")

    wb2 = Workbook(); ws2 = wb2.active
    _hdr(ws2, 1, ["编号","任务名称","责任单位","数量","隐藏UUID"])
    ws2.append(["H001","任务A","财务部",12,"UUID-ABC123"])   # 保留
    ws2.append(["H002","任务B","工程部",20,None])              # 丢失！
    ws2.append(["H003","任务C","安全部",30,"UUID-GHI789"])    # 保留
    ws2.append([None,"任务D","市场部",5,"UUID-ABC123"])        # UUID重复！
    _save(wb2, d / "反馈.xlsx")
    _cj(d, {"case_id":"C36","type":"反馈",
            "scenario":"隐藏UUID列：H001保留/H002丢失/H003保留/新行使用了重复UUID",
            "expected_changes": [
                {"task":"H001","uuid":"保留","match":"✓"},
                {"task":"H002","uuid":"丢失","match":"需回退到编号匹配"},
                {"task":"H003","uuid":"保留","match":"✓"},
                {"task":"新行","uuid":"UUID-ABC123(与H001重复)","match":"⚠ 冲突"},
            ],
            "ambiguity":["UUID重复(H001和新行都是ABC123)如何区分？需人工确认"]})

def C37(d):
    """重复导入同一文件"""
    orig = _base_tasks()[:3]
    _mk_orig(d, orig)
    wb2 = Workbook(); ws2 = wb2.active
    _hdr(ws2, 1, ["编号","任务名称","责任单位","月份","数量","备注"])
    for t in orig:
        ws2.append(list(t))
    _save(wb2, d / "反馈.xlsx")
    # case.json 说明：同一文件导入两次
    _cj(d, {"case_id":"C37","type":"反馈",
            "scenario":"同一反馈文件被导入两次",
            "expected":"第二次导入应检测到重复（文件级或行级去重），数据不应翻倍",
            "invariant":"最终结果数量不变"})

def C38(d):
    """旧月份文件混入"""
    # 原始表是2026年10月的
    orig = [("O001","十月任务A","财务部","10月",10),
            ("O002","十月任务B","工程部","10月",20)]
    _mk_orig(d, orig, ("编号","任务名称","责任单位","月份","数量"))
    # 混入一个9月的旧文件
    wb2 = Workbook(); ws2 = wb2.active; ws2.title = "9月旧文件"
    _hdr(ws2, 1, ["编号","任务名称","责任单位","月份","数量"])
    ws2.append(["S001","九月任务C","财务部","9月",5])
    _save(wb2, d / "旧月份文件.xlsx")
    _cj(d, {"case_id":"C38","type":"反馈",
            "scenario":"10月原始表回收时混入了一个9月的旧反馈文件",
            "expected":"S001(9月任务)不应出现在10月汇总中",
            "invariant":"非当月数据不应污染当月汇总"})

# ============================================================ D 保真补充

def D39(d):
    """命名范围"""
    wb = Workbook()
    ws = wb.active; ws.title = "数据"
    _hdr(ws, 1, ["编号","数量","金额"])
    for i in range(5):
        ws.append([f"NM{i+1}", i+1, (i+1)*100])
    wb.defined_names.add(DefinedName("数量区域", attr_text=f"数据!$B$2:$B$6"))
    wb.defined_names.add(DefinedName("总额", attr_text=f"数据!$C$2:$C$6"))
    ws2 = wb.create_sheet("计算")
    ws2.append(["数量合计", "=SUM(数量区域)"])
    ws2.append(["金额合计", "=SUM(总额)"])
    _save(wb, d / "D39_命名范围.xlsx")
    _cj(d, {"case_id":"D39","type":"保真",
            "checks":{"命名范围": "数量区域=B2:B6, 总额=C2:C6",
                      "公式": "=SUM(数量区域)和=SUM(总额)引用命名范围"}})

def D40(d):
    """列换序后引用范围不连续"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["A值","B值","C值","D值","合计"])
    for i in range(5):
        ws.append([i+1, i*10, i*100, i*1000, f"=A{i+2}+C{i+2}"])  # 只引用A和C,跳过B
    # 模拟筛选后删除B列 → 引用范围断裂
    ws2 = wb.create_sheet("筛选后")
    _hdr(ws2, 1, ["A值","C值","D值","合计"])
    for i in range(5):
        ws2.append([i+1, i*100, i*1000, f"=A{i+2}+B{i+2}"])  # B列变了(原来是C)
    _save(wb, d / "D40_列换序引用.xlsx")
    _cj(d, {"case_id":"D40","type":"保真",
            "checks":{"原始": "合计=A+C(跳过B)",
                      "筛选后": "删除B列后合计=A+B(原C变B)",
                      "关键": "公式引用是否随列位置变化而正确调整"}})

def D41(d):
    """数据区必要空行保留"""
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务","数量"])
    ws.append(["G001","任务A",10])
    ws.append([None]*3)  # 空行 — 业务含义：分组分隔
    ws.append(["G002","任务B",20])
    ws.append([None]*3)  # 空行
    ws.append(["G003","任务C",30])
    _save(wb, d / "D41_必要空行.xlsx")
    _cj(d, {"case_id":"D41","type":"保真",
            "scenario":"数据区空行用于视觉分组",
            "ambiguity":["空行是'数据间隙'还是'无意义空行'？保留还是删除？"],
            "expected":"如保留→分组结构不变；如删除→需在报告中说明"})

# ============================================================ 往返一致性

def RT42(d):
    """拆分→不修改→回收 = 零变更"""
    tasks = [(f"RT{i+1:03d}", TASKS[i % len(TASKS)], UNITS[i%5], f"{(i%12)+1}月", (i+1)*5)
             for i in range(10)]
    wb = Workbook(); ws = wb.active
    _hdr(ws, 1, ["编号","任务名称","责任单位","月份","数量"])
    for t in tasks:
        ws.append(list(t))
    _save(wb, d / "RT42_原始.xlsx")
    # 预期：拆分后不修改直接回收 → 变更清单为空
    _cj(d, {"case_id":"RT42","type":"往返一致性",
            "scenario":"拆分→不做任何修改→直接回收",
            "expected":"变更清单为空（零修改），全年表与原始表完全一致",
            "invariant":"所有任务编号、名称、单位、月份、数量不变"})

# ============================================================ 补全 W2~W6 预期输出

def W_outputs(d):
    """为 W2~W6 补充预期汇总/更新文件的说明性case.json（具体Excel因数据复杂由人工核对）"""
    for wid in ["W2", "W3", "W4", "W5", "W6"]:
        case_dir = ROOT / wid
        cj_path = case_dir / "case.json"
        if cj_path.exists():
            data = json.loads(cj_path.read_text(encoding="utf-8"))
            data["expected_monthly_summary"] = f"人工核对：按月汇总{wid}的任务"
            data["expected_annual_update"] = f"人工核对：全年表反映反馈变更"
            data["verification_status"] = "未验证（需在Excel/WPS中人工打开核对）"
            cj_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print("✓ W2~W6 case.json 补充预期说明和未验证标注")


# ============================================================ 主入口
NEW_CASES = {
    "A25": A25, "A26": A26, "A27": A27, "A28": A28,
    "B29": B29, "B30": B30, "B31": B31,
    "C32": C32, "C33": C33, "C34": C34, "C35": C35,
    "C36": C36, "C37": C37, "C38": C38,
    "D39": D39, "D40": D40, "D41": D41,
    "RT42": RT42,
}

if __name__ == "__main__":
    for cid, fn in NEW_CASES.items():
        d = ROOT / cid
        if d.exists():
            target = d.resolve()
            if target.parent != ROOT.resolve() or target.name not in NEW_CASES:
                raise ValueError(f"拒绝重建测试集之外的目录: {target}")
            shutil.rmtree(target)
        fn(d)
        n = len(list(d.glob("*.xlsx"))) + len(list(d.glob("*.json")))
        print(f"✓ {cid} ({n}个文件)")
    W_outputs(ROOT)
    print(f"\n补充 {len(NEW_CASES)} 个案例完成")
