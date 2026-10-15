# -*- coding: utf-8 -*-
"""写出层单元测试 + sidecar 协议测试。"""
import json
import subprocess
import sys
from pathlib import Path

from engine.tools.excel_merge.writer import big_number_guard, sanitize_sheet_name

ROOT = Path(__file__).resolve().parent.parent


def test_big_number_guard():
    assert big_number_guard(123456789012345) == (123456789012345, False)   # 15位安全
    assert big_number_guard(1234567890123456) == ("1234567890123456", True)  # 16位转文本
    assert big_number_guard(-1234567890123456) == ("-1234567890123456", True)
    assert big_number_guard(1.23e16)[1] is True
    assert big_number_guard(3.14) == (3.14, False)
    assert big_number_guard(True) == (True, False)                        # 布尔不算数字
    assert big_number_guard("文本") == ("文本", False)


def test_sanitize_sheet_name():
    used = set()
    assert sanitize_sheet_name("销售", used) == "销售"
    assert sanitize_sheet_name("销售", used) == "销售(2)"
    assert sanitize_sheet_name("销售", used) == "销售(3)"
    assert sanitize_sheet_name("a:b/c?d*e[f]", set()) == "a b c d e f"    # 非法字符替换
    assert len(sanitize_sheet_name("字" * 50, set())) <= 31               # 31字符上限


def test_sidecar_protocol():
    """ping / scan / merge 全链路（JSON 行协议）。"""
    d = ROOT / "测试集" / "场景E_表头智能识别"
    msgs = [
        {"cmd": "ping"},
        {"tool": "excel_merge", "cmd": "scan", "paths": [str(d / "报表_甲.xlsx")]},
        {"tool": "excel_merge", "cmd": "merge",
         "paths": [str(d / "报表_甲.xlsx"), str(d / "报表_乙.xlsx")],
         "out": str(d.parent.parent / "_sidecar_out.xlsx"), "options": {"mode": "single"}},
        {"cmd": "quit"},
    ]
    p = subprocess.run(
        [sys.executable, "-m", "engine.sidecar"],
        input="\n".join(json.dumps(m, ensure_ascii=False) for m in msgs),
        capture_output=True, text=True, encoding="utf-8", cwd=str(ROOT), timeout=120,
    )
    events = [json.loads(l) for l in p.stdout.splitlines() if l.strip()]
    assert any(e.get("evt") == "pong" for e in events)
    scan_evt = next(e for e in events if e.get("evt") == "report")
    det = scan_evt["report"]["header_detections"][0]
    assert det["header_row"] == 3                       # 表头行判定随协议透出
    merge_evt = [e for e in events if e.get("evt") == "report"][1]
    assert merge_evt["report"]["totals"]["rows_written"] == 26  # 甲14 + 乙12
    out = Path(merge_evt["report"]["output"])
    assert out.exists() and out.stat().st_size > 0
    out.unlink()
    (out.parent / (out.stem + "_运行日志.log")).unlink(missing_ok=True)
