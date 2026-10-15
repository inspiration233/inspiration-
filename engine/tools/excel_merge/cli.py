# -*- coding: utf-8 -*-
"""命令行入口（M1 验收用）。

用法示例：
    python -m engine.tools.excel_merge.cli 文件夹或文件... --scan     # 只体检
    python -m engine.tools.excel_merge.cli 文件夹或文件... -o 合并.xlsx
    python -m engine.tools.excel_merge.cli 1月.xlsx 2月.xls --mode workbook
"""
import argparse
import datetime as dt
import sys
import time
from pathlib import Path

from .merge import Options, merge, scan

SUFFIXES = {".xlsx", ".xls", ".csv"}


def collect(paths):
    out = []
    for p in paths:
        p = Path(p)
        if p.is_dir():
            out += sorted(q for q in p.rglob("*") if q.suffix.lower() in SUFFIXES)
        else:
            out.append(p)
    seen, dedup = set(), []
    for q in out:
        r = q.resolve()
        if r not in seen:
            seen.add(r)
            dedup.append(q)
    return dedup


def main(argv=None):
    ap = argparse.ArgumentParser(description="Excel 合并工具 · 引擎命令行版(M1)")
    ap.add_argument("inputs", nargs="+", help="输入文件或目录（目录递归扫描 xlsx/xls/csv）")
    ap.add_argument("-o", "--out", default=None, help="输出路径（默认当前目录 合并结果_时间戳.xlsx）")
    ap.add_argument("--mode", choices=["single", "workbook", "sheets", "group"],
                    default="single", help="合并模式（默认 single=多文件一张总表）")
    ap.add_argument("--align", choices=["strict", "loose"], default="loose",
                    help="表头对齐（默认 loose 宽容）")
    ap.add_argument("--style", choices=["unified", "preserve"], default="unified",
                    help="输出样式（默认 unified；preserve 保留原样式，较慢）")
    ap.add_argument("--keep-summary", action="store_true", help="保留合计/小计行（默认剔除）")
    ap.add_argument("--source-column", action="store_true", help="追加\"来源\"标记列")
    ap.add_argument("--include-hidden", action="store_true", help="包含隐藏 Sheet（默认跳过）")
    ap.add_argument("--title", action="store_true", help="以第一个文件的标题作为总表标题")
    ap.add_argument("--scan", action="store_true", help="只做体检（表头判定预览），不合并")
    args = ap.parse_args(argv)

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    files = collect(args.inputs)
    if not files:
        print("没有找到可合并的文件（支持 xlsx / xls / csv）")
        return 1
    print(f"输入 {len(files)} 个文件：")
    for f in files:
        print(f"  - {f}")

    opts = Options(mode=args.mode, align=args.align, style=args.style,
                   remove_summary=not args.keep_summary,
                   source_column=args.source_column,
                   include_hidden=args.include_hidden,
                   title_as_header=args.title)

    if args.scan:
        r = scan([str(f) for f in files], opts)
        for d in r["header_detections"]:
            print(f"\n[{d['display']}] 表头第 {d['header_row']} 行 · 高 {d['height']} 行 · "
                  f"{len(d['columns'])} 列")
            if d["title"]:
                print(f"  标题(不入总表): {d['title']}")
            if d["low_confidence"]:
                print("  ⚠ 两个候选表头行得分接近，建议核对")
            for w in d["warnings"]:
                print(f"  ⚠ {w}")
            print(f"  列: {' | '.join(d['columns'])}")
        for f in r["files"]:
            if f["status"] == "failed":
                print(f"\n✗ {f['path']}: {f['reason']}")
        return 0

    out = args.out or (Path.cwd() / f"合并结果_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
    t0 = time.monotonic()

    def prog(fi, total, name, rows):
        print(f"\r  合并中 {fi}/{total} · {name} · 已写入 {rows} 行", end="", flush=True)

    try:
        r = merge([str(f) for f in files], out, opts, progress=prog)
    except Exception as e:
        print(f"\n✗ 合并失败: {e}")
        return 1
    print()
    t = r["totals"]
    print(f"\n✓ 完成 · 耗时 {time.monotonic() - t0:.1f}s")
    print(f"  输出: {Path(r['output']).resolve()}")
    print(f"  文件: 成功 {t['files_ok']} · 失败 {t['files_failed']} · "
          f"写入 {t['rows_written']} 行")
    print(f"  清理: 剔除汇总行 {t['summary_rows_removed']} · 跳过空行 {t['empty_rows_skipped']}"
          + (f" · 隐藏Sheet跳过 {t['hidden_sheets_skipped']}" if t["hidden_sheets_skipped"] else ""))
    if t["big_numbers_converted"]:
        print(f"  防损: {t['big_numbers_converted']} 个超15位数字已转文本保存")
    for f in r["files"]:
        if f["status"] == "failed":
            print(f"  ✗ {f['path']}: {f['reason']}")
    for w in r["warnings"]:
        print(f"  ⚠ {w}")
    print(f"  日志: {Path(out).stem}_运行日志.log")
    return 0


if __name__ == "__main__":
    sys.exit(main())
