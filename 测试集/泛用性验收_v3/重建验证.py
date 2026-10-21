# -*- coding: utf-8 -*-
"""在空临时目录重建并自检，比较文件清单、JSON及XLSX内部数据/规则。"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent
STEPS = ["生成验收集.py", "补充案例.py", "修订.py", "收尾修订.py"]
OBSOLETE = {"W3/W3_预期_月度汇总.xlsx"}


def artifacts(root):
    return {p.relative_to(root).as_posix(): p
            for d in root.iterdir() if d.is_dir() and (d / "case.json").exists()
            for p in d.iterdir() if p.suffix == ".xlsx" or p.name == "case.json"}


def fingerprint(path):
    """只忽略ZIP封装和文档创建/修改时间；公式、样式、关系、规则均参与比较。"""
    if path.suffix == ".json":
        return json.dumps(json.loads(path.read_text(encoding="utf-8")),
                          sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    result = {}
    with ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            raw = archive.read(name)
            if name.endswith((".xml", ".rels", ".vml")):
                if name == "docProps/core.xml":
                    xml = ET.fromstring(raw)
                    for node in xml.iter():
                        if node.tag in {"{http://purl.org/dc/terms/}created",
                                        "{http://purl.org/dc/terms/}modified"}:
                            node.text = "IGNORED_TIMESTAMP"
                    raw = ET.tostring(xml, encoding="utf-8")
                raw = ET.canonicalize(raw).encode("utf-8")
            result[name] = hashlib.sha256(raw).hexdigest()
    return result


def compare(expected_root, actual_root):
    expected, actual = artifacts(expected_root), artifacts(actual_root)
    differences = []
    for name in sorted(expected.keys() | actual.keys()):
        if name not in actual:
            differences.append(f"缺少文件: {name}")
        elif name not in expected:
            differences.append(f"多出文件: {name}")
        else:
            a, b = fingerprint(expected[name]), fingerprint(actual[name])
            if a != b:
                parts = sorted(k for k in a.keys() | b.keys() if a.get(k) != b.get(k)) if isinstance(a, dict) else ["JSON"]
                differences.append(f"内容不同: {name} ({', '.join(parts)})")
    return differences


def build(destination):
    for script in ROOT.glob("*.py"):
        shutil.copy2(script, destination / script.name)
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONDONTWRITEBYTECODE": "1"}
    for name in STEPS + ["自检.py"]:
        command = [sys.executable, "-B", str(destination / name)]
        if name == "自检.py":
            command.append("--quiet")
        process = subprocess.run(command, cwd=destination, env=env,
                                 capture_output=True, text=True, encoding="utf-8")
        if process.returncode:
            print(process.stdout)
            print(process.stderr)
            raise RuntimeError(f"{name} 失败，退出码 {process.returncode}")
        print(f"完成: {name}")
        if name == "自检.py":
            print(process.stdout.strip())


def update(destination):
    """仅更新本测试集的生成文件，不删除案例目录或用户其他文件。"""
    generated, current = artifacts(destination), artifacts(ROOT)
    extras = current.keys() - generated.keys()
    unexpected = extras - OBSOLETE
    if unexpected:
        raise RuntimeError(f"发现未登记文件，已保留；请先核对：{sorted(unexpected)}")
    for name in sorted(extras):
        target = current[name].resolve()
        if not target.is_relative_to(ROOT) or target.suffix != ".xlsx":
            raise RuntimeError(f"拒绝移除非生成工作簿: {target}")
        target.unlink()
        print(f"移除过期工作簿: {name}")
    for name, source in generated.items():
        target = (ROOT / name).resolve()
        if not target.is_relative_to(ROOT):
            raise RuntimeError(f"目标超出测试集目录: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        # 未变的文件不重写，避免仅时间戳不同导致无意义的二进制差异。
        if not target.exists() or fingerprint(source) != fingerprint(target):
            shutil.copy2(source, target)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--update", action="store_true",
                        help="重建、自检成功后将生成文件同步到当前测试集")
    args = parser.parse_args()
    temp_parent = Path(tempfile.gettempdir()).resolve()
    destination = Path(tempfile.mkdtemp(prefix="excel-v3-rebuild-", dir=temp_parent)).resolve()
    try:
        build(destination)
        if args.update:
            update(destination)
        differences = compare(destination, ROOT)
        for message in differences:
            print(message)
        count = len(artifacts(destination))
        print(f"重建对照: {count} 个生成文件，{len(differences)} 项差异"
              "（忽略文档时间戳和ZIP封装）")
        return 1 if differences else 0
    except Exception as exc:
        print(f"重建失败: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    finally:
        # 校验绝对路径后，仅清理本次新建的临时目录。
        if destination.parent != temp_parent or not destination.name.startswith("excel-v3-rebuild-"):
            raise RuntimeError(f"拒绝清理意外路径: {destination}")
        shutil.rmtree(destination)


if __name__ == "__main__":
    raise SystemExit(main())
