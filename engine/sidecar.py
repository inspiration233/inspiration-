# -*- coding: utf-8 -*-
"""Python 引擎 sidecar：JSON 行协议（stdin/stdout）。

用法：python -m engine.sidecar
消息格式（每行一条 JSON）：
  → {"tool":"excel_merge","cmd":"scan","paths":[...],"options":{...}}
  → {"tool":"excel_merge","cmd":"merge","paths":[...],"out":"...","options":{...}}
  → {"cmd":"cancel"}
  ← {"evt":"progress","file":1,"total":6,"current":"xx.xlsx","rows":1234}
  ← {"evt":"log","level":"info","msg":"..."}
  ← {"evt":"report",...} / {"evt":"error","msg":"..."}
"""
import json
import sys
import threading

from .tools.excel_merge import __version__
from .tools.excel_merge.merge import Cancelled, Options, merge, scan

_lock = threading.Lock()
_cancel = threading.Event()
_job_lock = threading.Lock()
_current_job = None


def emit(obj):
    with _lock:
        sys.stdout.write(json.dumps(obj, ensure_ascii=False, default=str) + "\n")
        sys.stdout.flush()


def _options(payload):
    allowed = {f: getattr(Options(), f) for f in Options().__dataclass_fields__}
    kw = {}
    for k, v in (payload.get("options") or {}).items():
        if k in allowed:
            if isinstance(allowed[k], tuple) and isinstance(v, list):
                v = tuple(v)
            kw[k] = v
    return Options(**kw)


def _run_merge(msg):
    global _current_job
    _current_job = threading.current_thread()
    paths = [str(p) for p in msg.get("paths", [])]
    out = msg.get("out")
    _cancel.clear()

    def progress(fi, total, current, rows):
        emit({"evt": "progress", "file": fi, "total": total,
              "current": current, "rows": rows})
        return False

    try:
        with _job_lock:
            report = merge(paths, out, _options(msg), progress=progress,
                           cancel=_cancel.is_set)
        emit({"evt": "report", "report": report})
    except Cancelled:
        emit({"evt": "error", "code": "cancelled", "msg": "已取消"})
    except Exception as e:
        emit({"evt": "error", "code": type(e).__name__, "msg": str(e)})


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            emit({"evt": "error", "code": "bad_json", "msg": line[:200]})
            continue
        cmd = msg.get("cmd")
        if cmd == "quit":
            break
        if cmd == "ping":
            emit({"evt": "pong", "version": __version__})
        elif cmd == "cancel":
            _cancel.set()
        elif cmd == "scan":
            try:
                emit({"evt": "report", "report": scan([str(p) for p in msg.get("paths", [])],
                                                      _options(msg))})
            except Exception as e:
                emit({"evt": "error", "code": type(e).__name__, "msg": str(e)})
        elif cmd == "merge":
            threading.Thread(target=_run_merge, args=(msg,), daemon=True).start()
        else:
            emit({"evt": "error", "code": "unknown_cmd", "msg": str(cmd)})


if __name__ == "__main__":
    main()
    if _current_job and _current_job.is_alive():
        _current_job.join(timeout=600)
