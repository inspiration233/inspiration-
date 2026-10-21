"""UTF-8 JSON-lines bridge. Jobs run serially; cancel remains responsive."""
import hashlib
import json
from pathlib import Path
import queue
import sys
import threading

from .tools.excel_merge import __version__
from .tools.excel_merge.merge import Cancelled, Options, merge, scan

_lock = threading.Lock()
_jobs = queue.Queue()
_event_lock = threading.Lock()
_events = []


def emit(obj):
    with _lock:
        sys.stdout.write(json.dumps(obj, ensure_ascii=False, default=str) + "\n")
        sys.stdout.flush()


def _options(payload):
    allowed = Options().__dataclass_fields__
    kw = {k: v for k, v in (payload.get("options") or {}).items() if k in allowed}
    if "summary_keywords" in kw:
        kw["summary_keywords"] = tuple(kw["summary_keywords"])
    return Options(**kw)


def source_fingerprint(payload):
    paths = list(payload.get("paths") or [])
    if payload.get("path"):
        paths.append(payload["path"])
    for item in payload.get("feedback", []) + payload.get("selections", []):
        if item.get("path"):
            paths.append(item["path"])
    baseline = payload.get("baseline")
    if baseline and baseline.get("path"):
        paths.append(baseline["path"])
    digest = hashlib.sha256()
    for path in sorted({str(Path(p).resolve()) for p in paths}):
        digest.update(path.encode("utf-8"))
        with open(path, "rb") as stream:
            for data in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(data)
    return digest.hexdigest()


def _dispatch(msg, event):
    cmd = msg["cmd"]

    def progress(fi, total, current, rows):
        emit({"evt": "progress", "file": fi, "total": total, "current": current, "rows": rows})
        return False

    if event.is_set():
        raise Cancelled()
    if cmd == "scan":
        return scan(msg.get("paths", []), _options(msg))
    if cmd == "preview":
        from .tools.excel_merge.preview import preview
        return preview(msg)
    if cmd == "merge" and not msg.get("workbench"):
        return merge(msg.get("paths", []), msg.get("out"), _options(msg), progress, event.is_set)
    from .tools.excel_merge.workbench import merge_preview, merge_export
    from .tools.excel_merge.native_split import split_preview, split_export
    from .tools.excel_merge.workflow import workflow_preview, workflow_export
    handlers = {"merge_preview": merge_preview, "merge": merge_export,
                "split_preview": split_preview, "split": split_export,
                "reconcile_preview": workflow_preview, "reconcile": workflow_export}
    if cmd not in handlers:
        raise ValueError(f"未知操作：{cmd}")
    before = source_fingerprint(msg)
    if msg.get("preview_fingerprint") and msg["preview_fingerprint"] != before:
        raise ValueError("源文件在预览后发生变化，请重新预览后导出")
    if cmd in ("split_preview", "reconcile_preview"):
        report = handlers[cmd](msg)
    else:
        report = handlers[cmd](msg, cancel=event.is_set, progress=progress)
    if cmd.endswith("_preview"):
        if event.is_set():
            raise Cancelled()
        if source_fingerprint(msg) != before:
            raise ValueError("源文件在预览过程中发生变化，请重试")
        report["source_fingerprint"] = before
    return report


def _worker():
    while True:
        job = _jobs.get()
        if job is None:
            return
        msg, event = job
        try:
            report = _dispatch(msg, event)
            emit({"evt": "report", "report": report})
        except Cancelled:
            emit({"evt": "error", "code": "cancelled", "msg": "已取消"})
        except Exception as exc:
            result = {"evt": "error", "code": type(exc).__name__, "msg": str(exc)}
            if hasattr(exc, "preview"):
                result["report"] = exc.preview
            emit(result)
        finally:
            with _event_lock:
                if event in _events:
                    _events.remove(event)


def main():
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    worker = threading.Thread(target=_worker, daemon=True)
    worker.start()
    for line in sys.stdin:
        if not line.strip():
            continue
        try:
            msg = json.loads(line)
            if not isinstance(msg, dict):
                raise ValueError("消息必须是对象")
        except (ValueError, json.JSONDecodeError):
            emit({"evt": "error", "code": "bad_json", "msg": "无效 JSON 消息"})
            continue
        cmd = msg.get("cmd")
        if cmd == "quit":
            break
        if cmd == "ping":
            emit({"evt": "pong", "version": __version__})
        elif cmd == "cancel":
            with _event_lock:
                if _events:
                    _events[0].set()
        elif cmd in ("scan", "preview", "merge_preview", "merge", "split_preview", "split",
                     "reconcile_preview", "reconcile"):
            event = threading.Event()
            with _event_lock:
                _events.append(event)
            _jobs.put((msg, event))
        else:
            emit({"evt": "error", "code": "unknown_cmd", "msg": str(cmd)})
    _jobs.put(None)
    worker.join(timeout=600)


if __name__ == "__main__":
    main()
