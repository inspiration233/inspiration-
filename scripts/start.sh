#!/usr/bin/env bash
# macOS / Linux 启动脚本(与 start.ps1 等价): 检查 Python 依赖后启动 Electron
set -e
cd "$(dirname "$0")/.."
APP=app
if [ ! -x "$APP/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron" ] && [ ! -x "$APP/node_modules/.bin/electron" ]; then
  echo "缺少 Electron。请在 app 目录执行 npm ci 后重试。" >&2
  exit 1
fi
PY=python3
if [ -n "$TOOLBOX_PYTHON" ]; then PY="$TOOLBOX_PYTHON"
elif [ -x ".venv/bin/python" ]; then PY=".venv/bin/python"
elif [ -x "venv/bin/python" ]; then PY="venv/bin/python"; fi
"$PY" -c "import openpyxl, xlrd; assert tuple(map(int, openpyxl.__version__.split('.')[:2])) >= (3, 1)" 2>/dev/null \
  || { echo "Python 依赖不完整: 需要 openpyxl>=3.1 与 xlrd。请执行: $PY -m pip install -r requirements.txt" >&2; exit 1; }
cd "$APP" && exec ./node_modules/.bin/electron .
