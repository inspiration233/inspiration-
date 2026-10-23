# -*- coding: utf-8 -*-
"""打包入口: PyInstaller 以此脚本构建独立可执行文件,
行为与 ``python -m engine.sidecar`` 完全一致(stdin/stdout JSON 行协议)。"""
from engine.sidecar import main

if __name__ == "__main__":
    main()
