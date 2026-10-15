# 团队工具箱

本地桌面工具箱，Windows / macOS 双平台，**完全离线运行，文件不出本机**。

当前工具：**Excel 合并**——四种合并模式、表头智能识别（表头行自动探测 / 多行表头 / 合计行剔除）、流式处理百万行级数据。

## 文档

- [工具箱平台说明.md](工具箱平台说明.md) — 壳架构、工具接入规范（manifest）、发布模式（CI + 内部渠道）
- [需求说明.md](需求说明.md) — Excel 合并工具需求定稿（V1.7）
- [测试集/README.md](测试集/README.md) — 33 个测试文件、脏点设计与预期行为（验收标准）

## 测试集

小规模测试夹具已入库；大规模性能文件不入库，按需生成：

```bash
python 测试集/生成测试集.py --rows 100000   # 生成 10 万行 ×3 的大规模文件
python 测试集/生成测试集.py --skip-large    # 只重建小规模夹具
```

## 开发

- 引擎：Python 3（openpyxl / xlrd，流式读写），sidecar 形态由壳调用
- 界面：Electron + HTML/CSS/JS
- 打包：PyInstaller + electron-builder，GitHub Actions 双平台产出，内部渠道分发

> 仓库结构随开发推进更新（M1 引擎 → M2 壳与界面 → M3 CI）。
