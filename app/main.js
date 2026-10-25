// -*- coding: utf-8 -*-
// 工具箱壳 · Electron 主进程：无边框窗口 / Python sidecar 管理 / 对话框与路径 IPC
const { app, BrowserWindow, Menu, ipcMain, dialog, shell: shellOs, nativeTheme } = require('electron');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..');            // 仓库根（engine 所在）
// Python 解释器发现: 环境变量优先, 其次项目内 venv(mac/linux 在 bin/, Windows 在 Scripts/),
// 最后回退系统命令(python3 优先于 python, macOS 常无 python 别名)
function discoverPython() {
  if (process.env.TOOLBOX_PYTHON) return process.env.TOOLBOX_PYTHON;
  const candidates = process.platform === 'win32'
    ? [path.join(ROOT, '.venv', 'Scripts', 'python.exe'), path.join(ROOT, 'venv', 'Scripts', 'python.exe')]
    : [path.join(ROOT, '.venv', 'bin', 'python'), path.join(ROOT, 'venv', 'bin', 'python')];
  for (const candidate of candidates) {
    try { if (fs.existsSync(candidate)) return candidate; } catch { /* 忽略 */ }
  }
  return process.platform === 'win32' ? 'python' : 'python3';
}
const PYTHON = discoverPython();
const IS_SMOKE = process.argv.includes('--smoke');

// 已知问题缓解: Windows 下窗口遮挡计算会造成还原/缩放白闪
// 图形后端可配置(userData/angle.txt: 空=默认 | d3d9 | gl | d3d11)。
// 拉伸黑闪多为"当前通路×驱动"的兼容缺陷, 换一条通路即可绕开; 设置中心可切换。
const ANGLE_BACKENDS = new Set(['', 'd3d9', 'gl', 'd3d11']);
let angleBackend = '';
try {
  const saved = fs.readFileSync(path.join(app.getPath('userData'), 'angle.txt'), 'utf8').trim();
  if (ANGLE_BACKENDS.has(saved)) angleBackend = saved;
} catch { /* 未配置 */ }
// Windows 默认走 Direct3D 9(实测 D3D11 通路在部分 NVIDIA 驱动上拉伸卡顿/黑闪);
// 设置中心仍可改回。
if (!angleBackend && process.platform === 'win32') angleBackend = 'd3d9';
if (angleBackend) app.commandLine.appendSwitch('use-angle', angleBackend);

// CalculateNativeWinOcclusion: 缓解还原/缩放白闪(Windows 遮挡计算)。
// CanvasOopRasterization: 日志出现 "GPU state invalid after WaitForGetOffsetInRange"
// (GPU 命令缓冲重置→合成停帧→拉伸露黑), 与 NVIDIA 上的画布外光栅化已知问题吻合, 关闭。
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion,CanvasOopRasterization');
const themeBg = () => (nativeTheme.shouldUseDarkColors ? '#0f1420' : '#eef1f7');
const IS_WORKBENCH_TEST = process.argv.includes('--workbench-test') || process.argv.includes('--uitest');
// 真实会话探针: 复用用户真实 userData(不重定向), 只读诊断
const PROBE_REAL = !!process.env.TOOLBOX_PROBE_REAL;
if ((IS_WORKBENCH_TEST && !PROBE_REAL) || IS_SMOKE) {
  const os = require('os');
  app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'excel-workbench-ui-')));
}

let win = null;
let sidecar = null;
let sidecarBuf = '';
const pending = [];                                      // 请求队列（一次一个操作）
let waiting = null;                                      // 当前等待应答的 resolver

// ---------------------------------------------------------------- sidecar

function startSidecar() {
  // 打包版优先使用随应用分发的引擎二进制(resources/engine-sidecar);
  // 开发版回退 python -m engine.sidecar
  let command = PYTHON, args = ['-m', 'engine.sidecar'];
  if (app.isPackaged) {
    const name = process.platform === 'win32' ? 'engine-sidecar.exe' : 'engine-sidecar';
    const binary = path.join(process.resourcesPath, 'engine-sidecar', name);
    try { if (fs.existsSync(binary)) { command = binary; args = []; } } catch { /* 回退 python */ }
  }
  const processHandle = spawn(command, args, {
    cwd: ROOT,
    env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  sidecar = processHandle;
  sidecarBuf = '';
  sidecar.stdout.setEncoding('utf8');
  sidecar.stdout.on('data', (chunk) => {
    sidecarBuf += chunk;
    let idx;
    while ((idx = sidecarBuf.indexOf('\n')) >= 0) {
      const line = sidecarBuf.slice(0, idx).trim();
      sidecarBuf = sidecarBuf.slice(idx + 1);
      if (!line) continue;
      let msg;
      try { msg = JSON.parse(line); } catch { continue; }
      if (msg.evt === 'progress') {
        win?.webContents?.send('engine:progress', msg);
        continue;
      }
      if (msg.evt === 'report' || msg.evt === 'error') {
        const cb = waiting;
        waiting = null;
        drainNext();
        cb?.(msg);
      }
    }
  });
  sidecar.stderr.on('data', (d) => console.error('[sidecar]', String(d).trim()));
  const fail = (message) => {
    if (sidecar !== processHandle) return;
    sidecar = null;
    sidecarBuf = '';
    const active = waiting;
    waiting = null;
    active?.({ evt: 'error', code: 'engine_unavailable', msg: message });
    while (pending.length) pending.shift().resolve({ evt: 'error', code: 'engine_unavailable', msg: message });
  };
  sidecar.on('error', (error) => fail(`无法启动 Python 引擎：${error.message}。请检查 TOOLBOX_PYTHON 或 Python 依赖。`));
  sidecar.on('exit', (code) => fail(`Python 引擎已退出（${code ?? '已终止'}），请重试。`));
}

function engineSend(obj) {
  return new Promise((resolve) => {
    const launch = () => {
      if (!sidecar) startSidecar();
      waiting = resolve;
      sidecar.stdin.write(JSON.stringify(obj) + '\n', (error) => {
        if (error && waiting === resolve) {
          waiting = null;
          resolve({ evt: 'error', code: 'engine_write', msg: error.message });
          drainNext();
        }
      });
    };
    launch.resolve = resolve;
    if (waiting === null && !pending.length) {
      if (!sidecar) startSidecar();
      launch();
    } else {
      pending.push(launch);
    }
  });
}

function drainNext() {
  const next = pending.shift();
  if (next && waiting === null) next();
}

function killSidecar() {
  if (sidecar) {
    try { sidecar.kill(); } catch { /* ignore */ }
    sidecar = null;
  }
}

// ---------------------------------------------------------------- 窗口

function createWindow() {
  win = new BrowserWindow({
    width: 960,
    height: 640,
    minWidth: 800,
    minHeight: 560,
    // 原生窗口边框: 外层窗口(标题栏/最小化最大化关闭)跟随系统, 应用内保持自身风格
    backgroundColor: themeBg(),   // 初始底色跟随系统主题，避免缩放时闪异色
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: false,
      backgroundThrottling: false,   // 拉伸窗口期间保持重绘, 减少黑屏暴露
    },
  });
  win.loadFile(path.join(__dirname, 'shell', 'index.html'));
  win.once('ready-to-show', () => { if (!IS_WORKBENCH_TEST && !IS_SMOKE) win.show(); });

  // 精简中文应用菜单: 去掉 Electron 默认菜单(英文/外链)。
  // Windows/Linux: 菜单栏默认隐藏, 按 Alt 唤出(系统惯例), 复制粘贴等快捷键保持可用;
  // macOS: 菜单在屏幕顶部系统栏(不在窗口上), 放应用菜单/编辑/视图/窗口四组。
  const editMenu = { label: '编辑', submenu: [
    { role: 'undo', label: '撤销' }, { role: 'redo', label: '重做' }, { type: 'separator' },
    { role: 'cut', label: '剪切' }, { role: 'copy', label: '复制' }, { role: 'paste', label: '粘贴' },
    { role: 'selectAll', label: '全选' }] };
  const viewMenu = { label: '视图', submenu: [
    { role: 'reload', label: '刷新页面' }, { role: 'toggleDevTools', label: '开发者工具' }] };
  const windowMenu = { label: '窗口', submenu: [
    { role: 'minimize', label: '最小化' }, { role: 'close', label: '关闭' }] };
  const template = process.platform === 'darwin'
    ? [{ role: 'appMenu', label: '工具箱', submenu: [
        { role: 'about', label: '关于工具箱' }, { type: 'separator' },
        { role: 'hide', label: '隐藏工具箱' }, { role: 'hideOthers', label: '隐藏其他' },
        { type: 'separator' }, { role: 'quit', label: '退出工具箱' }] },
       editMenu, viewMenu, windowMenu]
    : [editMenu, viewMenu, windowMenu];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  if (process.platform !== 'darwin') {
    win.setMenuBarVisibility(false);
    win.setAutoHideMenuBar(true);
  }

  if (IS_WORKBENCH_TEST || PROBE_REAL) {
    win.webContents.once('did-finish-load', async () => {
      try {
        const result = await require('./workbench-smoke').run(win, ROOT);
        console.log('WORKBENCH-OK ' + JSON.stringify(result));
        app.exit(0);
      } catch (error) {
        console.error('WORKBENCH-FAIL ' + (error.stack || error));
        app.exit(1);
      }
    });
  }

  if (IS_SMOKE && !IS_WORKBENCH_TEST) {
    const errors = [];
    win.webContents.on('console-message', (_e, level, message) => {
      if (level >= 2 && !/GPU|command_buffer|EGL|ANGLE|GLES/i.test(message)) {
        errors.push(message);
      }
    });
    const scanPath = JSON.stringify(
      path.join(ROOT, '测试集', '场景E_表头智能识别', '报表_甲.xlsx'));
    win.webContents.on('did-finish-load', async () => {
      try {
        const headerRow = await win.webContents.executeJavaScript(`(async () => {
          document.querySelector('.tool-card')?.click();          // 进入工具页
          await new Promise((r) => setTimeout(r, 700));
          if (!window.toolbox) return 'no-toolbox';
          const rep = await window.toolbox.scan({ paths: [${scanPath}], options: {} });
          return rep.report?.header_detections?.[0]?.header_row;
        })()`);
        const ok = headerRow === 3;                              // 报表_甲 表头应判定第3行
        if (errors.length) {
          console.log('SMOKE-FAIL\n' + errors.join('\n'));
          app.exit(1);
        } else {
          console.log(ok ? 'SMOKE-OK' : `SMOKE-FAIL(scan=${headerRow})`);
          app.exit(ok ? 0 : 1);
        }
      } catch (e) {
        console.log('SMOKE-FAIL\n' + e);
        app.exit(1);
      }
    });
  }
}

// ---------------------------------------------------------------- IPC

// 路径展开：目录递归扫描 + 扩展名过滤 + 去重（按钮选择与拖拽统一走这里）
const TABLE_EXTS = new Set(['.xlsx', '.xls', '.csv']);
ipcMain.handle('paths:expand', async (_e, paths) => {
  const files = [];
  let skipped = 0;
  const seen = new Set();
  const seenDirs = new Set();
  const visit = (p) => {
    let st;
    try { st = fs.statSync(p); } catch { skipped += 1; return; }
    if (st.isDirectory()) {
      const real = fs.realpathSync(p).toLowerCase();
      if (seenDirs.has(real)) return;
      seenDirs.add(real);
      let names = [];
      try { names = fs.readdirSync(p); } catch { return; }
      names.forEach((n) => visit(path.join(p, n)));
    } else if (TABLE_EXTS.has(path.extname(p).toLowerCase()) && !path.basename(p).startsWith('~$')) {
      const key = path.resolve(p).toLowerCase();
      if (!seen.has(key)) { seen.add(key); files.push(p); }
    } else {
      skipped += 1;
    }
  };
  (paths || []).forEach(visit);
  return { files, skipped };
});

ipcMain.handle('dlg:files', async () => {
  const r = await dialog.showOpenDialog(win, {
    title: '选择表格文件',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: '表格文件', extensions: ['xlsx', 'xls', 'csv'] }],
  });
  return r.canceled ? [] : r.filePaths;
});

ipcMain.handle('dlg:folder', async () => {
  const r = await dialog.showOpenDialog(win, {
    title: '选择文件夹（递归扫描 xlsx/xls/csv）',
    properties: ['openDirectory'],
  });
  return r.canceled ? [] : r.filePaths;
});

ipcMain.handle('dlg:save', async (_e, defaultName) => {
  const r = await dialog.showSaveDialog(win, {
    title: '结果保存为',
    defaultPath: defaultName,
    filters: [{ name: 'Excel 工作簿', extensions: ['xlsx'] }],
  });
  return r.canceled ? null : r.filePath;
});

ipcMain.handle('dlg:output-folder', async () => {
  const r = await dialog.showOpenDialog(win, {
    title: '选择拆分结果保存位置', properties: ['openDirectory', 'createDirectory'],
  });
  return r.canceled ? null : r.filePaths[0];
});
ipcMain.handle('engine:scan', (_e, payload) => engineSend({ tool: 'excel_merge', cmd: 'scan', ...payload }));
ipcMain.handle('engine:merge', (_e, payload) => engineSend({ ...payload, tool: 'excel_merge', cmd: 'merge', workbench: true }));
for (const cmd of ['preview', 'merge_preview', 'split_preview', 'split', 'reconcile_preview', 'reconcile']) {
  ipcMain.handle(`engine:${cmd}`, (_e, payload) => engineSend({ ...payload, tool: 'excel_merge', cmd }));
}
ipcMain.on('engine:cancel', () => {
  if (sidecar) sidecar.stdin.write('{"cmd":"cancel"}\n');
});
ipcMain.handle('open:show', (_e, p) => shellOs.showItemInFolder(p));
// 主题切换时同步窗口原生底色（消除浅深切换/缩放时的底色闪烁）
ipcMain.on('win:set-bg', (_e, color) => {
  if (win && /^#[0-9a-fA-F]{6}$/.test(color)) win.setBackgroundColor(color);
});

ipcMain.on('win:ctl', (_e, action) => {
  if (!win) return;
  if (action === 'min') win.minimize();
  else if (action === 'max') win.isMaximized() ? win.unmaximize() : win.maximize();
  else if (action === 'close') win.close();
});

ipcMain.handle('gpu:set-angle', (_e, value) => {
  if (!ANGLE_BACKENDS.has(value)) throw Error('未知的图形后端：' + value);
  fs.writeFileSync(path.join(app.getPath('userData'), 'angle.txt'), value, 'utf8');
  app.relaunch();
  app.exit(0);
  return 'relaunching';
});

ipcMain.handle('app:info', () => {
  let gpu = 'unknown';
  try { gpu = app.getGPUFeatureStatus().gpu_compositing; } catch { /* 忽略 */ }
  return {
    appVersion: app.getVersion(),
    platform: process.platform,
    electron: process.versions.electron,
    python: PYTHON,
    angleBackend,
    gpuCompositing: gpu,        // 'disabled_software' = 软件渲染, 需自动材质降级
  };
});

// ---------------------------------------------------------------- 生命周期

app.whenReady().then(createWindow);
app.on('window-all-closed', () => { killSidecar(); app.quit(); });
app.on('before-quit', killSidecar);
