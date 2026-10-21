// -*- coding: utf-8 -*-
// 工具箱壳 · Electron 主进程：无边框窗口 / Python sidecar 管理 / 对话框与路径 IPC
const { app, BrowserWindow, ipcMain, dialog, shell: shellOs, nativeTheme } = require('electron');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..');            // 仓库根（engine 所在）
const PYTHON = process.env.TOOLBOX_PYTHON || 'python';
const IS_SMOKE = process.argv.includes('--smoke');

// 已知问题缓解: Windows 下窗口遮挡计算会造成还原/缩放白闪
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
const themeBg = () => (nativeTheme.shouldUseDarkColors ? '#0f1420' : '#eef1f7');
const IS_WORKBENCH_TEST = process.argv.includes('--workbench-test') || process.argv.includes('--uitest');
if (IS_WORKBENCH_TEST || IS_SMOKE) {
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
  const processHandle = spawn(PYTHON, ['-m', 'engine.sidecar'], {
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
    // 'hidden' 而非 frame:false —— 视觉上同样无原生标题栏,
    // 但缩放仍走系统原生路径, 消除 frame:false 重实现缩放时的黑闪(Windows DWM)
    titleBarStyle: 'hidden',
    backgroundColor: themeBg(),   // 初始底色跟随系统主题，避免缩放时闪异色
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  win.loadFile(path.join(__dirname, 'shell', 'index.html'));
  win.once('ready-to-show', () => { if (!IS_WORKBENCH_TEST && !IS_SMOKE) win.show(); });

  if (IS_WORKBENCH_TEST) {
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

ipcMain.handle('app:info', () => {
  let gpu = 'unknown';
  try { gpu = app.getGPUFeatureStatus().gpu_compositing; } catch { /* 忽略 */ }
  return {
    appVersion: app.getVersion(),
    platform: process.platform,
    electron: process.versions.electron,
    python: PYTHON,
    gpuCompositing: gpu,        // 'disabled_software' = 软件渲染, 需自动材质降级
  };
});

// ---------------------------------------------------------------- 生命周期

app.whenReady().then(createWindow);
app.on('window-all-closed', () => { killSidecar(); app.quit(); });
app.on('before-quit', killSidecar);
