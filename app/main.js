// -*- coding: utf-8 -*-
// 工具箱壳 · Electron 主进程：无边框窗口 / Python sidecar 管理 / 对话框与路径 IPC
const { app, BrowserWindow, ipcMain, dialog, shell: shellOs } = require('electron');
const { spawn } = require('child_process');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');            // 仓库根（engine 所在）
const PYTHON = process.env.TOOLBOX_PYTHON || 'python';
const IS_SMOKE = process.argv.includes('--smoke');

let win = null;
let sidecar = null;
let sidecarBuf = '';
const pending = [];                                      // 请求队列（一次一个操作）
let waiting = null;                                      // 当前等待应答的 resolver

// ---------------------------------------------------------------- sidecar

function startSidecar() {
  sidecar = spawn(PYTHON, ['-m', 'engine.sidecar'], {
    cwd: ROOT,
    env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
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
  sidecar.on('exit', (code) => { sidecar = null; });
}

function engineSend(obj) {
  return new Promise((resolve) => {
    const launch = () => {
      waiting = resolve;
      sidecar.stdin.write(JSON.stringify(obj) + '\n');
    };
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
    frame: false,
    backgroundColor: '#0f1420',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  win.loadFile(path.join(__dirname, 'shell', 'index.html'));
  win.once('ready-to-show', () => win.show());
  if (IS_SMOKE) {
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

ipcMain.handle('dlg:files', async () => {
  const r = await dialog.showOpenDialog(win, {
    title: '选择要合并的文件',
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
    title: '合并结果保存为',
    defaultPath: defaultName,
    filters: [{ name: 'Excel 工作簿', extensions: ['xlsx'] }],
  });
  return r.canceled ? null : r.filePath;
});

ipcMain.handle('engine:scan', (_e, payload) => engineSend({ tool: 'excel_merge', cmd: 'scan', ...payload }));
ipcMain.handle('engine:merge', (_e, payload) => engineSend({ tool: 'excel_merge', cmd: 'merge', ...payload }));
ipcMain.on('engine:cancel', () => {
  if (sidecar) sidecar.stdin.write('{"cmd":"cancel"}\n');
});
ipcMain.handle('open:show', (_e, p) => shellOs.showItemInFolder(p));
ipcMain.handle('app:info', () => ({
  appVersion: app.getVersion(),
  platform: process.platform,
  electron: process.versions.electron,
  python: PYTHON,
}));

// ---------------------------------------------------------------- 生命周期

app.whenReady().then(createWindow);
app.on('window-all-closed', () => { killSidecar(); app.quit(); });
app.on('before-quit', killSidecar);
