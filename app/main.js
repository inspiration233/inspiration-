// -*- coding: utf-8 -*-
// 工具箱壳 · Electron 主进程：无边框窗口 / Python sidecar 管理 / 对话框与路径 IPC
const { app, BrowserWindow, ipcMain, dialog, shell: shellOs } = require('electron');
const { spawn } = require('child_process');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');            // 仓库根（engine 所在）
const PYTHON = process.env.TOOLBOX_PYTHON || 'python';
const IS_SMOKE = process.argv.includes('--smoke');
const IS_UITEST = process.argv.includes('--uitest');

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

  if (IS_UITEST) {
    const errors = [];
    win.webContents.on('console-message', (_e, level, message) => {
      if (level >= 2 && !/GPU|command_buffer|EGL|ANGLE|GLES/i.test(message)) {
        errors.push(message);
      }
    });
    const p1 = JSON.stringify(path.join(ROOT, '测试集', '场景E_表头智能识别', '报表_甲.xlsx'));
    const p2 = JSON.stringify(path.join(ROOT, '测试集', '场景E_表头智能识别', '报表_丙.csv'));
    const out = JSON.stringify(path.join(ROOT, '_uitest_out.xlsx'));
    win.webContents.on('did-finish-load', async () => {
      try {
        const r = await win.webContents.executeJavaScript(`(async () => {
          const sleep = (ms) => new Promise((res) => setTimeout(res, ms));
          document.querySelector('.tool-card')?.click();
          await sleep(700);
          const T = window.__emTest;
          if (!T) return 'no-hook';
          T.addPaths([${p1}, ${p2}]);
          for (let i = 0; i < 20; i++) {           // 等体检回填
            await sleep(300);
            const metas = [...document.querySelectorAll('.file-row .meta')];
            if (metas.length === 2 && metas.every((m) => !m.textContent.includes('识别中'))) break;
          }
          const rows = document.querySelectorAll('.file-row');
          if (rows.length !== 2) return 'rows=' + rows.length;
          const okMeta = document.querySelector('.file-row .meta.ok');
          if (!okMeta) return 'scan-not-applied';
          if (!okMeta.textContent.includes('表头第3行')) return 'meta=' + okMeta.textContent;
          const ok = await T.startMerge(${out});    // 真实合并
          if (!ok) return 'merge-failed';
          for (let i = 0; i < 60; i++) {
            await sleep(500);
            if (!document.querySelector('#done-pane').classList.contains('hidden')) break;
          }
          if (document.querySelector('#done-pane').classList.contains('hidden')) return 'no-done-pane';
          return 'DONE:' + document.querySelector('#done-stats').textContent;
        })()`);
        const pass = typeof r === 'string' && r.startsWith('DONE:') && r.includes('写入');
        if (errors.length) {
          console.log('UITEST-FAIL(控制台错误): ' + errors.join(' | '));
          app.exit(1);
        } else {
          console.log(pass ? 'UITEST-OK ' + r.slice(5) : 'UITEST-FAIL ' + r);
          app.exit(pass ? 0 : 1);
        }
      } catch (e) {
        console.log('UITEST-FAIL: ' + e);
        app.exit(1);
      }
    });
  }
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
ipcMain.on('win:ctl', (_e, action) => {
  if (!win) return;
  if (action === 'min') win.minimize();
  else if (action === 'max') win.isMaximized() ? win.unmaximize() : win.maximize();
  else if (action === 'close') win.close();
});

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
