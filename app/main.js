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
    const pDir = JSON.stringify(path.join(ROOT, '测试集', '场景A_多文件纵向合并'));
    const out = JSON.stringify(path.join(ROOT, '_uitest_out.xlsx'));
    const shot = (name) => win.webContents.capturePage().then((img) => {
      fs.writeFileSync(path.join(ROOT, name), img.toPNG());
    });
    win.webContents.on('did-finish-load', async () => {
      try {
        const r1 = await win.webContents.executeJavaScript(`(async () => {
          const sleep = (ms) => new Promise((res) => setTimeout(res, ms));
          document.querySelector('.tool-card')?.click();
          await sleep(700);
          const T = window.__emTest;
          if (!T) return 'no-hook';
          // 回归防线: 状态面板初始必须真的隐藏(渲染层面)
          const dn = getComputedStyle(document.querySelector('#done-pane')).display;
          if (dn !== 'none') return 'done-pane-visible(display=' + dn + ')';
          const pp = getComputedStyle(document.querySelector('#progress-pane')).display;
          if (pp !== 'none') return 'progress-pane-visible(display=' + pp + ')';
          T.addPaths({ files: [${p1}, ${p2}] });
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
          // 文件夹展开（"选择文件夹"按钮同路径）：场景A 目录应扫出 6 个表格文件
          const exp = await window.toolbox.expandPaths([${pDir}]);
          if (exp.files.length !== 6) return 'folder-expand=' + exp.files.length;
          // 抽屉展开：无遮挡（action-bar 不得压住抽屉内容）、无截断
          document.querySelector('#settings-head').click();
          await sleep(500);
          const sBody = document.querySelector('#settings-body');
          const aBar = document.querySelector('.action-bar');
          if (aBar.getBoundingClientRect().top < sBody.getBoundingClientRect().bottom - 2) return 'drawer-overlap';
          if (sBody.clientHeight < sBody.scrollHeight - 2) return 'drawer-clipped';
          return 'PHASE1-OK';  // 抽屉保持展开，供截屏核验
        })()`);
        if (r1 !== 'PHASE1-OK') { console.log('UITEST-FAIL ' + r1); app.exit(1); return; }
        await shot('_uitest_1_filelist.png');
        const r2 = await win.webContents.executeJavaScript(`(async () => {
          const sleep = (ms) => new Promise((res) => setTimeout(res, ms));
          document.querySelector('#settings-head')?.click();     // 收回抽屉
          await sleep(300);
          const ok = await window.__emTest.startMerge(${out});   // 真实合并
          if (!ok) return 'merge-failed';
          for (let i = 0; i < 60; i++) {
            await sleep(500);
            if (!document.querySelector('#done-pane').classList.contains('hidden')) break;
          }
          if (document.querySelector('#done-pane').classList.contains('hidden')) return 'no-done-pane';
          return 'DONE:' + document.querySelector('#done-stats').textContent;
        })()`);
        await shot('_uitest_2_done.png');
        const r = r2;
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

// 路径展开：目录递归扫描 + 扩展名过滤 + 去重（按钮选择与拖拽统一走这里）
const TABLE_EXTS = new Set(['.xlsx', '.xls', '.csv']);
ipcMain.handle('paths:expand', async (_e, paths) => {
  const files = [];
  let skipped = 0;
  const seen = new Set();
  const visit = (p) => {
    let st;
    try { st = fs.statSync(p); } catch { skipped += 1; return; }
    if (st.isDirectory()) {
      let names = [];
      try { names = fs.readdirSync(p); } catch { return; }
      names.forEach((n) => visit(path.join(p, n)));
    } else if (TABLE_EXTS.has(path.extname(p).toLowerCase())) {
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
