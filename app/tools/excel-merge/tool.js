// Excel 合并 · 工具页逻辑（文件列表/体检/设置/合并/结算）
/* global toolbox, tbToast */
window.toolMount = (root, ctx) => {
  const $ = (id) => root.querySelector('#' + id);
  const state = {
    files: [],            // {path, name, status: ok|bad|pending, detections:[], reason}
    merging: false,
    offProgress: null,
  };

  // ---------------- 设置（记忆上次） ----------------
  const DEF = {
    mode: 'single', align: 'loose', style: 'unified', templateIndex: 0,
    headerAuto: '0', headerRow: 1, removeSummary: true, keywords: ['合计', '小计', '总计', '汇总', '累计'],
    source: false, hidden: false, title: false, outDir: '',
    stripSerial: true, renumberSerial: true,
  };
  let cfg = { ...DEF, ...JSON.parse(localStorage.getItem('excelMergeCfg') || '{}') };
  const saveCfg = () => localStorage.setItem('excelMergeCfg', JSON.stringify(cfg));

  // ---------------- 拖放（文件+文件夹递归） ----------------
  const dropzone = $('dropzone');
  const collectEntries = async (items) => {
    const files = [];
    const walk = (entry) => new Promise((res) => {
      if (entry.isFile) {
        entry.file((f) => { files.push(f); res(); }, res);
      } else if (entry.isDirectory) {
        const reader = entry.createReader();
        const readAll = () => reader.readEntries(async (ents) => {
          if (!ents.length) return res();
          await Promise.all(ents.map(walk));
          readAll();                       // 目录可能分批返回
        }, res);
        readAll();
      } else res();
    });
    for (const it of items) {
      const entry = it.webkitGetAsEntry();
      if (entry) await walk(entry);
    }
    return files;
  };

  // 统一入口：主进程做目录递归/扩展名过滤/去重（按钮与拖拽同一条路，文件夹按钮由此修复）
  const addFromPaths = async (rawPaths) => {
    if (!rawPaths || !rawPaths.length) return;
    try {
      addPaths(await toolbox.expandPaths(rawPaths));
    } catch (e) {
      tbToast(`添加失败：${e}`, 'error');
    }
  };

  root.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
  root.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  root.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    let paths = [...e.dataTransfer.files].map((f) => toolbox.pathForFile(f));
    if (!paths.length) {                      // 少数场景 files 为空但有 entries
      const items = [...(e.dataTransfer.items || [])];
      if (items.length && items[0].webkitGetAsEntry) {
        const files = await collectEntries(items);
        paths = files.map((f) => toolbox.pathForFile(f));
      }
    }
    addFromPaths(paths);
  });

  $('btn-add-files').onclick = async () => addFromPaths(await toolbox.selectFiles());
  $('btn-add-folder').onclick = async () => addFromPaths(await toolbox.selectFolder());
  $('btn-add-more').onclick = () => $('btn-add-files').click();

  // ---------------- 文件列表 ----------------
  const summarize = (dets) => dets
    .map((d) => (dets.length > 1 ? `${d.sheet}: ` : '') + `表头第${d.header_row}行·${d.columns.length}列`)
    .join('、');

  function addPaths({ files, skipped = 0 }) {
    let added = 0;
    let dup = 0;
    for (const p of files) {
      if (state.files.some((f) => f.path === p)) { dup++; continue; }
      state.files.push({ path: p, name: p.split(/[\\/]/).pop(), status: 'pending' });
      added++;
    }
    if (added) { refreshList(); scheduleScan(); }
    const skippedTotal = skipped + dup;
    if (added || skippedTotal) {
      tbToast(`已添加 ${added} 个文件` +
        (skippedTotal ? `，跳过 ${skippedTotal} 个（不支持或重复）` : ''));
    }
  }

  function refreshList() {
    const wrap = $('file-list-wrap');
    wrap.classList.toggle('hidden', state.files.length === 0);
    $('dz-idle').classList.toggle('hidden', state.files.length > 0);
    $('dz-thin').classList.toggle('hidden', state.files.length === 0);
    $('file-count').textContent = state.files.length;
    $('btn-merge').disabled = state.files.every((f) => f.status === 'bad');

    const rows = $('file-rows');
    rows.innerHTML = '';
    state.files.forEach((f, i) => {
      const row = document.createElement('div');
      row.className = 'file-row row-item';
      row.style.setProperty('--i', i);
      row.draggable = true;
      const meta = f.status === 'bad' ? `<span class="meta bad">✗ ${f.reason || '无法读取'}</span>`
        : f.detections ? `<span class="meta ok">✓ ${summarize(f.detections)}</span>`
        : '<span class="meta">识别中…</span>';
      const warns = (f.detections || []).some((d) => d.warnings?.length || d.low_confidence);
      row.innerHTML = `
        <span class="drag-handle">⠿</span>
        <span class="fname" title="${f.path}">${f.name}</span>
        ${warns ? '<span class="warn-dot" title="有警告，点击行展开">⚠</span>' : ''}
        ${meta}
        <span class="rm" title="移除">✕</span>`;
      row.querySelector('.rm').onclick = (e) => { e.stopPropagation(); removeFile(i); };
      row.onclick = () => toggleDetail(row, f);
      row.ondragstart = (e) => { e.dataTransfer.setData('text/plain', String(i)); row.classList.add('dragging'); };
      row.ondragend = () => row.classList.remove('dragging');
      row.ondragover = (e) => e.preventDefault();
      row.ondrop = (e) => {
        e.preventDefault(); e.stopPropagation();
        const from = Number(e.dataTransfer.getData('text/plain'));
        if (Number.isInteger(from) && from !== i) {
          const [m] = state.files.splice(from, 1);
          state.files.splice(i, 0, m);
          refreshList();
        }
      };
      rows.appendChild(row);
    });
    refreshTemplateSelect();
  }

  function toggleDetail(row, f) {
    const next = row.nextElementSibling;
    if (next?.classList.contains('detail')) { next.remove(); return; }
    root.querySelectorAll('.detail').forEach((d) => d.remove());
    if (!f.detections) return;
    const div = document.createElement('div');
    div.className = 'detail';
    div.innerHTML = f.detections.map((d) => `
      <div class="line"><b>${d.display}</b></div>
      <div class="line">排列: ${d.orientation === 'horizontal' ? '横向表（已自动转正）' : '纵向'} · 表头第 ${d.header_row} 行 · ${d.height} 行高 · ${d.columns.length} 列${d.title ? ` · 标题(不入表): ${d.title}` : ''}</div>
      ${d.low_confidence ? '<div class="line w">⚠ 两个候选表头行得分接近，请核对</div>' : ''}
      ${(d.warnings || []).map((w) => `<div class="line w">⚠ ${w}</div>`).join('')}
      <div class="line">列: ${d.columns.join(' | ')}</div>`).join('<hr style="border:none;border-top:1px solid var(--hairline-2);margin:6px 0">');
    row.after(div);
  }

  function removeFile(i) {
    state.files.splice(i, 1);
    refreshList();
  }
  $('btn-clear').onclick = () => { state.files = []; refreshList(); };

  // ---------------- 体检 ----------------
  let scanTimer = null;
  function scheduleScan() {
    clearTimeout(scanTimer);
    scanTimer = setTimeout(runScan, 250);
  }

  async function runScan() {
    const paths = state.files.map((f) => f.path);
    if (!paths.length) return;
    try {
      const reply = await toolbox.scan({ paths, options: engineOptions() });
      if (reply.evt === 'error') {
        tbToast(`体检失败：${reply.msg}`, 'error');
        return;
      }
      applyScanReport(reply.report);
    } catch (e) {
      tbToast(`体检失败：${e}`, 'error');
    }
  }

  function applyScanReport(report) {
    const byFile = {};
    for (const f of report.files) byFile[f.path] = f;
    for (const f of state.files) {
      const rec = byFile[f.path];
      if (!rec) continue;
      if (rec.status === 'failed') { f.status = 'bad'; f.reason = rec.reason; f.detections = null; }
      else {
        f.status = 'ok';
        f.detections = rec.sheets.filter((s) => s.header_row !== undefined);
        if (!f.detections.length) { f.status = 'bad'; f.reason = '无可识别的表格'; }
      }
    }
    refreshList();
    const bad = report.totals.files_failed;
    if (bad) tbToast(`${bad} 个文件无法读取，合并时将自动跳过`, 'warn');
  }

  // ---------------- 设置面板 ----------------
  const bindCards = (id, key) => {
    const wrap = $(id);
    const cards = [...wrap.querySelectorAll('.radio-card')];
    cards.forEach((c) => c.classList.toggle('sel', c.dataset.v === cfg[key]));
    cards.forEach((c) => {
      c.onclick = () => {
        cards.forEach((x) => x.classList.remove('sel'));
        c.classList.add('sel');
        cfg[key] = c.dataset.v;
        saveCfg();
      };
    });
  };
  bindCards('opt-mode', 'mode');
  bindCards('opt-align', 'align');
  bindCards('opt-style', 'style');

  const bindSwitch = (id, key, onChange) => {
    const sw = $(id);
    sw.classList.toggle('on', !!cfg[key]);
    const toggle = () => { cfg[key] = !cfg[key]; sw.classList.toggle('on', cfg[key]); saveCfg(); onChange?.(); };
    sw.onclick = (e) => { e.stopPropagation(); toggle(); };
    // 整行 label 可点（点文字也能切换）
    sw.parentElement.addEventListener('click', (e) => {
      if (e.target !== sw) toggle();
    });
  };
  bindSwitch('opt-summary', 'removeSummary');
  bindSwitch('opt-serial-strip', 'stripSerial');
  bindSwitch('opt-serial-num', 'renumberSerial');
  bindSwitch('opt-source', 'source');
  bindSwitch('opt-hidden', 'hidden');
  bindSwitch('opt-title', 'title');

  const kwWrap = $('opt-keywords');
  const renderKw = () => {
    kwWrap.innerHTML = cfg.keywords.map((k) => `<span class="chip" title="点击移除">${k}</span>`).join('');
    kwWrap.querySelectorAll('.chip').forEach((c, i) => {
      c.onclick = () => { cfg.keywords.splice(i, 1); saveCfg(); renderKw(); };
    });
  };
  renderKw();
  $('kw-input').onkeydown = (e) => {
    if (e.key === 'Enter' && e.target.value.trim()) {
      cfg.keywords.push(e.target.value.trim());
      e.target.value = '';
      saveCfg();
      renderKw();
    }
  };

  $('opt-header-auto').value = cfg.headerAuto;
  $('opt-header-row').value = cfg.headerRow;
  $('opt-header-row').classList.toggle('hidden', cfg.headerAuto === '0');
  $('opt-header-auto').onchange = (e) => {
    cfg.headerAuto = e.target.value;
    $('opt-header-row').classList.toggle('hidden', cfg.headerAuto === '0');
    saveCfg();
    scheduleScan();
  };
  $('opt-header-row').onchange = (e) => { cfg.headerRow = Number(e.target.value) || 1; saveCfg(); scheduleScan(); };

  function refreshTemplateSelect() {
    const sel = $('opt-template');
    const cur = Math.min(cfg.templateIndex, Math.max(state.files.length - 1, 0));
    sel.innerHTML = state.files.map((f, i) => `<option value="${i}">${i === 0 ? '默认 · ' : ''}${f.name}</option>`).join('')
      || '<option>（先添加文件）</option>';
    if (state.files.length) sel.value = String(cur);
    sel.onchange = () => { cfg.templateIndex = Number(sel.value); saveCfg(); };
  }

  const toggleSettings = (show) => {
    const body = $('settings-body');
    const open = show ?? body.classList.contains('hidden');
    body.classList.toggle('hidden', !open);
    $('settings').classList.toggle('hidden', false);
    $('settings-head').querySelector('.caret').textContent = open ? '▴' : '▾';
  };
  $('settings-head').onclick = () => toggleSettings();
  $('btn-settings-toggle').onclick = () => { toggleSettings(); $('settings').scrollIntoView({ behavior: 'smooth' }); };

  // ---------------- 合并 ----------------
  function engineOptions() {
    return {
      mode: cfg.mode, align: cfg.align, style: cfg.style,
      template_index: Number(cfg.templateIndex) || 0,
      remove_summary: cfg.removeSummary, summary_keywords: cfg.keywords,
      strip_serial_rows: cfg.stripSerial, renumber_serial: cfg.renumberSerial,
      source_column: cfg.source, include_hidden: cfg.hidden,
      title_as_header: cfg.title,
      force_header_row: cfg.headerAuto === '1' ? (Number(cfg.headerRow) || 1) : 0,
    };
  }

  $('btn-merge').onclick = async () => {
    const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '').replace(/(\d{8})(\d{4})/, '$1_$2');
    const out = await toolbox.saveDialog(`合并结果_${stamp}.xlsx`);
    if (!out) return;
    startMerge(out);
  };

  async function startMerge(out) {
    state.merging = true;
    showProgress();
    const off = toolbox.onProgress((m) => {
      $('prog-file').textContent = m.current || '-';
      $('prog-count').textContent = `${m.file} / ${m.total}`;
      $('prog-rows').textContent = m.rows ?? 0;
      const pct = Math.min(100, Math.round(((m.file || 0) / Math.max(m.total || 1, 1)) * 100));
      $('prog-bar').style.width = `${pct}%`;
    });
    try {
      const reply = await toolbox.merge({
        paths: state.files.map((f) => f.path),
        out,
        options: engineOptions(),
      });
      off();
      state.merging = false;
      if (reply.evt === 'error') {
        hideAllPanes();
        tbToast(`合并失败：${reply.msg}`, 'error');
        return false;
      }
      $('prog-bar').style.width = '100%';
      await new Promise((r) => setTimeout(r, 250));
      showDone(reply.report, out);
      return true;
    } catch (e) {
      off();
      state.merging = false;
      hideAllPanes();
      tbToast(`合并失败：${e}`, 'error');
      return false;
    }
  }

  $('btn-cancel').onclick = () => $('cancel-zone').classList.toggle('hidden', false);
  $('btn-cancel-no').onclick = () => $('cancel-zone').classList.add('hidden');
  $('btn-cancel-yes').onclick = () => { toolbox.cancel(); $('cancel-zone').classList.add('hidden'); };

  function showProgress() {
    hideAllPanes();
    $('progress-pane').classList.remove('hidden');
    $('prog-bar').style.width = '0%';
  }

  function showDone(report, out) {
    hideAllPanes();
    $('done-pane').classList.remove('hidden');
    const t = report.totals || {};
    const cancelled = report.cancelled;
    $('done-icon').textContent = cancelled ? '⏸' : '✓';
    $('done-icon').classList.toggle('cancelled', !!cancelled);
    $('done-title').textContent = cancelled ? '已取消（保留已完成部分）' : '合并完成';
    $('done-stats').textContent =
      `${report.files ? report.files.length : 0} 个文件 · 写入 ${t.rows_written ?? 0} 行 · 耗时 ${report.duration_sec ?? 0}s` +
      (t.summary_rows_removed ? ` · 剔除汇总行 ${t.summary_rows_removed}` : '') +
      (t.big_numbers_converted ? ` · 大数字转文本 ${t.big_numbers_converted}` : '');
    $('done-out').textContent = out;
    $('btn-open-folder').onclick = () => toolbox.showInFolder(out);
    $('btn-again').onclick = () => { hideAllPanes(); $('files-pane') && root.querySelector('.files-pane').classList.remove('hidden'); refreshList(); };
    $('btn-report').onclick = () => $('report-detail').classList.toggle('hidden');

    const lines = [];
    for (const w of report.warnings || []) lines.push(`<div class="r-line r-warn">⚠ ${w}</div>`);
    for (const f of report.files || []) {
      if (f.status === 'failed') lines.push(`<div class="r-line r-bad">✗ ${f.path}：${f.reason}</div>`);
      for (const s of f.sheets || []) {
        if (s.skipped) lines.push(`<div class="r-line">○ ${f.path} ${s.sheet}（跳过）</div>`);
        else lines.push(`<div class="r-line">✓ ${s.display} · 表头第 ${s.header_row} 行</div>`);
      }
    }
    $('report-detail').innerHTML = lines.join('') || '<div class="r-line">无附加信息</div>';
  }

  function hideAllPanes() {
    ['progress-pane', 'done-pane'].forEach((id) => $(id).classList.add('hidden'));
  }

  $('btn-back').onclick = () => ctx.goHome();

  // 初始渲染
  toggleSettings(false);
  refreshList();

  // UI 自动化测试钩子（--uitest 冒烟用；不影响正常使用）
  window.__emTest = { addPaths, startMerge, state };
};
window.toolUnmount = () => { delete window.__emTest; };
