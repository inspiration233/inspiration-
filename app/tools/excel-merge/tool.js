/* global toolbox, tbToast */
// Three independent local workflows. A validated preview is the only export entry.
(() => {
  let dispose = () => {};
  window.toolUnmount = () => dispose();
  window.toolMount = (root, ctx) => {
    const $ = id => root.querySelector('#' + id);
    const make = (tag, cls, text) => { const el = document.createElement(tag); if (cls) el.className = cls; if (text !== undefined) el.textContent = text; return el; };
    const clone = value => JSON.parse(JSON.stringify(value));
    const text = value => value == null ? '' : typeof value === 'object' ? JSON.stringify(value) : String(value);
    const basename = path => path.split(/[\\/]/).pop();
    const colName = n => { let result = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; };
    const cellRef = (row, col) => colName(col) + row;
    const parseCell = raw => { const m = /^\$?([A-Z]+)\$?([1-9]\d*)$/i.exec(raw.trim()); if (!m) throw Error('请使用 A1 这样的单元格坐标。'); return { row: Number(m[2]), col: [...m[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) }; };
    const parseRange = raw => { const parts = raw.trim().split(':'); if (parts.length > 2) throw Error('区域格式应为 A3:F4。'); const a = parseCell(parts[0]), b = parseCell(parts[1] || parts[0]); return { r1: Math.min(a.row, b.row), r2: Math.max(a.row, b.row), c1: Math.min(a.col, b.col), c2: Math.max(a.col, b.col) }; };
    const defaults = () => ({ files: [], feedback: [], baseline: null, active: null, role: 'group', group: [], groupMode: 'combination', filters: [], outputFilters: [], filteredEnabled: false, filteredOptions: {resolutions:{}}, blankGroup: 'separate', outputFields: [], filteredFields: [], keys: [], updates: [], compares: [], groupField: null, resolutions: {}, retained: [], shared: [], options: { mode: 'auto', align: 'loose', shape: 'auto', templateIndex: 0, templateSheet: null, hidden: false, source: false, removeSummary: false, keywords: '合计,小计,总计,汇总,累计', stripSerial: true, minFillRatio: 20, renumber: true, title: false, exampleRows: true, retainAll: false }, result: null, payload: null, version: 0, previewVersion: -1, schemeProblems: [] });
    const sessions = { merge: defaults(), split: defaults(), reconcile: defaults() };
    sessions.reconcile.role = 'keys';
    const state = { mode: 'merge', busy: false, merging: false, disposed: false, gridRequest: 0, operationRequest: 0, picking: null, firstCorner: null, rangeDirty: false, offProgress: null };
    const current = () => sessions[state.mode];
    Object.defineProperty(state, 'files', { get: () => current().files });
    const activeSource = () => { const s = current(); return s.active?.bucket === 'baseline' ? s.baseline : s[s.active?.bucket || 'files']?.[s.active?.index || 0]; };
    const config = source => source?.configs?.[source.sheet];
    const primary = () => current().files[0];
    const primaryConfig = () => config(primary());
    const schemeSource = () => state.mode==='merge'?(current().files[current().options.templateIndex]||primary()):primary();
    const schemeSheet = () => state.mode==='merge'?(current().options.templateSheet||schemeSource()?.sheet):schemeSource()?.sheet;
    const schemeConfig = () => schemeSource()?.configs[schemeSheet()];
    const fieldList = () => primaryConfig()?.report?.fields || [];
    const fieldText = (f, axis = primaryConfig()?.selection?.axis) => `${axis === 'columns' ? '第 ' + f.index + ' 行' : colName(f.index) + ' 列'} · ${f.label}`;
    const fieldByIndex = index => fieldList().find(f => f.index === Number(index));
    const fieldLabel = index => fieldByIndex(index)?.label || `待重选字段 ${index}`;
    // 同名字段(如整行XXX)必须能分辨: 标签带行/列位置
    const fieldChipLabel = index => { const f = fieldByIndex(index); return f ? fieldText(f) : fieldLabel(index); };
    const warningMessages = warnings => [...new Set((warnings||[]).map(w=>typeof w==='string'?w:w?.message||text(w)).map(w=>w.trim()).filter(Boolean))];
    const nativeSource = source => /\.xlsx$/i.test(source?.path || '');
    const matchSignature = field => field?.match_signature ?? field?.signature;
    const fieldContext = fields => (fields||[]).map(f=>({index:f.index,match_signature:matchSignature(f),empty:!!f.empty}));
    function mappingReviews() { const s=current();return state.mode==='merge'?s.files.flatMap(source=>Object.entries(source.configs).filter(([sheet,cfg])=>cfg.mapping_review&&!source.excluded.includes(sheet)&&(s.options.hidden||source.sheets.find(sh=>sh.name===sheet)?.state==='visible')).map(([sheet])=>source.name+' / '+sheet)):[]; }
    function clearMappings() { const s=current();for(const source of [...s.files,...s.feedback,...(s.baseline?[s.baseline]:[])])for(const cfg of Object.values(source.configs)){cfg.field_map={};cfg.mapping_review=false;} }
    function preserveMappings(source, sheet, before, after) {
      if (!before) return;
      const oldFields=before.report.fields||[],newFields=after.report.fields||[];
      const unchanged=before.selection.axis===after.selection.axis&&JSON.stringify(oldFields.map(f=>[f.index,matchSignature(f)]))===JSON.stringify(newFields.map(f=>[f.index,matchSignature(f)]));
      if (unchanged) return; // A data-range edit does not change field correspondence.
      const relocate=index=>{
        const old=oldFields.find(f=>f.index===Number(index));if(!old)return null;
        if(old.empty||String(matchSignature(old)).startsWith('@empty:'))return null;
        const previous=oldFields.filter(f=>matchSignature(f)===matchSignature(old)),matches=newFields.filter(f=>matchSignature(f)===matchSignature(old));
        return previous.length===1&&matches.length===1?matches[0].index:null;
      };
      const targetChanged=source===schemeSource()&&sheet===schemeSheet();
      const s=current();
      for(const file of [...s.files,...s.feedback,...(s.baseline?[s.baseline]:[])])for(const cfg of Object.values(file.configs)){
        if(cfg!==after&&!targetChanged)continue;
        const kept={};let lost=false;
        for(const [target,index] of Object.entries(cfg.field_map||{})){
          const targetIndex=targetChanged?relocate(target):Number(target);
          const sourceIndex=index==null?null:cfg===after?relocate(index):Number(index);
          if(targetIndex!=null&&(index==null||sourceIndex!=null))kept[targetIndex]=sourceIndex;
          else lost=true;
        }
        cfg.field_map=kept;
        if(lost)cfg.mapping_review=true;
      }
    }
    function clearRoles() { const s=current();s.group=[];s.keys=[];s.updates=[];s.compares=[];s.filters=[];s.outputFilters=[];s.outputFields=[];s.filteredFields=[];s.schemeProblems=[];clearMappings(); }
    const notice = (message = '', error = false) => { $('notice').textContent = message; $('notice').classList.toggle('hidden', !message); $('notice').classList.toggle('error', error); };
    const fail = error => { notice(error?.message || String(error), true); if (typeof tbToast === 'function') tbToast(error?.message || String(error), 'error'); };
    const call = async (method, payload) => { if (typeof toolbox[method] !== 'function') throw Error('当前引擎尚未提供此功能，请重新启动更新后的程序。'); const reply = await toolbox[method](payload); if (reply?.evt === 'error') throw Error(reply.msg || '处理失败'); return reply?.report ?? reply; };
    const option = (value, label) => { const el = make('option', '', label); el.value = String(value); return el; };
    const setOptions = (select, values, selected) => { select.replaceChildren(...values.map(([v, l]) => option(v, l))); if (selected !== undefined) select.value = String(selected); };
    const isPrimary = source => source === primary();

    function invalidate(message = '设置已变化，请重新预览。', keepResolutions = false) {
      const s = current(); s.version++; s.result = null; s.payload = null; s.previewVersion = -1;
      if (!keepResolutions) { s.resolutions = {}; s.retained = []; s.shared = []; s.filteredOptions = {resolutions:{}}; }
      $('result-pane').classList.add('hidden'); $('btn-export').disabled = true;
      $('action-hint').textContent = message;
      updateAction();
    }
    function updateAction() {
      const pendingMappings=mappingReviews();
      const ready = current().files.length > 0 && !state.busy && !state.rangeDirty && !current().schemeProblems.length&&!pendingMappings.length;
      $('btn-preview').disabled = !ready;
      $('btn-preview').textContent = state.busy ? '正在预览…' : '预览结果';
      $('btn-export').disabled = state.busy || !current().result?.can_export || current().previewVersion !== current().version;
      if (!current().files.length) $('action-hint').textContent = '先选择文件';
      if (state.rangeDirty) $('action-hint').textContent = '请先应用刚刚修改的表格范围';
      if (current().schemeProblems.length) $('action-hint').textContent = '方案有未能唯一对应的字段，请补选并确认。';
      if(pendingMappings.length)$('action-hint').textContent='表头发生变化，请核对字段对应：'+pendingMappings.join('、');
    }
    function newSource(path) { return { path, name: basename(path), configs: {}, sheets: [], sheet: null, excluded: [], status: 'pending' }; }
    async function loadSource(source, sheet = source.sheet, page, selection, selectionOrigin) {
      const cfg = source.configs[sheet];
      const origin=selectionOrigin||(selection?'manual':cfg?.selection_origin||'auto');
      const selected=selection||(origin==='manual'?cfg?.selection:null);
      const report = await call('preview', { path: source.path, ...(sheet ? { sheet } : {}), ...(selected ? { selection:selected } : {}), selection_origin:origin, start_row: page?.row || 1, start_col: page?.col || 1, row_limit: 80, col_limit: 40 });
      source.sheets = report.sheets || []; source.status = 'ok'; source.reason = '';
      if (!source.sheet) source.sheet = report.sheet;
      const previous = source.configs[report.sheet];
      source.configs[report.sheet] = { ...previous, selection: report.selection, selection_origin:origin, report, field_map: previous?.field_map || {}, page: { row: report.viewport?.start_row || page?.row || 1, col: report.viewport?.start_col || page?.col || 1 } };
      const next=source.configs[report.sheet];
      preserveMappings(source,report.sheet,previous,next);
      // Explicit range/plan changes already invalidate in their owning action.
      // A plain reread must also invalidate a result made with the old layout.
      if(previous&&!selection&&(previous.report.signature!==report.signature||JSON.stringify(previous.selection)!==JSON.stringify(next.selection)))invalidate('原表结构或范围已变化，字段对应已重新核对，请重新预览。');
      return source.configs[report.sheet];
    }
    async function showSource(source, bucket = 'files', index = 0, sheet = source?.sheet, page) {
      if (!source) { $('workspace').classList.add('hidden'); return; }
      current().active = { bucket, index }; state.picking = null; state.firstCorner = null; state.rangeDirty = false;
      const request = ++state.gridRequest, mode = state.mode;
      const newlyOpened = $('workspace').classList.contains('hidden');
      $('workspace').classList.remove('hidden'); $('grid-status').textContent = '正在读取原表…';
      renderFiles();
      try {
        if (sheet) source.sheet = sheet;
        const cfg = await loadSource(source, sheet, page);
        if (request !== state.gridRequest || mode !== state.mode || state.disposed) return;
        source.sheet = cfg.report.sheet;
        // 文件加入后工作区在首屏之下, 不自动滚过去用户根本找不到点击区
        if (newlyOpened) $('workspace').scrollIntoView({ behavior: 'smooth', block: 'start' });
        renderWorkspace(); renderRoles(); renderFilters(); renderFiles(); updateAction();
        await renderMapping();
      } catch (error) {
        if (request !== state.gridRequest || state.disposed) return;
        source.status = 'bad'; source.reason = error.message; $('grid-wrap').replaceChildren(make('p', 'notice error', error.message)); $('grid-status').textContent = '无法读取'; renderFiles(); fail(error);
      }
    }
    async function addPaths(input, bucket = 'files') {
      const s = current(), paths = Array.isArray(input) ? input : input?.files || [];
      if (!paths.length) return;
      if (bucket === 'files' && state.mode !== 'merge') { clearRoles();s.files = [newSource(paths[0])]; if (paths.length > 1) notice('当前操作使用一份底表，已选择第一个文件。其他文件可在反馈区添加。'); }
      else if (bucket === 'baseline') s.baseline = newSource(paths[0]);
      else for (const path of paths) if (!s[bucket].some(f => f.path.toLowerCase() === path.toLowerCase())) s[bucket].push(newSource(path));
      invalidate('已添加文件，请在原表上确认范围。');
      const source = bucket === 'baseline' ? s.baseline : s[bucket][bucket === 'feedback' ? s[bucket].length - 1 : 0];
      renderFiles(); await showSource(source, bucket, bucket === 'feedback' ? s[bucket].length - 1 : 0);
      if (state.mode === 'merge') recommend();
    }
    async function addFromPaths(paths, bucket = 'files') {
      if (!paths?.length) return;
      try { await addPaths(await toolbox.expandPaths(paths), bucket); } catch (error) { fail(error); }
    }
    function renderFiles() {
      const s = current();
      const render = (container, files, bucket) => {
        container.replaceChildren();
        files.forEach((file, index) => {
          const row = make('div', 'file-row' + (s.active?.bucket === bucket && s.active?.index === index ? ' active' : ''));
          if (bucket === 'files' && state.mode === 'merge') {
            row.draggable = true; row.append(make('span', 'drag-handle', '⠿'));
            row.ondragstart = e => { e.dataTransfer.setData('application/x-excel-source', String(index)); e.dataTransfer.effectAllowed = 'move'; };
            row.ondragover = e => e.preventDefault();
            row.ondrop = e => { const raw = e.dataTransfer.getData('application/x-excel-source'); if (!raw) return; e.preventDefault(); e.stopPropagation(); const from = Number(raw); const selected = activeSource(),template=s.files[s.options.templateIndex]; const [fileToMove] = s.files.splice(from, 1); s.files.splice(index, 0, fileToMove); s.options.templateIndex=Math.max(0,s.files.indexOf(template));s.active = { bucket, index: s.files.indexOf(selected) }; invalidate(); renderFiles(); };
          }
          const name = make('span', 'file-name', file.name); name.title = file.path;
          row.append(name, make('span', 'file-meta', file.status === 'bad' ? file.reason : file.sheets.length ? `${file.sheets.length} 张工作表` : '点击预览'));
          const remove = make('button', 'ghost rm', '×'); remove.setAttribute('aria-label', '移除 ' + file.name);
          remove.onclick = e => { e.stopPropagation(); const template=s.files[s.options.templateIndex];files.splice(index, 1);if(bucket==='files'){if(state.mode!=='merge')clearRoles();else clearMappings();}s.options.templateIndex=Math.max(0,s.files.indexOf(template));invalidate(); s.active = null; renderFiles(); showSource(s.files[0]); };
          row.append(remove); row.onclick = () => showSource(file, bucket, index); container.append(row);
        });
      };
      render($('file-rows'), s.files, 'files'); render($('feedback-rows'), s.feedback, 'feedback');
      $('feedback-count').textContent = s.feedback.length ? `· ${s.feedback.length} 份` : '';
      $('baseline-name').textContent = s.baseline?.name || '未选择'; $('baseline-name').style.cursor = s.baseline ? 'pointer' : '';
      $('operation-options').classList.toggle('hidden', !s.files.length);
      setOptions($('opt-template'), s.files.map((f, i) => [i, f.name]), Math.min(s.options.templateIndex, s.files.length - 1));
      const template=s.files[s.options.templateIndex]||s.files[0];if(template?.sheets.length&&!template.sheets.some(sh=>sh.name===s.options.templateSheet))s.options.templateSheet=template.sheet||template.sheets[0].name;
      setOptions($('opt-template-sheet'),template?.sheets.map(sh=>[sh.name,sh.name])||[],s.options.templateSheet);
      updateAction();
    }
    function rangesFromSelection(selection) {
      if (selection.axis === 'columns') return { header: `${cellRef(selection.field_start, selection.header_start)}:${cellRef(selection.field_end, selection.header_end)}`, data: `${cellRef(selection.field_start, selection.data_start)}:${cellRef(selection.field_end, selection.data_end)}` };
      return { header: `${cellRef(selection.header_start, selection.field_start)}:${cellRef(selection.header_end, selection.field_end)}`, data: `${cellRef(selection.data_start, selection.field_start)}:${cellRef(selection.data_end, selection.field_end)}` };
    }
    function renderWorkspace() {
      const source = activeSource(), cfg = config(source); if (!cfg) return;
      const report = cfg.report, ranges = rangesFromSelection(cfg.selection);
      $('source-caption').textContent = `${current().active?.bucket === 'feedback' ? '反馈 · ' : current().active?.bucket === 'baseline' ? '原始版本 · ' : ''}${source.name}`;
      setOptions($('sheet-select'), source.sheets.map(s => [s.name, s.name + (s.state !== 'visible' ? '（隐藏）' : '')]), source.sheet);
      $('sheet-enabled-wrap').classList.toggle('hidden', state.mode !== 'merge'); $('sheet-enabled').checked = !source.excluded.includes(source.sheet);
      $('axis-select').value = cfg.selection.axis; $('header-range').value = ranges.header; $('data-range').value = ranges.data;
      // 横向配置表(记录沿列)由程序识别范围, 手动改向/框选会破坏"字段=行,记录=列"语义
      const crossSheet = !!cfg.report?.recognition?.native_cross;
      for(const id of ['axis-select','header-range','data-range','pick-header','pick-data','apply-range'])$(id).disabled=!nativeSource(source)||crossSheet;
      $('source-warnings').textContent = warningMessages(report.warnings).join('；');
      $('source-warnings').classList.toggle('hidden',!$('source-warnings').textContent);
      renderGrid(report); renderToolbar(); pickHint();
    }
    function applyCellStyle(td, style) {
      const allowed = ['fontFamily','fontSize','color','background','fontWeight','fontStyle','textAlign','verticalAlign','whiteSpace','textDecoration','borderTop','borderBottom','borderLeft','borderRight'];
      for (const key of allowed) if (style?.[key] != null) td.style[key] = String(style[key]);
    }
    function roleFields() { const s = current(); return state.mode === 'split' ? (s.role==='outputFields'?s.outputFields:s.group.filter(g => g.kind === 'field').map(g => g.index)) : state.mode === 'reconcile' ? (s.role==='groupField'?[s.groupField].filter(v=>v!=null):s[s.role] || []) : []; }
    function renderGrid(report) {
      const wrap = $('grid-wrap'), oldLeft = wrap.scrollLeft, oldTop = wrap.scrollTop;
      const viewport = report.viewport || { start_row: 1, start_col: 1, end_row: Math.min(report.max_row, 80), end_col: Math.min(report.max_column, 40) };
      const { start_row: sr, end_row: er, start_col: sc, end_col: ec } = viewport;
      const table = make('table', 'excel-grid'); table.setAttribute('aria-label', report.sheet + ' 原表');
      const cg = make('colgroup'); const gutter = make('col'); gutter.style.width = '42px'; cg.append(gutter);
      for (let c = sc; c <= ec; c++) { const col = make('col'); col.style.width = `${Math.max(25, Math.min(600, report.column_widths?.[c] || 100))}px`; cg.append(col); }
      table.append(cg);
      const fields = new Set((report.fields || []).map(f => f.index)), selected = new Set(isPrimary(activeSource()) ? roleFields() : []), sel = report.selection;
      const thead = make('thead'), heading = make('tr'); heading.append(make('th', 'corner', ''));
      for (let c = sc; c <= ec; c++) { const th = make('th', sel.axis === 'rows' && fields.has(c) ? 'field-head' : '', colName(c)); th.dataset.col = c; if (sel.axis === 'rows' && selected.has(c)) th.classList.add('selected'); th.onclick = () => selectField(c, 'rows'); heading.append(th); }
      thead.append(heading); table.append(thead);
      const lookup = new Map((report.cells || []).map(cell => [`${cell.row}:${cell.col}`, cell]));
      const covered = new Set(), merged = new Map();
      for (const m of report.merges || []) {
        const r1 = Math.max(sr, m.min_row), r2 = Math.min(er, m.max_row), c1 = Math.max(sc, m.min_col), c2 = Math.min(ec, m.max_col);
        if (r2 < r1 || c2 < c1) continue;
        merged.set(`${r1}:${c1}`, { ...m, rowSpan: r2-r1+1, colSpan: c2-c1+1 });
        for (let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++) if(r!==r1||c!==c1) covered.add(`${r}:${c}`);
      }
      const tbody = make('tbody');
      for (let r = sr; r <= er; r++) {
        const tr = make('tr'); tr.style.height = `${Math.max(16, report.row_heights?.[r] || 20)}px`;
        const th = make('th', sel.axis === 'columns' && fields.has(r) ? 'field-head' : '', r); th.dataset.row = r; if (sel.axis === 'columns' && selected.has(r)) th.classList.add('selected'); th.onclick = () => selectField(r, 'columns'); tr.append(th);
        for (let c = sc; c <= ec; c++) {
          const key = `${r}:${c}`; if (covered.has(key)) continue;
          const merge = merged.get(key), cell = (merge ? lookup.get(`${merge.min_row}:${merge.min_col}`) : null) || lookup.get(key);
          const td = make('td', '', cell?.text ?? cell?.value ?? ''); td.dataset.row = r; td.dataset.col = c;
          td.title = `${cellRef(r,c)}${cell?.formula ? ' · '+cell.formula : ''}${merge ? ' · 合并区域 '+cellRef(merge.min_row,merge.min_col)+':'+cellRef(merge.max_row,merge.max_col) : ''}`;
          if (merge) { td.rowSpan = merge.rowSpan; td.colSpan = merge.colSpan; }
          applyCellStyle(td, cell?.style);
          const record = sel.axis === 'rows' ? r : c, field = sel.axis === 'rows' ? c : r;
          if (field >= sel.field_start && field <= sel.field_end) {
            if (record >= sel.header_start && record <= sel.header_end) td.classList.add('in-header');
            if (record >= sel.data_start && record <= sel.data_end) td.classList.add('in-data');
            if (selected.has(field)) td.classList.add('selected-field');
          }
          if (state.firstCorner?.row === r && state.firstCorner?.col === c) td.classList.add('corner-picked');
          td.onclick = () => cellClicked(r,c,cell); tr.append(td);
        }
        tbody.append(tr);
      }
      table.append(tbody); wrap.replaceChildren(table); wrap.scrollLeft = oldLeft; wrap.scrollTop = oldTop; applyGridZoom();
      $('grid-status').textContent = `${report.max_row} 行 × ${report.max_column} 列 · 当前 ${cellRef(sr,sc)}:${cellRef(er,ec)}${report.truncated ? '（分段预览）' : ''}`;
      $('page-prev').disabled = sr <= 1; $('page-next').disabled = er >= report.max_row; $('page-left').disabled = sc <= 1; $('page-right').disabled = ec >= report.max_column;
    }
    // 预览缩放: 默认"适应列宽"(不低于35%), 可手动增减。
    // 用 CSS zoom(布局级缩放)而不是 transform: 滚动条/冻结表头/跳转/点击
    // 坐标全部按缩放后的真实布局工作, 不需要任何坐标换算补丁。
    const GRID_MIN_ZOOM = 0.35;
    const gridZoom = { mode: 'fit', scale: 1 };
    function fitZoom(table, wrap) {
      table.style.zoom = '';
      const natural = table.scrollWidth;
      let z = Math.min(1, Math.max(GRID_MIN_ZOOM, (wrap.clientWidth - 2) / Math.max(1, natural)));
      table.style.zoom = z;
      // 应用后容器可能因滚动条出现/消失而变宽窄, 再校正一次
      const corrected = Math.min(1, Math.max(GRID_MIN_ZOOM, (wrap.clientWidth - 2) / Math.max(1, natural)));
      if (Math.abs(corrected - z) > 0.005) { z = corrected; table.style.zoom = z; }
      return z;
    }
    function applyGridZoom() {
      const wrap = $('grid-wrap'), table = wrap.querySelector('table');
      if (!table) return;
      gridZoom.scale = gridZoom.mode === 'fit' ? fitZoom(table, wrap) : gridZoom.scale;
      table.style.zoom = gridZoom.scale;
      $('zoom-level').textContent = (gridZoom.mode === 'fit' ? '适应 ' : '') + Math.round(gridZoom.scale * 100) + '%';
    }
    function syncModals() {
      const open = ['done-pane','progress-pane'].some(id => !$(id).classList.contains('hidden') || $(id).classList.contains('closing'));
      document.body.classList.toggle('modal-open', open);
    }
    // 弹窗动效: 开=背景淡入+卡片弹性上浮, 关=反向退场后再隐藏(有开有关)
    function openPane(id) {
      const el = $(id);
      el.classList.remove('closing'); el.classList.remove('hidden');
      void el.offsetWidth;                     // 重新触发进入动画
      syncModals();
    }
    function closePane(id) {
      const el = $(id);
      if (el.classList.contains('hidden') || el.classList.contains('closing')) return;
      el.classList.add('closing');
      let fired = false;
      const finish = () => {
        if (fired) return; fired = true;
        el.classList.remove('closing'); el.classList.add('hidden'); syncModals();
      };
      el.addEventListener('animationend', e => { if (e.target === el) finish(); }, { once: true });
      setTimeout(finish, 300);                 // 兜底: 动画事件缺席也能收尾
    }
    function renderToolbar() {
      const bar = $('field-toolbar'); bar.replaceChildren();
      if (state.mode === 'merge') { bar.append(make('span', '', '蓝色=标题行，浅蓝=数据行；识别有误时可修改上方范围。')); return; }
      if (!isPrimary(activeSource())) { bar.append(make('span', '', '当前查看反馈或旧版本；选择列请先点选底稿文件。')); return; }
      bar.append(make('span', '', '选用途，再点列标题：'));
      const choices = state.mode === 'split' ? [['group','拆分依据'],['outputFields','输出字段']] : [['groupField','单位列（分组）'],['keys','认记录（主键）'],['updates','允许修改的列'],['compares','改后标红的列'],...(current().filteredEnabled?[['filteredFields','筛选结果字段']]:[])];
      for (const [value,label] of choices) { const button = make('button', current().role === value ? 'active' : '', label); button.onclick = () => { current().role = value; renderToolbar(); renderGrid(config(activeSource()).report); }; bar.append(button); }
      // 选中反馈就在网格正上方——chips 区可能在屏幕外, 这里给即时可见的确认
      const picked = roleFields();
      if (picked.length) bar.append(make('span', 'hint', `已选 ${picked.length} 项：${picked.slice(0, 4).map(fieldChipLabel).join('、')}${picked.length > 4 ? ' 等' : ''}`));
    }
    function selectField(index, axis) {
      const source = activeSource(), cfg = config(source); if (!cfg || state.mode === 'merge') return;
      if (cfg.selection.axis !== axis) {
        // 点错方向不给"静默无效": 明确告知字段在哪一侧
        notice(cfg.selection.axis === 'columns'
          ? '这张表是横向布局（记录沿列）：字段排在左侧行号上，请点击行号选择；如需按列处理，可先在上方把方向改为纵向。'
          : '这张表是纵向布局：字段排在顶部列号上，请点击列号选择；如需按行处理，可先在上方把方向改为横向。');
        return;
      }
      if (!isPrimary(source)) { notice('请先点选底稿文件，再选择用于拆分或回填的字段。'); return; }
      if (!cfg.report.fields.some(f => f.index === index)) { notice('这个行或列不在当前字段范围内。请先调整范围。'); return; }
      const s = current();
      if (state.mode === 'reconcile' && s.role === 'groupField') { s.groupField = s.groupField === index ? null : index; }
      else if (state.mode === 'split'&&s.role!=='outputFields') { const pos = s.group.findIndex(g => g.kind === 'field' && g.index === index); if (pos >= 0) s.group.splice(pos,1); else s.group.push({kind:'field',index}); }
      else { const list = s[s.role], pos = list.indexOf(index); if (pos >= 0) list.splice(pos,1); else list.push(index); }
      invalidate(); renderRoles(); renderToolbar(); renderGrid(cfg.report);
    }
    function pickHint() {
      $('selection-hint').classList.toggle('picking', !!state.picking);
      $('selection-hint').textContent = state.picking === 'cell' ? '点一个格子，以它的文字命名。' : state.picking ? `${state.picking === 'header' ? '标题行' : '数据行'}圈选：${state.firstCorner ? '已点 '+cellRef(state.firstCorner.row,state.firstCorner.col)+'，请点对角单元格。' : '先点一个角。'} 再点按钮取消。` : config(activeSource())?.selection_origin==='manual'?'已按圈选应用；如需调整可修改上方范围。':'范围已自动识别，可直接预览；如需调整请修改上方范围。';
    }
    function startPick(kind) { if (!config(activeSource())) return; state.picking = state.picking === kind ? null : kind; state.firstCorner = null; pickHint(); }
    function cellClicked(row,col,cell) {
      if (!state.picking) { $('selection-hint').textContent = `${cellRef(row,col)} · ${text(cell?.formula || cell?.value || '')}`; return; }
      if (state.picking === 'cell') {
        if (!isPrimary(activeSource())) { notice('请先回到底稿选择命名单元格。'); return; }
        const ref = cellRef(row,col); if (!current().group.some(g => g.kind === 'cell' && g.ref === ref)) current().group.push({kind:'cell',ref,expected:cell?.value});
        state.picking = null; invalidate(); renderRoles(); pickHint(); return;
      }
      if (!state.firstCorner) { state.firstCorner = {row,col}; renderGrid(config(activeSource()).report); pickHint(); return; }
      const a = state.firstCorner; $(state.picking === 'header' ? 'header-range' : 'data-range').value = `${cellRef(Math.min(row,a.row),Math.min(col,a.col))}:${cellRef(Math.max(row,a.row),Math.max(col,a.col))}`;
      state.picking = null; state.firstCorner = null; state.rangeDirty = true; invalidate(); pickHint(); renderGrid(config(activeSource()).report);
    }
    async function applyRanges() {
      const source = activeSource(); if (!source) return false;
      const request=++state.gridRequest,mode=state.mode,version=current().version;
      try {
        const header = parseRange($('header-range').value), data = parseRange($('data-range').value), axis = $('axis-select').value;
        const rows = axis === 'rows';
        if (rows ? header.c1 !== data.c1 || header.c2 !== data.c2 : header.r1 !== data.r1 || header.r2 !== data.r2) throw Error('表头区域与数据区域必须覆盖相同的字段。');
        const selection = {axis,header_start:rows?header.r1:header.c1,header_end:rows?header.r2:header.c2,data_start:rows?data.r1:data.c1,data_end:rows?data.r2:data.c2,field_start:rows?header.c1:header.r1,field_end:rows?header.c2:header.r2};
        const previous = config(source)?.report?.fields || [], saved = isPrimary(source) ? captureRoles(previous) : null;
        const cfg = await loadSource(source,source.sheet,config(source)?.page,selection);
        if(request!==state.gridRequest||mode!==state.mode||version!==current().version||source!==activeSource())return false;
        state.rangeDirty = false;invalidate();
        if (saved&&JSON.stringify(previous)!==JSON.stringify(cfg.report.fields)) restoreRoles(saved,cfg.report.fields,false);
        renderWorkspace(); renderRoles(); renderFilters(); await renderMapping(); updateAction(); return true;
      } catch(error) { fail(error); return false; }
    }
    function renderRoles() {
      const s = current();
      const chips = (id, values, label, remove, roleKey) => { const box = $(id); box.replaceChildren();
        if (!values.length) box.append(make('span','empty-chip','尚未选择（点表格的 A、B、C 列头）'));
        // 入口按钮常驻: 先把"当前点选用途"切到本区, 再跳到表格——
        // 否则列头会一直加进上一次的用途里
        const go = make('button','ghost compact','＋ 在表上点选');
        go.onclick = async () => { const s=current(); if (!s.files.length) { notice('请先在上方选择文件。', true); return; }
          if (roleKey && s.role !== roleKey) { s.role = roleKey; renderToolbar(); }
          // 正在查看反馈/原始版本时, 点列会提示"请先点选底表"——入口按钮负责带回底表
          const active = s.active;
          if (!active || active.bucket !== 'files') { await showSource(s.files[0], 'files', 0); }
          $('workspace').scrollIntoView({behavior:'smooth',block:'start'}); };
        values.forEach((v,i) => { const chip=make('button','field-chip',label(v)+' ×'); chip.onclick=()=>{remove(i);invalidate();renderRoles();renderToolbar();if(config(activeSource()))renderGrid(config(activeSource()).report);};box.append(chip); });
        box.append(go); };
      chips('group-chips',s.group,g=>g.kind==='field'?fieldChipLabel(g.index):g.kind==='cell'?'单元格 '+g.ref:'工作表名',i=>s.group.splice(i,1),'group');
      const unitBox=$('group-unit-chips');unitBox.replaceChildren();
      const unitGo=make('button','ghost compact','＋ 在表上点选');
      unitGo.onclick=async()=>{const s=current();if(!s.files.length){notice('请先在上方选择文件。',true);return;}
        if(s.role!=='groupField'){s.role='groupField';renderToolbar();}
        const active=s.active;
        if(!active||active.bucket!=='files'){await showSource(s.files[0],'files',0);}
        $('workspace').scrollIntoView({behavior:'smooth',block:'start'});};
      if(s.groupField!=null){const chip=make('button','field-chip',fieldChipLabel(s.groupField)+' ×');chip.onclick=()=>{s.groupField=null;invalidate();renderRoles();renderToolbar();if(config(activeSource()))renderGrid(config(activeSource()).report);};unitBox.append(chip);}
      else unitBox.append(make('span','empty-chip','未选择 · 按主键值直接匹配（不分组）'));
      unitBox.append(unitGo);
      for(const [id,key] of [['key-chips','keys'],['update-chips','updates'],['compare-chips','compares'],['output-chips','outputFields'],['filtered-field-chips','filteredFields']]) chips(id,s[key],fieldChipLabel,i=>s[key].splice(i,1),key);
      $('blank-group').value=s.blankGroup;$('split-example-rows').checked=!!s.options.exampleRows;$('split-retain-all').checked=!!s.options.retainAll;
      $('group-by-sheet').checked=s.group.some(g=>g.kind==='sheet');$('group-mode').value=s.groupMode;
      renderToolbar();
    }
    function fieldSelect(selected, label = '选择字段') { const select=make('select');setOptions(select,[['',label],...fieldList().map(f=>[f.index,fieldText(f)])],selected??'');return select; }
    function renderFilters() {
      const s=current();$('filter-count').textContent=s.filters.length?`· ${s.filters.length} 条`:'';
      $('filtered-enabled').checked=s.filteredEnabled;$('filtered-controls').classList.toggle('hidden',!s.filteredEnabled);
      for(const [boxId,list] of [['filter-rows',s.filters],['filtered-rows',s.outputFilters]]){const box=$(boxId);box.replaceChildren();
      list.forEach((filter,index)=>{
        const row=make('div','filter-row'),field=fieldSelect(filter.field),op=make('select'),value=make('input'),end=fieldSelect(filter.end_field,'选择截止字段'),rm=make('button','ghost','×');value.type='text';value.value=filter.value||'';value.placeholder=filter.op==='month'?'3 或 2026-03':filter.op==='interval'?'目标值或月份':['gt','lt','ge','le'].includes(filter.op)?'要比较的数字':filter.op==='oneof'?'多个值，逗号分隔':'要保留的内容';
        setOptions(op,[['equals','等于'],['contains','包含文字'],['oneof','任一匹配（逗号分隔）'],['gt','大于'],['lt','小于'],['ge','不小于'],['le','不大于'],['month','属于月份'],['interval','起止区间覆盖']],filter.op||'equals');
        end.classList.toggle('hidden',filter.op!=='interval');
        field.onchange=()=>{filter.field=field.value?Number(field.value):null;invalidate();};op.onchange=()=>{filter.op=op.value;invalidate();renderFilters();};value.oninput=()=>{filter.value=value.value;invalidate();};end.onchange=()=>{filter.end_field=end.value?Number(end.value):null;invalidate();};rm.onclick=()=>{list.splice(index,1);invalidate();renderFilters();};
        row.append(field,op,end,value,rm);box.append(row);
      });}
    }
    function autoMap(targetFields, sourceFields) {
      const result={};
      for(const target of targetFields){const signature=matchSignature(target);if(!signature||target.empty||String(signature).startsWith('@empty:'))continue;const t=targetFields.filter(f=>matchSignature(f)===signature),matches=sourceFields.filter(f=>!f.empty&&matchSignature(f)===signature);if(t.length===1&&matches.length===1)result[target.index]=matches[0].index;}
      return result;
    }
    async function mappingTarget(source) {
      if(state.mode==='reconcile')return primaryConfig();
      if(state.mode!=='merge')return null;
      const target=current().files[current().options.templateIndex]||primary();if(!target)return null;
      if(!target.sheets.length)await loadSource(target);
      const sheet=current().options.templateSheet||target.sheet;
      if(target===source&&sheet===source.sheet)return null;
      return target.configs[sheet]||await loadSource(target,sheet);
    }
    async function renderMapping() {
      const source=activeSource(),cfg=config(source),panel=$('mapping-details');
      if(!cfg||state.mode==='split'||(state.mode==='reconcile'&&isPrimary(source))){panel.classList.add('hidden');return;}
      const mode=state.mode,request=state.gridRequest;
      const target=await mappingTarget(source);if(request!==state.gridRequest||mode!==state.mode)return;
      if(!target){panel.classList.add('hidden');return;}
      panel.classList.remove('hidden');const targetFields=target.report.fields||[],sourceFields=cfg.report.fields||[];
      $('mapping-review').classList.toggle('hidden',!cfg.mapping_review);if(cfg.mapping_review)panel.open=true;
      const mapping={...autoMap(targetFields,sourceFields),...cfg.field_map};
      const table=make('table','mapping-table'),head=make('tr');head.append(make('th','','底表 / 模板字段'),make('th','','当前文件字段'));table.append(head);
      let missing=0;
      for(const field of targetFields){const tr=make('tr'),td=make('td'),select=make('select');setOptions(select,[['','不对应 / 待选择'],...sourceFields.map(f=>[f.index,fieldText(f,cfg.selection.axis)])],mapping[field.index]??'');if(!select.value)missing++;select.onchange=()=>{cfg.field_map[field.index]=select.value?Number(select.value):null;invalidate();};tr.append(make('td','',fieldText(field,target.selection.axis)));td.append(select);tr.append(td);table.append(tr);}
      $('mapping-table').replaceChildren(table);$('mapping-summary').textContent=missing?`· ${missing} 个未对应`:'· 已按唯一同名字段对应';
    }
    function recommend() {
      const s=current(),sources=s.files.filter(f=>f.sheets.length);
      let mode='single',reason='把记录汇入一张表';
      if(sources.length===1&&sources[0].sheets.length>1){mode='sheets';reason='汇总当前文件中的工作表';}
      else if(sources.length>1){const first=sources[0].sheets.filter(sh=>!sources[0].excluded.includes(sh.name)&&(s.options.hidden||sh.state==='visible')).map(sh=>sh.name);if(first.length>1&&sources.every(f=>{const names=f.sheets.filter(sh=>!f.excluded.includes(sh.name)&&(s.options.hidden||sh.state==='visible')).map(sh=>sh.name);return names.length===first.length&&first.every(name=>names.includes(name));})){mode='group';reason='同名工作表分别汇总';}}
      s.recommendation={mode,reason};$('mode-reco').textContent=s.options.mode==='auto'?`将生成：${reason}`:'';return mode;
    }
    function engineOptions() { const o=current().options;return {mode:o.mode==='auto'?recommend():o.mode,align:o.align,style:'fidelity',merge_shape:o.shape,template_index:Number(o.templateIndex)||0,remove_summary:o.removeSummary,summary_keywords:o.keywords.split(/[,，]/).map(k=>k.trim()).filter(Boolean),strip_serial_rows:o.stripSerial,renumber_serial:o.renumber,source_column:o.source,include_hidden:o.hidden,title_as_header:o.title,force_header_row:0, min_fill_ratio:(Number.isFinite(Number(o.minFillRatio))?Math.max(0,Number(o.minFillRatio)):20)/100}; }
    function renderOptions() {
      const s=current(),o=s.options;
      root.querySelectorAll('[data-mode]').forEach(b=>{b.classList.toggle('active',b.dataset.mode===state.mode);b.setAttribute('aria-pressed',String(b.dataset.mode===state.mode));});
      $('source-title').textContent={merge:'第 1 步 · 选择文件',split:'第 1 步 · 选择文件',reconcile:'第 1 步 · 选底稿与反馈'}[state.mode];
      $('source-help').textContent=state.mode==='merge'?'可一次多选，也可直接拖入。':'底稿：要保留并更新的表。';
      $('btn-add-folder').classList.toggle('hidden',state.mode!=='merge');$('reconcile-sources').classList.toggle('hidden',state.mode!=='reconcile');
      for(const mode of ['merge','split','reconcile'])$(mode+'-options').classList.toggle('hidden',mode!==state.mode);
      root.querySelectorAll('[name="merge-output"]').forEach(input=>input.checked=input.value===o.mode);
      for(const [id,key] of [['opt-shape','shape'],['opt-align','align'],['opt-keywords','keywords']])$(id).value=o[key];
      $('opt-fill-ratio').value=o.minFillRatio;
      for(const [id,key] of [['opt-hidden','hidden'],['opt-source','source'],['opt-summary','removeSummary'],['opt-serial-strip','stripSerial'],['opt-serial-num','renumber'],['opt-title','title']])$(id).checked=o[key];
      renderFiles();renderRoles();renderFilters();renderSchemes();
    }
    async function switchMode(mode) {
      if(state.busy||state.merging||!sessions[mode])return;
      state.mode=mode;state.gridRequest++;state.operationRequest++;state.rangeDirty=false;notice();renderOptions();
      // 切模式是场景重置: 立即关闭, 不播退场动画
      $('done-pane').classList.remove('closing');$('progress-pane').classList.remove('closing');
      $('done-pane').classList.add('hidden');$('progress-pane').classList.add('hidden');syncModals();
      $('result-pane').classList.add('hidden');$('workspace').classList.add('hidden');
      const s=current(),source=activeSource()||s.files[0];if(source)await showSource(source,s.active?.bucket||'files',s.active?.index||0);
      if(s.result&&s.previewVersion===s.version)renderResult(s.result);
    }
    async function buildPayload() {
      const s=current();if(state.rangeDirty)throw Error('请先应用表头与数据范围。');if(!s.files.length)throw Error('请先选择文件。');if(s.schemeProblems.length)throw Error('设置中有未能唯一对应的字段，请补选并确认。');
      if(mappingReviews().length)throw Error('以下表头已变化，请打开相应工作表核对字段对应：'+mappingReviews().join('、'));
      if(state.mode==='merge'){
        const selections=[];
        for(const source of s.files){if(!source.sheets.length)await loadSource(source);for(const sheet of source.sheets){if(source.excluded.includes(sheet.name)||(!s.options.hidden&&sheet.state!=='visible'))continue;const cfg=source.configs[sheet.name]||await loadSource(source,sheet.name);selections.push({path:source.path,sheet:sheet.name,selection_origin:cfg.selection_origin||'auto',...(nativeSource(source)?{selection:cfg.selection}:{}),...(Object.keys(cfg.field_map).length?{field_map:cfg.field_map}:{})});}}
        if(!selections.length)throw Error('请至少选择一张参与合并的工作表。');
        const template=s.files[s.options.templateIndex]||primary(),templateSheet=s.options.templateSheet||template.sheet;
        if(!selections.some(item=>item.path===template.path&&item.sheet===templateSheet))throw Error('所选模板工作表未参与合并，请勾选该页或更换模板。');
        for(const item of selections)if(item.field_map){const source=s.files.find(f=>f.path===item.path);item.field_map_context={source:fieldContext(source.configs[item.sheet].report.fields),target:fieldContext(schemeConfig().report.fields)};}
        return {paths:s.files.map(f=>f.path),options:engineOptions(),template:{path:template.path,sheet:templateSheet},selections};
      }
      const source=primary();if(!config(source))await loadSource(source);const cfg=config(source);
      const base={path:source.path,sheet:source.sheet,selection:cfg.selection,selection_origin:cfg.selection_origin||'auto'};
      if(state.mode==='split'){
        if(!s.group.length)throw Error('请点选至少一个拆分字段，或选择命名单元格 / 工作表名。');
        for(const f of s.filters)if(!f.field||!String(f.value??'').trim()||(f.op==='interval'&&!f.end_field))throw Error('请补全筛选字段、条件值，以及区间条件的截止字段。');
        return {...base,group_fields:s.group.map(g=>g.kind==='cell'?{kind:'cell',ref:g.ref}:clone(g)),group_mode:s.groupMode,blank_group:s.blankGroup,include_example_rows:!!s.options.exampleRows,retain_all_sheets:!!s.options.retainAll,...(s.outputFields.length?{output_fields:[...s.outputFields]}:{}),filters:clone(s.filters),retain_sheets:clone(s.retained),shared_label_ranges:clone(s.shared),resolutions:clone(s.resolutions),allow_warnings:false};
      }
      if(s.filteredEnabled){if(!s.outputFilters.length)throw Error('请为额外筛选结果至少添加一个条件。');for(const f of s.outputFilters)if(!f.field||!String(f.value??'').trim()||(f.op==='interval'&&!f.end_field))throw Error('请补全额外筛选结果的条件。');}
      if(!s.feedback.length)throw Error('请添加至少一份反馈文件。');if(!s.keys.length)throw Error('请选择用于对应同一条记录的字段。');if(!s.updates.length)throw Error('请选择允许回填的字段。');
      const feedback=[];
      for(const file of s.feedback){if(!config(file))await loadSource(file);const fc=config(file);feedback.push({path:file.path,sheet:file.sheet,selection:fc.selection,selection_origin:fc.selection_origin||'auto',field_map:{...autoMap(cfg.report.fields,fc.report.fields),...fc.field_map}});}
      let baseline;
      if(s.baseline){if(!config(s.baseline))await loadSource(s.baseline);const bc=config(s.baseline);baseline={path:s.baseline.path,sheet:s.baseline.sheet,selection:bc.selection,selection_origin:bc.selection_origin||'auto',field_map:{...autoMap(cfg.report.fields,bc.report.fields),...bc.field_map}};}
      return {...base,feedback,...(baseline?{baseline}:{}),key_fields:[...s.keys],update_fields:[...s.updates],compare_fields:[...s.compares],...(s.groupField!=null?{group_field:s.groupField}:{}),resolutions:clone(s.resolutions),...(s.filteredEnabled?{filtered_output:{...clone(s.filteredOptions),filters:clone(s.outputFilters),...(s.filteredFields.length?{output_fields:[...s.filteredFields]}:{})}}:{})};
    }
    async function previewResult() {
      if(state.busy||state.merging)return false;
      const mode=state.mode,s=current(),version=s.version,request=++state.operationRequest;
      if(s.schemeProblems.length){showSchemeProblems();return false;}
      state.busy=true;updateAction();notice();
      try{
        const payload=await buildPayload();
        if(version!==s.version||mode!==state.mode||request!==state.operationRequest)return false;
        const report=await call({merge:'mergePreview',split:'splitPreview',reconcile:'reconcilePreview'}[mode],payload);
        if(version!==s.version||mode!==state.mode||request!==state.operationRequest||state.disposed)return false;
        s.payload=clone(payload);if(report.source_fingerprint)s.payload.preview_fingerprint=report.source_fingerprint;
        s.result=report;s.previewVersion=version;renderResult(report);renderFiles();
        $('result-pane').scrollIntoView({behavior:'smooth',block:'start'});return true;
      }catch(error){if(!state.disposed&&mode===state.mode)fail(error);return false;}
      finally{state.busy=false;if(!state.disposed)updateAction();}
    }
    function applyIssueChoice(issue,choice) {
      const s=current();
      if(choice.patch?.filtered_output){const patch=choice.patch.filtered_output;for(const [key,value] of Object.entries(patch)){if(Array.isArray(value))s.filteredOptions[key]=[...new Set([...(s.filteredOptions[key]||[]),...value])];else if(value&&typeof value==='object')s.filteredOptions[key]={...(s.filteredOptions[key]||{}),...value};else s.filteredOptions[key]=value;}}
      else if(state.mode==='split'&&choice.patch){
        const patch=choice.patch;
        if(patch.retain_sheets)s.retained=[...new Set([...s.retained,...patch.retain_sheets])];
        if(patch.shared_label_ranges)s.shared=[...new Set([...s.shared,...patch.shared_label_ranges])];
        if(patch.resolutions)Object.assign(s.resolutions,patch.resolutions);
        if(patch.blank_group)s.blankGroup=patch.blank_group;
      }else s.resolutions[issue.id]=choice.value;
      invalidate('已记录处理方式，正在重新核对结果。',true);previewResult();
    }
    function renderResult(report) {
      $('result-pane').classList.remove('hidden');
      const groups=report.groups||[],changes=report.changes||[],issues=report.issues||[],totals=report.totals||{};
      if(state.mode==='merge'){
        const records=totals.rows_written??totals.records??totals.rows;
        const independent=(report.mode||current().payload?.options?.mode)==='workbook';
        const rowsKnown=groups.length>0&&groups.every(group=>group.row_count!=null);
        const detail=!independent&&records!=null?`${records} 条记录`:rowsKnown?`合计 ${groups.reduce((sum,group)=>sum+Number(group.row_count),0)} 行（含标题和表头）`:'';
        $('result-summary').textContent=`${groups.length} 张结果工作表${detail?' · '+detail:''}`;
      }else $('result-summary').textContent=state.mode==='reconcile'?`${totals.records??'—'} 条底表记录 · ${totals.changes??changes.length} 处变化 · ${totals.additions??changes.filter(c=>c.kind==='new').length} 条新增${totals.additions_skipped?` · ${totals.additions_skipped} 条新增未并入`:''} · ${totals.missing??report.missing?.length??0} 条未收到反馈`:`${groups.length} 个输出文件 · ${totals.records??totals.rows??groups.reduce((n,g)=>n+(Number(g.count)||0),0)} 条记录`;
      const warnings=[...(report.warnings||[])];if(report.retained_sheets?.length)warnings.push('同时保留依赖工作表：'+report.retained_sheets.map(s=>typeof s==='string'?s:s.name).join('、'));
      const messages=warningMessages(warnings),warningBox=$('result-warnings');warningBox.replaceChildren();warningBox.classList.toggle('hidden',!messages.length);
      // 默认折叠: 大文件的兼容提示可能几十条, 不折叠会把下方预览挤到很深
      if(messages.length){const summary=make('summary','',`兼容与处理提示 · ${messages.length} 项${report.can_export?' · 可继续导出':''}（点开查看）`),list=make('ul');for(const message of messages)list.append(make('li','',message));warningBox.append(summary,list);}
      const panel=$('issues-panel');panel.replaceChildren();panel.classList.toggle('hidden',!issues.length);
      if(issues.length)panel.append(make('b','',report.can_export?'可选处理与提示':'请集中处理以下问题'));
      for(const issue of issues){
        const row=make('div','issue'+(issue.resolved?' solved':'')),label=make('div','',issue.message||issue.id);if(issue.location)label.append(make('div','hint',text(issue.location)));if(issue.append_blocked)label.append(make('div','hint',issue.append_blocked));row.append(label);
        if(issue.choices?.length){const select=make('select');select.setAttribute('aria-label',issue.message||issue.id);setOptions(select,[['',issue.blocking===false?'按默认方式继续':'选择处理方式'],...issue.choices.map(c=>[c.value,c.label])],issue.resolution??current().resolutions[issue.id]??'');select.onchange=()=>{const choice=issue.choices.find(c=>String(c.value)===select.value);if(choice)applyIssueChoice(issue,choice);};row.append(select);}else row.append(make('span','hint',issue.resolved?'已处理':issue.blocking===false?'可按默认方式继续':'请调整上方范围或字段对应后重试'));
        if(issue.suggested_key){const sk=issue.suggested_key,fix=make('button','ghost key-fix','改用 '+(sk.title||('第 '+sk.field+' 列'))+' 作主键');fix.onclick=async()=>{fix.disabled=true;const s=current();s.keys=[sk.field];invalidate();renderRoles();renderToolbar();const c=config(activeSource());if(c)renderGrid(c.report);await previewResult();};row.append(fix);}
        panel.append(row);
      }
      $('result-samples').replaceChildren();$('result-changes').replaceChildren();
      if(state.mode==='reconcile'){renderChanges(changes);if(report.filtered_preview){$('result-samples').append(make('h3','','额外筛选结果'));renderSamples(report.filtered_preview.groups||[]);}}else renderSamples(groups);
      renderMissing(report.missing||[]);
      $('action-hint').textContent=report.can_export?'预览已准备好，可核对后导出。':'预览发现需要处理的问题。';updateAction();
    }
    function renderSamples(groups) {
      const box=$('result-samples');if(!groups.length){box.append(make('p','hint','没有符合条件的记录。'));return;}
      const tabs=make('div','sample-tabs'),content=make('div');
      const show=(group,index)=>{
        [...tabs.children].forEach((b,i)=>b.classList.toggle('active',i===index));content.replaceChildren();
        const rows=group.sample||[],grid=group.sample_kind==='grid';
        const headers=group.columns||(grid?Array.from({length:Math.max(0,...rows.map(row=>row.length||0))},(_,i)=>colName(i+1)):fieldList().map(f=>f.label));
        const summaryText=grid?`${group.row_count!=null?'工作表共 '+group.row_count+' 行':'工作表行数未提供'}，预览${(group.sample_start_row||1)===1?'前 '+rows.length+' 行':'第 '+group.sample_start_row+'–'+(group.sample_start_row+rows.length-1)+' 行'}`:`${group.count!=null?group.count+' 条记录 · ':''}${group.sources?.length?group.sources.length+' 个来源 · ':''}显示前 ${rows.length} 条样本`;
        const summary=make('p','hint',summaryText);content.append(summary);
        const scroll=make('div','sample-scroll'),table=make('table','sample-table'),head=make('tr');
        if(grid)head.append(make('th','','行'));
        const list=Array.isArray(headers)?headers:[];list.forEach((h,i)=>head.append(make('th','',typeof h==='object'?h.label||h.title||String(i+1):h)));table.append(head);
        rows.forEach((values,i)=>{const tr=make('tr');if(grid)tr.append(make('th','',Number(group.sample_start_row||1)+i));const cells=Array.isArray(values)?values:values.values||Object.values(values);cells.forEach((value,j)=>{const td=make('td','',value&&typeof value==='object'&&'value'in value?text(value.value):text(value));if(state.mode==='split'){const sourceRow=group.records?.[i],cfg=primaryConfig(),field=typeof list[j]==='object'?list[j]:cfg?.report.fields?.[j];const original=cfg?.report.cells?.find(c=>cfg.selection.axis==='rows'?c.row===sourceRow&&c.col===field?.index:c.col===sourceRow&&c.row===field?.index);applyCellStyle(td,original?.style);}tr.append(td);});table.append(tr);});
        if(!rows.length)content.append(make('p','hint','引擎未返回样本内容，请检查记录数量及范围。'));else{scroll.append(table);content.append(scroll);}
      };
      groups.forEach((group,index)=>{const quantity=group.count!=null?`${group.count} 条记录`:group.row_count!=null?`共 ${group.row_count} 行`:'';const button=make('button','',`${group.label||group.id||'结果 '+(index+1)}${quantity?' · '+quantity:''}`);button.onclick=()=>show(group,index);tabs.append(button);});box.append(tabs,content);show(groups[0],0);
    }
    function renderChanges(changes) {
      const box=$('result-changes');if(!changes.length){box.append(make('p','hint','当前选定范围内没有值变化。'));return;}
      const scroll=make('div','changes-scroll'),table=make('table','changes-table'),head=make('tr');for(const title of ['记录','字段','原值','新值','来源'])head.append(make('th','',title));table.append(head);
      for(const change of changes){const tr=make('tr');tr.append(make('td','',text(change.key)),make('td','',change.field_title||(change.field==null?'新增记录':fieldLabel(change.field))),make('td','',text(change.before)),make('td',change.marked?'changed-value':'',text(change.after)),make('td','',typeof change.source==='object'?`${basename(change.source.path||'')} / ${change.source.sheet||''}`:text(change.source)));table.append(tr);}
      scroll.append(table);box.append(scroll,make('p','hint','红色表示选定比较字段中的调整。导出时保留原表样式，并对确认的调整标红。'));
    }
    function renderMissing(missing) {
      const box=$('missing-records');box.replaceChildren();if(!missing.length)return;
      const details=make('details'),summary=make('summary','',`${missing.length} 条记录未收到反馈，保留现有内容 · 点击查看全部`),scroll=make('div','sample-scroll'),table=make('table','sample-table');
      const header=make('tr');header.append(make('th','','记录'),make('th','','底表位置'),make('th','','查看'));table.append(header);
      const axis=primaryConfig()?.selection.axis;
      for(const item of missing){const row=make('tr'),record=make('td','',text(item.key??item)),position=make('td','',item.position?`第 ${item.position} ${axis==='columns'?'列':'行'}`:'—'),action=make('td');
        if(item.position){const button=make('button','ghost compact','定位到底表');button.onclick=async()=>{await showSource(primary(),'files',0,primary().sheet,{row:axis==='rows'?Math.max(1,item.position-3):1,col:axis==='columns'?Math.max(1,item.position-3):1});$('workspace').scrollIntoView({behavior:'smooth',block:'start'});};action.append(button);}
        row.append(record,position,action);table.append(row);
      }
      scroll.append(table);details.append(summary,scroll);box.append(details);
    }
    async function exportResult(outOverride) {
      const s=current();if(!s.result||!s.result.can_export||s.previewVersion!==s.version||state.busy||state.merging){notice('请先生成可导出的最新预览。',true);return false;}
      const mode=state.mode,version=s.version,payload=clone(s.payload),_d=new Date(Date.now()+8*3600e3),_p=n=>String(n).padStart(2,'0'),stamp=`${_d.getUTCFullYear()}${_p(_d.getUTCMonth()+1)}${_p(_d.getUTCDate())}_${_p(_d.getUTCHours())}${_p(_d.getUTCMinutes())}`;
      let out=outOverride;
      if(!out)out=mode==='split'?await toolbox.selectOutputFolder():await toolbox.saveDialog(`${mode==='merge'?'合并结果':'回填结果'}_${stamp}.xlsx`);
      if(!out)return false;if(version!==s.version||mode!==state.mode){notice('设置已发生变化，请重新预览。',true);return false;}
      payload[mode==='split'?'out_dir':'out']=out;
      state.merging=true;openPane('progress-pane');$('progress-title').textContent='正在导出…';$('prog-file').textContent='';$('prog-count').textContent='';$('prog-bar').style.width='0%';
      state.offProgress=toolbox.onProgress?.(m=>{if(state.disposed)return;$('prog-file').textContent=m.current||'';$('prog-count').textContent=m.total?`${m.file||0} / ${m.total}`:'';$('prog-bar').style.width=`${Math.min(99,Math.round((m.file||0)/Math.max(1,m.total||1)*100))}%`;});
      try{
        const report=await call({merge:'merge',split:'split',reconcile:'reconcile'}[mode],payload);
        if(report.exported===false||report.can_export===false){s.result=report;s.payload=null;s.previewVersion=-1;renderResult(report);throw Error('文件尚未导出，请处理报告中的问题后重新预览。');}
        if(!state.disposed)showDone(report,out);return true;
      }catch(error){if(!state.disposed)fail(error);return false;}
      finally{state.offProgress?.();state.offProgress=null;state.merging=false;if(!state.disposed){closePane('progress-pane');updateAction();}}
    }
    function showDone(report,out) {
      openPane('done-pane');$('done-icon').textContent=report.cancelled?'—':'✓';$('done-title').textContent=report.cancelled?'处理已取消':'导出完成';
      const totals=report.totals||{};$('done-stats').textContent=state.mode==='split'?`${report.files?.length??report.outputs?.length??current().result?.groups?.length??0} 个文件`:`${totals.rows_written??totals.records??current().result?.totals?.records??''} 条记录${report.changes_path?' · 已生成变更记录':''}`;
      const location=report.output||report.out||out,outputs=report.outputs||(state.mode==='split'?report.files?.filter(f=>typeof f==='object'&&f.path).map(f=>f.path):null)||[location];
      $('done-out').textContent=outputs.map(text).join('\n');$('done-out').style.whiteSpace='pre-wrap';$('report-detail').textContent=JSON.stringify(report,null,2);$('report-detail').classList.add('hidden');
      const messages=warningMessages([...(current().result?.warnings||[]),...(report.warnings||[])]);
      $('done-warnings').classList.toggle('hidden',!messages.length);$('done-warnings').open=false;
      $('done-warnings-title').textContent=`兼容与处理提示 · ${messages.length} 项（点开查看）`;
      $('done-warnings-list').replaceChildren(...messages.map(message=>make('li','',message)));
      if(report.filtered_output_path)$('done-stats').textContent+=' · 已同时导出筛选结果';
      $('btn-open-folder').onclick=()=>toolbox.showInFolder(typeof location==='string'?location:out);
    }

    // Schemes store semantic header paths, never reusable absolute field numbers.
    function identity(field) { return field ? {signature:field.signature,...(field.match_signature!==undefined?{match_signature:field.match_signature}:{}),label:field.label,empty:!!field.empty} : {signature:null,label:'未知字段',empty:true}; }
    function uniqueField(saved,fields) { const signature=matchSignature(saved);if(!signature||saved.empty||String(signature).startsWith('@empty:'))return null;const matches=fields.filter(f=>(saved.match_signature!==undefined?matchSignature(f):f.signature)===signature&&!f.empty);return matches.length===1?matches[0]:null; }
    function captureRoles(fields=fieldList()) {
      const s=current(),lookup=index=>identity(fields.find(f=>f.index===Number(index)));
      return {outputFields:s.outputFields.map(lookup),filteredFields:s.filteredFields.map(lookup),keys:s.keys.map(lookup),updates:s.updates.map(lookup),compares:s.compares.map(lookup),group:s.group.map(g=>g.kind==='field'?{kind:'field',field:lookup(g.index)}:g.kind==='cell'?{...g,expected:g.expected??primaryConfig()?.report.cells?.find(c=>cellRef(c.row,c.col)===g.ref)?.value}:clone(g)),filters:s.filters.map(f=>({...f,field:lookup(f.field),end_field:f.end_field?lookup(f.end_field):null})),outputFilters:s.outputFilters.map(f=>({...f,field:lookup(f.field),end_field:f.end_field?lookup(f.end_field):null}))};
    }
    function showSchemeProblems() {
      const s=current();if(!s.schemeProblems.length)return;
      notice('以下内容无法唯一对应，已停止套用，请在原表上补选：\n'+s.schemeProblems.join('；'),true);
      const button=make('button','','已补选，使用当前字段');button.style.margin='10px 0 0';button.onclick=()=>{s.schemeProblems=[];invalidate();notice('已采用当前手动选择的字段，请重新预览核对。');};$('notice').append(make('br'),button);
      updateAction();
    }
    function restoreRoles(saved,fields,validateCells=true) {
      const s=current(),problems=[];
      const resolve=(field,role)=>{const match=uniqueField(field,fields);if(!match)problems.push(role+'：'+field.label);return match?.index;};
      for(const [key,label]of[['keys','认记录（主键）'],['updates','允许修改的列'],['compares','改后标红的列'],['outputFields','输出字段'],['filteredFields','额外筛选结果字段']])s[key]=(saved[key]||[]).map(f=>resolve(f,label)).filter(v=>v!==undefined);
      if(saved.groupField!=null){const g=resolve(saved.groupField,'单位列（分组）');s.groupField=g!==undefined?g:null;}
      s.group=[];
      for(const g of saved.group||[]){if(g.kind==='field'){const index=resolve(g.field,'拆分依据');if(index!==undefined)s.group.push({kind:'field',index});}else if(g.kind==='cell'){
        const cells=primaryConfig()?.report.cells||[],matches=cells.filter(c=>c.value!==null&&text(c.value)===text(g.expected));
        if(!validateCells)s.group.push({kind:'cell',ref:g.ref,expected:g.expected});else if(g.expected!==undefined&&matches.length===1)s.group.push({kind:'cell',ref:cellRef(matches[0].row,matches[0].col),expected:g.expected});else problems.push('命名单元格 '+g.ref+'（请重新点选）');
      }else s.group.push({kind:'sheet'});}
      s.filters=(saved.filters||[]).map(f=>({...f,field:resolve(f.field,'筛选字段')??null,end_field:f.end_field?resolve(f.end_field,'区间截止字段')??null:undefined}));
      s.outputFilters=(saved.outputFilters||[]).map(f=>({...f,field:resolve(f.field,'额外结果筛选字段')??null,end_field:f.end_field?resolve(f.end_field,'额外结果截止字段')??null:undefined}));
      s.schemeProblems=problems;if(problems.length)showSchemeProblems();return problems;
    }
    function readSchemes() {try{const value=JSON.parse(localStorage.getItem('excelToolSchemes.v1')||'[]');return Array.isArray(value)?value:[];}catch{return[];}}
    function writeSchemes(schemes){localStorage.setItem('excelToolSchemes.v1',JSON.stringify(schemes));}
    function renderSchemes(){const selected=$('scheme-select').value;setOptions($('scheme-select'),[['','选择已保存的方案'],...readSchemes().filter(s=>s.mode===state.mode).map(s=>[s.id,s.name])],selected);}
    function saveScheme(){
      const s=current(),cfg=schemeConfig(),name=$('scheme-name').value.trim();if(!name){notice('请先给方案起一个名字。',true);return;}if(!cfg||state.rangeDirty){notice('请先查看并确认底表或模板的表头与数据范围。',true);return;}if(s.schemeProblems.length){showSchemeProblems();return;}
      const mappings=[];for(const source of [...s.files,...s.feedback,...(s.baseline?[s.baseline]:[])])for(const [sheet,c]of Object.entries(source.configs)){if(!Object.keys(c.field_map).length)continue;const entries=Object.entries(c.field_map).map(([target,index])=>({target:identity(cfg.report.fields.find(f=>f.index===Number(target))),source:index==null?null:identity(c.report.fields.find(f=>f.index===index))}));mappings.push({signature:c.report.signature,fields:c.report.fields.map(identity),entries,sheet});}
      const existing=readSchemes(),old=existing.find(x=>x.name===name&&x.mode===state.mode),scheme={version:1,id:old?.id||String(Date.now()),name,mode:state.mode,structure:{axis:cfg.selection.axis,depth:cfg.selection.header_end-cfg.selection.header_start+1,fields:cfg.report.fields.map(identity)},roles:captureRoles(),options:clone(s.options),groupMode:s.groupMode,filteredEnabled:s.filteredEnabled,blankGroup:s.blankGroup,region:{...cfg.selection,data_gap:cfg.selection.data_start-cfg.selection.header_end-1,tail:Math.max(0,(cfg.selection.axis==='rows'?cfg.report.max_row:cfg.report.max_column)-cfg.selection.data_end)},mappings};
      writeSchemes([...existing.filter(x=>x.id!==scheme.id),scheme]);renderSchemes();$('scheme-select').value=scheme.id;notice('设置已保存。应用时会先核对表头结构，再套用字段设置。');
    }
    async function locateSchemeRegion(scheme) {
      const source=schemeSource(),sheet=schemeSheet(),cfg=schemeConfig(),saved=scheme.region;
      if(!saved)return cfg;
      const expected=scheme.structure.fields,axis=scheme.structure.axis,depth=scheme.structure.depth;
      const match=report=>{
        if(report.selection.axis!==axis||report.selection.header_end-report.selection.header_start+1!==depth)return null;
        const fields=expected.map(f=>uniqueField(f,report.fields));
        // Empty and repeated headings have no reliable relocation identity.
        if(fields.some(f=>!f))return null;
        return {start:Math.min(...fields.map(f=>f.index)),end:Math.max(...fields.map(f=>f.index))};
      };
      const maxRecord=axis==='rows'?cfg.report.max_row:cfg.report.max_column,maxField=axis==='rows'?cfg.report.max_column:cfg.report.max_row;
      const makeSelection=(hs,start=1,end=maxField)=>({axis,header_start:hs,header_end:hs+depth-1,data_start:hs+depth+(saved.data_gap||0),data_end:maxRecord-(saved.tail||0),field_start:start,field_end:end});
      let found=match(cfg.report),selected=found?makeSelection(cfg.selection.header_start,found.start,found.end):null;
      if(!selected){
        const starts=new Set([saved.header_start,cfg.selection.header_start]);
        const first=expected.find(f=>!f.empty&&f.signature&&!String(f.signature).startsWith('@empty:'));
        let leaf;try{const path=JSON.parse(first?.signature||'[]');leaf=path[path.length-1];}catch{}
        if(leaf!==undefined)for(const cell of cfg.report.cells||[])if(text(cell.value).trim()===text(leaf).trim())starts.add((axis==='rows'?cell.row:cell.col)-depth+1);
        const matches=[];
        for(const hs of [...starts].filter(n=>n>=1&&n+depth<=maxRecord+1).slice(0,12)){
          try{
            const report=await call('preview',{path:source.path,sheet,selection:makeSelection(hs),selection_origin:'manual',row_limit:1,col_limit:1});
            const located=match(report);if(located)matches.push(makeSelection(hs,located.start,located.end));
          }catch{/* Try another candidate; no input or workbook is changed. */}
        }
        if(matches.length!==1)return null;
        selected=matches[0];
      }
      return loadSource(source,sheet,{row:1,col:1},selected);
    }
    async function applyScheme(){
      const scheme=readSchemes().find(s=>s.id===$('scheme-select').value),s=current();let cfg=schemeConfig();if(!scheme){notice('请先选择已保存的设置。');return false;}if(!cfg){notice('请先查看底稿或模板并确认标题行范围。',true);return false;}
      if(state.rangeDirty){notice('请先应用当前范围，再核对方案。',true);return false;}
      invalidate();
      const mode=state.mode,version=s.version;
      cfg=await locateSchemeRegion(scheme);
      if(mode!==state.mode||version!==s.version||state.disposed)return false;
      if(!cfg){clearRoles();s.schemeProblems=['表头结构无法唯一对应，请圈选范围并补选字段'];showSchemeProblems();renderRoles();renderFilters();return false;}
      clearMappings();
      s.options={...s.options,...scheme.options,templateIndex:s.options.templateIndex,templateSheet:s.options.templateSheet};s.groupMode=scheme.groupMode||'combination';s.filteredEnabled=!!scheme.filteredEnabled;s.blankGroup=scheme.blankGroup||'separate';
      const problems=restoreRoles(scheme.roles,cfg.report.fields);
      for(const source of [...s.files,...s.feedback,...(s.baseline?[s.baseline]:[])]){
        if(!config(source))await loadSource(source);const c=config(source);
        const candidates=[...new Map((scheme.mappings||[]).filter(m=>m.fields.every(f=>uniqueField(f,c.report.fields))).map(m=>[JSON.stringify(m.entries),m])).values()];
        if(candidates.length===1){c.field_map={};for(const entry of candidates[0].entries){const target=uniqueField(entry.target,cfg.report.fields),mapped=entry.source?uniqueField(entry.source,c.report.fields):null;if(target&&(!entry.source||mapped))c.field_map[target.index]=mapped?.index??null;else problems.push('字段对应：'+entry.target.label);}}
        else if(candidates.length>1)problems.push(source.name+' 有多个可用的字段对应方案');
      }
      if(mode!==state.mode||version!==s.version||state.disposed)return false;
      s.schemeProblems=problems;renderOptions();renderWorkspace();await renderMapping();
      if(problems.length)showSchemeProblems();else notice('方案已按当前表头唯一对应，恢复范围并适配本表记录数。请预览结果。');return !problems.length;
    }

    // Event bindings are scoped to this mounted view and removed with it.
    $('btn-back').onclick=()=>{if(!state.merging)ctx.goHome();};
    root.querySelectorAll('[data-mode]').forEach(button=>button.onclick=()=>switchMode(button.dataset.mode));
    $('btn-add-files').onclick=async()=>addFromPaths(await toolbox.selectFiles());
    $('btn-add-folder').onclick=async()=>{const paths=await toolbox.selectFolder();await addFromPaths(Array.isArray(paths)?paths:paths?[paths]:[]);};
    $('btn-feedback').onclick=async()=>addFromPaths(await toolbox.selectFiles(),'feedback');
    $('btn-baseline').onclick=async()=>addFromPaths(await toolbox.selectFiles(),'baseline');
    $('baseline-name').onclick=()=>{if(current().baseline)showSource(current().baseline,'baseline',0);};
    $('btn-remove-baseline').onclick=()=>{current().baseline=null;invalidate();renderFiles();showSource(primary());};
    $('btn-clear').onclick=()=>{const s=current();clearRoles();s.files=[];s.active=null;state.gridRequest++;state.rangeDirty=false;invalidate();renderFiles();$('workspace').classList.add('hidden');};
    $('sheet-select').onchange=async()=>{const source=activeSource(),s=current();const oldFields=config(source)?.report.fields||[],saved=isPrimary(source)?captureRoles(oldFields):null;source.sheet=$('sheet-select').value;if(saved)clearMappings();invalidate();await showSource(source,s.active.bucket,s.active.index);if(saved&&config(source)){restoreRoles(saved,config(source).report.fields);renderRoles();renderFilters();}};
    $('sheet-enabled').onchange=()=>{const source=activeSource();source.excluded=source.excluded.filter(name=>name!==source.sheet);if(!$('sheet-enabled').checked)source.excluded.push(source.sheet);invalidate();recommend();};
    for(const id of ['header-range','data-range','axis-select'])$(id).addEventListener(id==='axis-select'?'change':'input',()=>{state.rangeDirty=true;invalidate();});
    $('pick-header').onclick=()=>startPick('header');$('pick-data').onclick=()=>startPick('data');$('pick-group-cell').onclick=()=>startPick('cell');$('apply-range').onclick=applyRanges;
    const page=(rowDelta,colDelta)=>{const source=activeSource(),cfg=config(source),s=current();if(!cfg)return;showSource(source,s.active.bucket,s.active.index,source.sheet,{row:Math.max(1,cfg.page.row+rowDelta),col:Math.max(1,cfg.page.col+colDelta)});};
    $('page-prev').onclick=()=>page(-80,0);$('page-next').onclick=()=>page(80,0);$('page-left').onclick=()=>page(0,-40);$('page-right').onclick=()=>page(0,40);
    $('jump-go').onclick=()=>{try{const destination=parseCell($('jump-cell').value),source=activeSource(),cfg=config(source),s=current();if(destination.row>cfg.report.max_row||destination.col>cfg.report.max_column)throw Error('这个坐标超出了工作表范围。');showSource(source,s.active.bucket,s.active.index,source.sheet,destination);}catch(error){fail(error);}};
    $('group-by-sheet').onchange=()=>{const s=current();s.group=s.group.filter(g=>g.kind!=='sheet');if($('group-by-sheet').checked)s.group.push({kind:'sheet'});invalidate();renderRoles();};$('group-mode').onchange=()=>{current().groupMode=$('group-mode').value;invalidate();};$('blank-group').onchange=()=>{current().blankGroup=$('blank-group').value;invalidate();};$('split-example-rows').onchange=()=>{current().options.exampleRows=$('split-example-rows').checked;invalidate();};$('split-retain-all').onchange=()=>{current().options.retainAll=$('split-retain-all').checked;invalidate();};
    $('add-filter').onclick=()=>{current().filters.push({field:null,op:'equals',value:''});invalidate();renderFilters();};
    $('add-output-filter').onclick=()=>{current().outputFilters.push({field:null,op:'equals',value:''});invalidate();renderFilters();};$('filtered-enabled').onchange=()=>{current().filteredEnabled=$('filtered-enabled').checked;if(!current().filteredEnabled&&current().role==='filteredFields')current().role='keys';invalidate();renderFilters();renderRoles();};
    root.querySelectorAll('[name="merge-output"]').forEach(input=>input.onchange=()=>{current().options.mode=input.value;invalidate();recommend();});
    for(const [id,key]of [['opt-shape','shape'],['opt-align','align'],['opt-keywords','keywords']])$(id).addEventListener(id==='opt-keywords'?'input':'change',()=>{current().options[key]=$(id).value;invalidate();});
      $('opt-fill-ratio').addEventListener('input',()=>{const v=Number($('opt-fill-ratio').value);current().options.minFillRatio=Number.isFinite(v)&&v>=0?v:20;invalidate();});
    for(const [id,key]of [['opt-hidden','hidden'],['opt-source','source'],['opt-summary','removeSummary'],['opt-serial-strip','stripSerial'],['opt-serial-num','renumber'],['opt-title','title']])$(id).onchange=()=>{current().options[key]=$(id).checked;invalidate();recommend();};
    $('opt-template').onchange=()=>{current().options.templateIndex=Number($('opt-template').value);current().options.templateSheet=null;renderFiles();clearMappings();invalidate();renderMapping();};
    $('opt-template-sheet').onchange=()=>{current().options.templateSheet=$('opt-template-sheet').value;clearMappings();invalidate();renderMapping();};
    $('mapping-reviewed').onclick=()=>{const cfg=config(activeSource());if(!cfg)return;cfg.mapping_review=false;invalidate('已采用当前字段对应，请重新预览核对。');renderMapping();};
    document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('done-pane').classList.contains('hidden')&&!state.busy&&!state.merging)closePane('done-pane');});
    $('zoom-out').onclick=()=>{gridZoom.mode='manual';gridZoom.scale=Math.max(GRID_MIN_ZOOM,(gridZoom.scale||1)-0.1);applyGridZoom();};
    $('zoom-in').onclick=()=>{gridZoom.mode='manual';gridZoom.scale=Math.min(2,(gridZoom.scale||1)+0.1);applyGridZoom();};
    $('zoom-fit').onclick=()=>{gridZoom.mode='fit';applyGridZoom();};
    let zoomResizeTimer=null;
    window.addEventListener('resize',()=>{if(gridZoom.mode!=='fit')return;clearTimeout(zoomResizeTimer);zoomResizeTimer=setTimeout(applyGridZoom,120);});
    $('btn-preview').onclick=previewResult;$('btn-export').onclick=()=>exportResult();$('btn-cancel').onclick=()=>{toolbox.cancel?.();$('progress-title').textContent='正在停止…';};$('btn-again').onclick=()=>closePane('done-pane');$('btn-report').onclick=()=>$('report-detail').classList.toggle('hidden');
    $('scheme-save').onclick=saveScheme;$('scheme-apply').onclick=applyScheme;$('scheme-delete').onclick=()=>{const id=$('scheme-select').value;if(!id)return;writeSchemes(readSchemes().filter(s=>s.id!==id));renderSchemes();notice('已删除所选方案。');};
    const drop=$('dropzone');drop.ondragover=e=>{if([...e.dataTransfer.types].includes('Files')){e.preventDefault();drop.classList.add('dragover');}};drop.ondragleave=()=>drop.classList.remove('dragover');drop.ondrop=async e=>{if(!e.dataTransfer.files.length)return;e.preventDefault();drop.classList.remove('dragover');await addFromPaths([...e.dataTransfer.files].map(file=>toolbox.pathForFile(file)).filter(Boolean));};
    dispose=()=>{state.disposed=true;state.gridRequest++;state.operationRequest++;state.offProgress?.();delete window.__emTest;delete window.__excelTest;};
    window.__excelTest={state,sessions,current,addPaths,addFromPaths,switchMode,showSource,applyRanges,selectField,previewResult,exportResult,buildPayload,engineOptions,saveScheme,applyScheme,renderFiles,renderWorkspace,renderRoles,renderFilters,invalidate,get source(){return activeSource();},get config(){return config(activeSource());}};
    window.__emTest={state,addPaths,startMerge:async out=>{if(!current().result&&!(await previewResult()))return false;return exportResult(out);}};
    renderOptions();updateAction();
  };
})();
