// 工具箱壳 · 路由 / manifest 注册 / 全局浮层 / 主题 / toast
/* global toolbox */
(() => {
  const stage = document.getElementById('stage');
  const overlay = document.getElementById('overlay');
  const overlayPanel = document.getElementById('overlay-panel');
  const toasts = document.getElementById('toasts');
  let manifests = [];
  let currentTool = null;
  let unmountCurrent = null;

  // ---------------- 主题（跟随系统 + 手动覆盖 + 减少动效） ----------------
  const savedTheme = localStorage.getItem('theme') || 'auto';
  const reduceMotion = localStorage.getItem('reduceMotion') === '1';
  const applyTheme = () => {
    const dark = savedTheme === 'dark' ||
      (savedTheme === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
    // 同步窗口原生底色: 无边框窗口缩放时露出的区域不再闪异色
    toolbox.setBackground?.(dark ? '#0f1420' : '#eef1f7');
  };
  applyTheme();
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
  if (reduceMotion) document.documentElement.classList.add('reduce-motion');

  // 拉伸期间临时降级材质：暂停动画/隐藏极光/摘掉玻璃模糊，
  // 让合成器跟上 DWM 的窗口尺寸变化（消除掉帧露出的黑闪）
  let resizeTimer = null;
  window.addEventListener('resize', () => {
    document.documentElement.classList.add('resizing');
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      document.documentElement.classList.remove('resizing');
    }, 160);
  }, { passive: true });

  // ---------------- 窗口控制 ----------------
  document.getElementById('win-min').onclick = () => toolbox.winControl('min');
  document.getElementById('win-max').onclick = () => toolbox.winControl('max');
  document.getElementById('win-close').onclick = () => toolbox.winControl('close');

  // ---------------- toast ----------------
  window.tbToast = (msg, kind = '') => {
    const el = document.createElement('div');
    el.className = `toast glass-2 ${kind}`;
    el.textContent = msg;
    toasts.appendChild(el);
    setTimeout(() => {
      el.style.transition = 'opacity .2s, transform .2s';
      el.style.opacity = '0';
      el.style.transform = 'translateY(8px)';
      setTimeout(() => el.remove(), 220);
    }, 4000);
  };

  // ---------------- 首页 ----------------
  function renderHome() {
    unmountCurrent?.();
    unmountCurrent = null;
    currentTool = null;
    stage.innerHTML = '';
    const home = document.createElement('div');
    home.className = 'home';
    home.innerHTML = `
      <div class="hero">
        <h1>工具箱</h1>
        <p>本地运行 · 完全离线 · 文件不出本机</p>
      </div>
      <div class="tool-grid" id="tool-grid"></div>`;
    stage.appendChild(home);
    const grid = home.querySelector('#tool-grid');
    manifests.forEach((m, i) => {
      const card = document.createElement('div');
      card.className = 'tool-card glass-2 anim-rise';
      card.style.setProperty('--i', i);
      card.innerHTML = `<div class="icon">${m.icon}</div>
        <div class="name">${m.name}</div>
        <div class="desc">${m.description}</div>`;
      card.onclick = () => enterTool(m);
      grid.appendChild(card);
    });
    stage.scrollTop = 0;
  }

  // ---------------- 进入工具（fetch 视图 + 动态挂载脚本/样式） ----------------
  async function enterTool(m) {
    stage.innerHTML = '';
    const view = document.createElement('div');
    view.className = 'view-enter';
    const [html] = await Promise.all([fetch(m.view).then((r) => r.text())]);
    view.innerHTML = html;
    stage.appendChild(view);

    const css = document.createElement('link');
    css.rel = 'stylesheet';
    css.href = m.css;
    document.head.appendChild(css);

    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = m.script;
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });
    currentTool = m;
    unmountCurrent = window.toolUnmount || null;
    window.toolMount?.(view, { goHome: renderHome });
  }

  // ---------------- manifest 加载 ----------------
  async function loadManifests() {
    // v0.1: 清单列表内置（编译期接入；后续读取 manifests 目录）
    manifests = [
      {
        id: 'excel-merge', name: 'Excel 工具', icon: '▦',
        description: '合并 · 拆分 · 对照回填，点选原表即可处理',
        view: '../tools/excel-merge/index.html',
        script: '../tools/excel-merge/tool.js',
        css: '../tools/excel-merge/tool.css',
      },
    ];
  }

  // ---------------- 设置中心 / 关于 ----------------
  function openSettings() {
    overlayPanel.innerHTML = `
      <h2>设置中心</h2>
      <div class="sub">全局设置（对各工具生效）</div>
      <div class="kv"><span>外观</span>
        <span><select id="set-theme">
          <option value="auto">跟随系统</option><option value="light">浅色</option><option value="dark">深色</option>
        </select></span></div>
      <div class="kv"><span>减少动效（含材质降级）</span><span><span class="switch" id="set-motion"></span></span></div>
      <div class="sub" style="margin-top:16px">各工具的专属选项会按当前操作显示在工具页内。</div>`;
    overlay.classList.remove('hidden');
    const sel = overlayPanel.querySelector('#set-theme');
    sel.value = savedTheme;
    sel.onchange = () => {
      localStorage.setItem('theme', sel.value);
      location.reload();
    };
    const sw = overlayPanel.querySelector('#set-motion');
    sw.classList.toggle('on', reduceMotion);
    sw.onclick = () => {
      localStorage.setItem('reduceMotion', sw.classList.contains('on') ? '0' : '1');
      location.reload();
    };
  }

  async function openAbout() {
    const info = await toolbox.info();
    overlayPanel.innerHTML = `
      <h2>关于</h2>
      <div class="sub">团队本地工具箱</div>
      <div class="kv"><span>工具箱版本</span><span>v${info.appVersion}</span></div>
      <div class="kv"><span>Excel 处理引擎</span><span>M1 (sidecar)</span></div>
      <div class="kv"><span>平台</span><span>${info.platform} · Electron ${info.electron}</span></div>
      <div class="sub" style="margin-top:14px">完全离线运行，文件不出本机。检查更新请访问内部渠道。</div>`;
    overlay.classList.remove('hidden');
  }

  overlay.onclick = (e) => { if (e.target === overlay) overlay.classList.add('hidden'); };
  document.getElementById('btn-settings').onclick = openSettings;
  document.getElementById('btn-about').onclick = openAbout;

  // ---------------- 启动 ----------------
  loadManifests().then(renderHome);

  // GPU 软件渲染检测 → 自动材质降级（UI设计说明 §7：停用模糊与极光漂移）
  toolbox.info().then((info) => {
    if (info.gpuCompositing === 'disabled_software') {
      document.documentElement.classList.add('reduce-motion');
      tbToast('检测到软件渲染，已自动降级动效与材质以保证流畅', 'warn');
    }
  });
})();
