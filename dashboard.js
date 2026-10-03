const AUTH_KEY = 'csv_dashboard_auth';
const AUTH_VERSION = 'v1';

function readAuthState() {
  try {
    const raw = sessionStorage.getItem(AUTH_KEY);
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (parsed && parsed.version === AUTH_VERSION && parsed.value === '1' && typeof parsed.issuedAt === 'number') {
      return parsed;
    }

    return null;
  } catch (error) {
    return null;
  }
}

function clearAuthState() {
  sessionStorage.removeItem(AUTH_KEY);
}

if (!readAuthState()) {
  window.location.href = '/login';
} else {
  const $ = (s) => document.querySelector(s);
  const PAL = ['#0b7f86', '#e0892b', '#5b6fd6', '#c4465b', '#6aa84f', '#9b6dc9', '#2f9fd0', '#8a8f98'];
  const S = { rows: [], cols: [], wb: null, name: '', view: [], page: 0, sort: null };
  const charts = {};

  $('#logoutBtn').onclick = () => {
    clearAuthState();
    window.location.href = '/login';
  };

  const fmt = (n) => n == null || isNaN(n) ? '–' : Math.abs(n) >= 1e4 ? new Intl.NumberFormat('en', { notation: 'compact', maximumFractionDigits: 2 }).format(n) : new Intl.NumberFormat('en', { maximumFractionDigits: 2 }).format(n);
  const esc = (s) => String(s).replace(/[&<>\"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;' }[c]));
  const isNil = (v) => v === null || v === undefined || v === '' || (typeof v === 'number' && isNaN(v));
  const dateRe = /^\d{4}-\d{1,2}-\d{1,2}|^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/;
  const toNum = (v) => typeof v === 'number' ? v : parseFloat(String(v).trim().replace(/[,$€£₹%\s]/g, ''));

  function load(file) {
    $('#err').textContent = '';
    const r = new FileReader();
    r.onload = (e) => {
      try {
        S.wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        S.name = file.name;
        S.buf = e.target.result;
        S.src = 'upload';
        openSheet(S.wb.SheetNames[0]);
        afterLoad();
      } catch (x) {
        $('#err').textContent = 'Could not read this file. Check that it is a valid CSV or Excel file.';
      }
    };
    r.readAsArrayBuffer(file);
  }

  function openSheet(name) {
    const rows = XLSX.utils.sheet_to_json(S.wb.Sheets[name], { defval: null, raw: true });
    if (!rows.length) {
      $('#err').textContent = 'This sheet has no rows. Make sure the first row holds column headers.';
      return;
    }
    S.rows = rows;
    buildModel();
    const names = S.wb.SheetNames;
    $('#shwrap').classList.toggle('hide', names.length < 2);
    $('#sheet').innerHTML = names.map((n) => `<option ${n === name ? 'selected' : ''}>${esc(n)}</option>`).join('');
    $('#drop').classList.add('hide');
    $('#dash').classList.remove('hide');
    $('#fname').textContent = S.name;
    render();
  }

  function buildModel() {
    const names = [...new Set(S.rows.flatMap((r) => Object.keys(r)))];
    S.cols = names.map((n) => {
      const vals = S.rows.map((r) => r[n]).filter((v) => !isNil(v));
      const smp = vals.slice(0, 500);
      let nu = 0;
      let dt = 0;
      smp.forEach((v) => {
        if (v instanceof Date) dt++;
        else if (typeof v === 'number') nu++;
        else if (typeof v === 'string') {
          const t = v.trim();
          if (t && !isNaN(toNum(t)) && /^[-+$€£₹\d.,\s%()]+$/.test(t)) nu++;
          else if (dateRe.test(t) && !isNaN(Date.parse(t))) dt++;
        }
      });
      const k = smp.length || 1;
      const uniq = new Set(vals.map(String)).size;
      let type = nu / k >= .9 ? 'number' : dt / k >= .9 ? 'date' : (uniq <= Math.max(30, vals.length * .05) ? 'category' : 'text');
      if (!vals.length) type = 'text';
      S.rows.forEach((r) => {
        const v = r[n];
        if (isNil(v)) {
          r[n] = null;
          return;
        }
        if (type === 'number') r[n] = toNum(v);
        else if (type === 'date') r[n] = v instanceof Date ? v : new Date(v);
      });
      return { name: n, type, missing: S.rows.length - vals.length, uniq };
    });
  }

  const cols = (t) => S.cols.filter((c) => t.includes(c.type));
  const opts = (arr, sel) => arr.map((c) => `<option ${c.name === sel ? 'selected' : ''}>${esc(c.name)}</option>`).join('');

  function render() {
    const nums = cols(['number']);
    const cats = cols(['category']);
    const dts = cols(['date']);
    const n = S.rows.length;
    const totalMiss = S.cols.reduce((a, c) => a + c.missing, 0);
    $('#fmeta').textContent = `${n.toLocaleString()} rows · ${S.cols.length} columns`;
    const k = [['Rows', n.toLocaleString()], ['Columns', S.cols.length], ['Missing values', (100 * totalMiss / (n * S.cols.length)).toFixed(1) + '%']];
    nums.slice(0, 3).forEach((c) => {
      const v = S.rows.map((r) => r[c.name]).filter((x) => x != null);
      k.push([c.name + ' (total)', fmt(v.reduce((a, b) => a + b, 0))], [c.name + ' (average)', fmt(v.reduce((a, b) => a + b, 0) / (v.length || 1))]);
    });
    $('#kpis').innerHTML = k.slice(0, 9).map(([l, v]) => `<div class="kpi"><b>${esc(v)}</b><span>${esc(l)}</span></div>`).join('');

    const groupable = [...dts, ...cats, ...cols(['text']).filter((c) => c.uniq <= n * .8)];
    $('#bx').innerHTML = opts(groupable.length ? groupable : S.cols, (dts[0] || cats[0] || {}).name);
    $('#by').innerHTML = '<option value="">(row count)</option>' + opts(nums);
    $('#hx').innerHTML = opts(nums);
    $('#cx').innerHTML = opts(cats.length ? cats : S.cols.filter((c) => c.type !== 'number'), (cats[0] || {}).name);
    $('#sx').innerHTML = opts(nums, (nums[0] || {}).name);
    $('#sy').innerHTML = opts(nums, (nums[1] || nums[0] || {}).name);
    $('#corrSel').innerHTML = opts(nums, (nums[0] || {}).name);
    if (!$('#by').value && nums[0]) $('#by').value = nums[0].name;
    drawAll();
    profile();
    S.view = S.rows;
    S.page = 0;
    S.sort = null;
    $('#q').value = '';
    table();
  }

  function drawAll() {
    builder();
    hist();
    breakdown();
    scatter();
    corr();
  }

  function mk(id, cfg) {
    const st = getComputedStyle(document.documentElement);
    Chart.defaults.color = st.getPropertyValue('--muted').trim();
    Chart.defaults.borderColor = st.getPropertyValue('--line').trim();
    Chart.defaults.font.family = '"Schibsted Grotesk",system-ui,sans-serif';
    charts[id]?.destroy();
    cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false }, cfg.options || {});
    charts[id] = new Chart($('#' + id), cfg);
  }

  function empty(id, msg) {
    charts[id]?.destroy();
    const c = $('#' + id);
    const x = c.getContext('2d');
    x.clearRect(0, 0, c.width, c.height);
    x.fillStyle = getComputedStyle(document.body).color;
    x.font = '14px sans-serif';
    x.fillText(msg, 12, 28);
  }

  const colOf = (n) => S.cols.find((c) => c.name === n);

  function builder() {
    const xn = $('#bx').value;
    const yn = $('#by').value;
    const ag = yn ? $('#bag').value : 'count';
    const type = $('#bt').value;
    const xc = colOf(xn);
    if (!xc) return empty('c1', 'No columns to group by.');
    const dates = xc.type === 'date';
    let key = (v) => String(v);
    if (dates) {
      const ds = S.rows.map((r) => r[xn]).filter(Boolean);
      const span = (Math.max(...ds) - Math.min(...ds)) / 864e5;
      key = (v) => {
        const i = v.toISOString();
        return span > 120 ? i.slice(0, 7) : i.slice(0, 10);
      };
    }
    const g = new Map();
    S.rows.forEach((r) => {
      const x = r[xn];
      if (x == null) return;
      const k = key(x);
      const y = yn ? r[yn] : 1;
      if (yn && y == null) return;
      (g.get(k) || g.set(k, []).get(k)).push(y);
    });
    let data = [...g].map(([k, a]) => [k, ag === 'count' ? a.length : ag === 'sum' ? a.reduce((p, q) => p + q, 0) : ag === 'avg' ? a.reduce((p, q) => p + q, 0) / a.length : ag === 'min' ? Math.min(...a) : Math.max(...a)]);
    if (dates) data.sort((a, b) => a[0] < b[0] ? -1 : 1);
    else {
      data.sort((a, b) => b[1] - a[1]);
      data = data.slice(0, type === 'doughnut' ? 8 : 20);
    }
    mk('c1', {
      type: data.length ? type : 'bar',
      data: {
        labels: data.map((d) => d[0]),
        datasets: [{ label: yn ? ag + ' of ' + yn : 'Row count', data: data.map((d) => d[1]), backgroundColor: type === 'doughnut' ? PAL : PAL[0] + 'cc', borderColor: PAL[0], borderWidth: type === 'line' ? 2 : 0, tension: .25, pointRadius: type === 'line' ? 2 : 0, fill: false }]
      },
      options: { plugins: { legend: { display: type === 'doughnut' } }, scales: type === 'doughnut' ? {} : { y: { beginAtZero: true } } }
    });
  }

  function hist() {
    const n = $('#hx').value;
    if (!n) return empty('c2', 'No numeric columns found.');
    const v = S.rows.map((r) => r[n]).filter((x) => x != null);
    if (!v.length) return empty('c2', 'No data.');
    const mn = Math.min(...v);
    const mx = Math.max(...v);
    const b = Math.min(20, Math.max(5, Math.ceil(Math.sqrt(v.length))));
    const w = (mx - mn) / b || 1;
    const cnt = Array(b).fill(0);
    v.forEach((x) => cnt[Math.min(b - 1, Math.floor((x - mn) / w))]++);
    mk('c2', { type: 'bar', data: { labels: cnt.map((_, i) => fmt(mn + i * w) + '–' + fmt(mn + (i + 1) * w)), datasets: [{ label: 'Rows', data: cnt, backgroundColor: PAL[2] + 'cc', barPercentage: 1, categoryPercentage: .95 }] }, options: { plugins: { legend: { display: false } } } });
  }

  function breakdown() {
    const n = $('#cx').value;
    if (!n) return empty('c3', 'No category columns found.');
    const m = {};
    S.rows.forEach((r) => {
      const v = r[n];
      if (v != null) {
        const k = v instanceof Date ? v.toISOString().slice(0, 10) : String(v);
        m[k] = (m[k] || 0) + 1;
      }
    });
    let e = Object.entries(m).sort((a, b) => b[1] - a[1]);
    const top = e.slice(0, 7);
    const rest = e.slice(7).reduce((a, b) => a + b[1], 0);
    if (rest) top.push(['Other', rest]);
    mk('c3', { type: 'doughnut', data: { labels: top.map((d) => d[0]), datasets: [{ data: top.map((d) => d[1]), backgroundColor: PAL, borderWidth: 0 }] }, options: { plugins: { legend: { position: 'right' } } } });
  }

  function scatter() {
    const x = $('#sx').value;
    const y = $('#sy').value;
    if (!x || !y) return empty('c4', 'Needs two numeric columns.');
    const pts = S.rows.filter((r) => r[x] != null && r[y] != null).slice(0, 3000).map((r) => ({ x: r[x], y: r[y] }));
    mk('c4', { type: 'scatter', data: { datasets: [{ label: `${y} vs ${x}`, data: pts, backgroundColor: PAL[1] + '99', pointRadius: 3 }] }, options: { plugins: { legend: { display: false } }, scales: { x: { title: { display: true, text: x } }, y: { title: { display: true, text: y } } } } });
  }

  function pearson(a, b) {
    const p = S.rows.filter((r) => r[a] != null && r[b] != null);
    if (p.length < 3) return NaN;
    const xs = p.map((r) => r[a]);
    const ys = p.map((r) => r[b]);
    const mx = xs.reduce((s, v) => s + v, 0) / xs.length;
    const my = ys.reduce((s, v) => s + v, 0) / ys.length;
    let nu = 0;
    let dx = 0;
    let dy = 0;
    xs.forEach((v, i) => {
      nu += (v - mx) * (ys[i] - my);
      dx += (v - mx) ** 2;
      dy += (ys[i] - my) ** 2;
    });
    return nu / Math.sqrt(dx * dy);
  }

  function corr() {
    const t = $('#corrSel').value;
    $('#corrLabel').textContent = t || '…';
    const o = cols(['number']).filter((c) => c.name !== t).map((c) => [c.name, pearson(t, c.name)]).filter((d) => !isNaN(d[1])).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1])).slice(0, 12);
    if (!o.length) return empty('c5', 'Needs at least two numeric columns.');
    mk('c5', { type: 'bar', data: { labels: o.map((d) => d[0]), datasets: [{ data: o.map((d) => +d[1].toFixed(3)), backgroundColor: o.map((d) => d[1] >= 0 ? PAL[0] : PAL[3]) }] }, options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { min: -1, max: 1 } } } });
  }

  function profile() {
    const n = S.rows.length;
    $('#prof').innerHTML = '<tr><th>Column</th><th>Type</th><th>Missing</th><th>Unique</th><th>Min</th><th>Median</th><th>Mean</th><th>Max</th></tr>' + S.cols.map((c) => {
      let st = ['', '', '', ''];
      if (c.type === 'number') {
        const v = S.rows.map((r) => r[c.name]).filter((x) => x != null).sort((a, b) => a - b);
        if (v.length) st = [v[0], v[v.length >> 1], v.reduce((a, b) => a + b, 0) / v.length, v[v.length - 1]].map(fmt);
      } else if (c.type === 'date') {
        const v = S.rows.map((r) => r[c.name]).filter(Boolean).sort((a, b) => a - b);
        if (v.length) st = [v[0].toISOString().slice(0, 10), '', '', v[v.length - 1].toISOString().slice(0, 10)];
      }
      const pct = 100 * c.missing / n;
      return `<tr><td>${esc(c.name)}</td><td><span class="tag">${c.type}</span></td><td><span class="miss" style="width:${Math.max(2, pct * .6)}px"></span>${pct.toFixed(1)}%</td><td class="n">${c.uniq.toLocaleString()}</td><td class="n">${esc(st[0])}</td><td class="n">${esc(st[1])}</td><td class="n">${esc(st[2])}</td><td class="n">${esc(st[3])}</td></tr>`;
    }).join('');
  }

  const PS = 15;
  const cell = (v) => v == null ? '' : v instanceof Date ? v.toISOString().slice(0, 10) : typeof v === 'number' ? fmt(v) : v;

  function table() {
    const pages = Math.max(1, Math.ceil(S.view.length / PS));
    const s = S.page * PS;
    $('#tbl').innerHTML = '<tr>' + S.cols.map((c) => `<th class="s" data-c="${esc(c.name)}">${esc(c.name)}${S.sort && S.sort[0] === c.name ? (S.sort[1] > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('') + '</tr>' + S.view.slice(s, s + PS).map((r) => '<tr>' + S.cols.map((c) => `<td class="${c.type === 'number' ? 'n' : ''}">${esc(cell(r[c.name]))}</td>`).join('') + '</tr>').join('');
    $('#pinfo').textContent = `Page ${S.page + 1} of ${pages} · ${S.view.length.toLocaleString()} rows`;
    $('#prev').disabled = S.page <= 0;
    $('#next').disabled = S.page >= pages - 1;
  }

  $('#pick').onclick = () => $('#file').click();
  $('#file').onchange = (e) => e.target.files[0] && load(e.target.files[0]);
  const dz = $('#drop');
  ['dragenter', 'dragover'].forEach((t) => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', (e) => e.dataTransfer.files[0] && load(e.dataTransfer.files[0]));
  $('#reset').onclick = () => { $('#dash').classList.add('hide'); $('#drop').classList.remove('hide'); $('#file').value = ''; };
  $('#sheet').onchange = (e) => openSheet(e.target.value);
  ['bx', 'by', 'bag', 'bt'].forEach((i) => $('#' + i).onchange = builder);
  $('#hx').onchange = hist;
  $('#cx').onchange = breakdown;
  $('#sx').onchange = scatter;
  $('#sy').onchange = scatter;
  $('#corrSel').onchange = corr;
  $('#q').oninput = (e) => {
    const q = e.target.value.toLowerCase();
    S.view = q ? S.rows.filter((r) => S.cols.some((c) => String(cell(r[c.name])).toLowerCase().includes(q))) : S.rows;
    S.page = 0;
    table();
  };
  $('#prev').onclick = () => { S.page--; table(); };
  $('#next').onclick = () => { S.page++; table(); };
  $('#tbl').onclick = (e) => {
    const th = e.target.closest('th.s');
    if (!th) return;
    const c = th.dataset.c;
    const d = S.sort && S.sort[0] === c ? -S.sort[1] : 1;
    S.sort = [c, d];
    S.view = [...S.view].sort((a, b) => {
      const x = a[c];
      const y = b[c];
      if (x == null) return 1;
      if (y == null) return -1;
      return (x > y ? 1 : x < y ? -1 : 0) * d;
    });
    S.page = 0;
    table();
  };
  function updateThemeButton() {
    const r = document.documentElement;
    const dark = r.dataset.theme ? r.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme:dark)').matches;
    const btn = $('#theme');
    btn.textContent = dark ? '☀' : '☾';
    btn.setAttribute('aria-label', dark ? 'Switch to light theme' : 'Switch to dark theme');
    btn.title = dark ? 'Switch to light theme' : 'Switch to dark theme';
  }

  $('#theme').onclick = () => {
    const r = document.documentElement;
    const dark = r.dataset.theme ? r.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme:dark)').matches;
    r.dataset.theme = dark ? 'light' : 'dark';
    updateThemeButton();
    if (S.rows.length && !$('#dash').classList.contains('hide')) drawAll();
  };

  updateThemeButton();

  const CFG_KEY = 'sheetlens_cfg';
  function cfg() {
    let c = {};
    try { c = JSON.parse(localStorage.getItem(CFG_KEY) || '{}'); } catch (e) {}
    const h = location.hostname;
    if (!c.owner && h.endsWith('.github.io')) {
      c.owner = h.split('.')[0];
      c.repo = c.repo || location.pathname.split('/')[1] || '';
    }
    c.branch = c.branch || 'main';
    return c;
  }

  async function gh(path, o = {}) {
    const c = cfg();
    if (!c.owner || !c.repo) throw new Error('Set your repository in Settings first.');
    const h = { Accept: o.raw ? 'application/vnd.github.raw+json' : 'application/vnd.github+json' };
    if (c.token) h.Authorization = 'Bearer ' + c.token;
    const r = await fetch(`https://api.github.com/repos/${c.owner}/${c.repo}/contents/${path}${o.ref ? '?ref=' + encodeURIComponent(c.branch) : ''}`, { method: o.method || 'GET', headers: h, body: o.body ? JSON.stringify(o.body) : undefined });
    if (!r.ok) {
      let m = 'Request failed (' + r.status + ')';
      try { m = (await r.json()).message || m; } catch (e) {}
      const e = new Error(m);
      e.status = r.status;
      throw e;
    }
    return o.raw ? r.arrayBuffer() : r.json();
  }

  function b64(buf) {
    const u = new Uint8Array(buf);
    let s = '';
    for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  }

  const fpath = (n) => 'data/' + encodeURIComponent(n);

  async function saveToRepo(name, buf) {
    const c = cfg();
    let sha;
    try { sha = (await gh(fpath(name), { ref: 1 })).sha; } catch (e) { if (e.status !== 404) throw e; }
    await gh(fpath(name), { method: 'PUT', body: { message: (sha ? 'Update ' : 'Add ') + name, content: b64(buf), branch: c.branch, sha } });
  }

  async function doSave() {
    const st = $('#saveStatus');
    if (!cfg().token) {
      st.textContent = 'Not saved: no GitHub token is configured on this device.';
      return;
    }
    st.textContent = 'Saving to repo…';
    try {
      await saveToRepo(S.name, S.buf);
      st.textContent = 'Saved to data/' + S.name;
    } catch (e) {
      st.textContent = 'Save failed: ' + e.message;
    }
  }

  function afterLoad() {
    $('#saveBtn').classList.toggle('hide', S.src !== 'upload');
    $('#saveStatus').textContent = S.src === 'repo' ? '· loaded from data/' : '';
    if (S.src === 'upload' && $('#autosave').checked) doSave();
  }

  function showTab(t) {
    $('#tabDash').classList.toggle('hide', t !== 'dash');
    $('#files').classList.toggle('hide', t !== 'files');
    $('#tDash').classList.toggle('active', t === 'dash');
    $('#tFiles').classList.toggle('active', t === 'files');
    if (t === 'files') listFiles();
  }

  async function listFiles() {
    const c = cfg();
    const box = $('#flist');
    $('#repoLabel').textContent = c.owner && c.repo ? `${c.owner}/${c.repo} · ${c.branch} · data/` : 'No repository configured';
    box.innerHTML = '<p class="sub">Loading…</p>';
    try {
      let items = [];
      try { items = await gh('data', { ref: 1 }); } catch (e) { if (e.status !== 404) throw e; }
      items = items.filter((i) => i.type === 'file' && /\.(csv|tsv|xlsx|xls|xlsm)$/i.test(i.name));
      box.innerHTML = items.length ? '<table><tr><th>File</th><th>Size</th><th></th></tr>' + items.map((i) => `<tr><td>${esc(i.name)}</td><td class="n">${(i.size / 1024).toFixed(1)} KB</td><td style="text-align:right"><button class="btn sm" data-a="open" data-n="${esc(i.name)}">Open</button> <button class="btn alt sm del" data-a="del" data-n="${esc(i.name)}" data-sha="${i.sha}">Delete</button></td></tr>`).join('') + '</table>' : '<p class="sub">No saved files yet. Upload a file on the Dashboard tab with saving turned on.</p>';
    } catch (e) {
      box.innerHTML = `<p class="warn">${esc(e.message)}</p>`;
    }
  }

  async function openFromRepo(name) {
    try {
      const buf = await gh(fpath(name), { raw: 1, ref: 1 });
      S.wb = XLSX.read(buf, { type: 'array', cellDates: true });
      S.name = name;
      S.src = 'repo';
      showTab('dash');
      openSheet(S.wb.SheetNames[0]);
      afterLoad();
    } catch (e) {
      $('#flist').insertAdjacentHTML('afterbegin', `<p class="warn">Could not open ${esc(name)}: ${esc(e.message)}</p>`);
    }
  }

  $('#flist').onclick = async (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    const n = b.dataset.n;
    if (b.dataset.a === 'open') return openFromRepo(n);
    if (!cfg().token) return alert('Deleting needs a GitHub token configured on this device.');
    if (!confirm(`Delete ${n} from the repo? It stays in git history.`)) return;

    const row = b.closest('tr');
    if (row) row.remove();
    if (S.name === n) {
      $('#dash').classList.add('hide');
      $('#drop').classList.remove('hide');
      $('#file').value = '';
      S.wb = null;
      S.name = '';
      S.rows = [];
      S.cols = [];
      $('#fname').textContent = '';
      $('#saveStatus').textContent = '';
      $('#fmeta').textContent = '';
      $('#tbl').innerHTML = '';
      $('#prof').innerHTML = '';
      $('#kpis').innerHTML = '';
    }

    b.disabled = true;
    try {
      await gh(fpath(n), { method: 'DELETE', body: { message: 'Delete ' + n, sha: b.dataset.sha, branch: cfg().branch } });
      await listFiles();
    } catch (x) {
      alert('Delete failed: ' + x.message);
      if (row) {
        const current = $('#flist').querySelector('table');
        if (current) current.appendChild(row);
      }
      await listFiles();
    }
  };

  $('#tDash').onclick = () => showTab('dash');
  $('#tFiles').onclick = () => showTab('files');
  $('#refresh').onclick = listFiles;
  $('#saveBtn').onclick = doSave;
  $('#settingsBtn').onclick = () => {
    const c = cfg();
    $('#gOwner').value = c.owner || '';
    $('#gRepo').value = c.repo || '';
    $('#gBranch').value = c.branch;
    $('#gToken').value = c.token || '';
    $('#dlg').showModal();
  };
  $('#gCancel').onclick = () => $('#dlg').close();
  $('#gSave').onclick = () => {
    try { localStorage.setItem(CFG_KEY, JSON.stringify({ owner: $('#gOwner').value.trim(), repo: $('#gRepo').value.trim(), branch: $('#gBranch').value.trim() || 'main', token: $('#gToken').value.trim() })); } catch (e) {}
    $('#dlg').close();
    if (!$('#files').classList.contains('hide')) listFiles();
  };

  async function snap() {
    const el = $('#dash');
    const bg = getComputedStyle(document.body).backgroundColor;
    document.body.classList.add('exporting');
    try {
      return await html2canvas(el, { backgroundColor: bg, scale: 2, useCORS: true, windowWidth: Math.max(el.scrollWidth, 1100) });
    } finally {
      document.body.classList.remove('exporting');
    }
  }

  const baseName = () => S.name.replace(/\.[^.]+$/, '') + '-dashboard';

  function dl(url, name) {
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  async function exportAs(kind, btn) {
    const label = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Rendering…';
    try {
      const c = await snap();
      const img = c.toDataURL('image/jpeg', .92);
      if (kind === 'jpg') dl(img, baseName() + '.jpg');
      else {
        const w = c.width / 2;
        const h = c.height / 2;
        const pdf = new window.jspdf.jsPDF({ orientation: w > h ? 'l' : 'p', unit: 'pt', format: [w, h] });
        pdf.addImage(img, 'JPEG', 0, 0, w, h);
        pdf.save(baseName() + '.pdf');
      }
    } catch (e) {
      alert('Export failed: ' + e.message);
    }
    btn.disabled = false;
    btn.textContent = label;
  }

  $('#expPdf').onclick = (e) => exportAs('pdf', e.currentTarget);
  $('#expJpg').onclick = (e) => exportAs('jpg', e.currentTarget);

  const settingsBtn = $('#settingsBtn');
  if (settingsBtn && /settings/.test(location.search + location.hash)) settingsBtn.classList.remove('hide');
  document.addEventListener('keydown', (e) => { if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 's') { e.preventDefault(); settingsBtn && settingsBtn.click(); } });

  const sampleBtn = $('#sample');
  if (sampleBtn) {
    sampleBtn.onclick = () => {
      const reg = ['North', 'South', 'East', 'West'];
      const cat = ['Hardware', 'Software', 'Services', 'Support'];
      let seed = 7;
      const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
      const rows = [['Order date', 'Region', 'Category', 'Units', 'Unit price', 'Revenue', 'Discount %']];
      for (let i = 0; i < 320; i++) {
        const d = new Date(2025, 0, 1 + Math.floor(i * 1.1));
        const u = 1 + Math.floor(rnd() * 30);
        const p = +(20 + rnd() * 180).toFixed(2);
        const ds = Math.floor(rnd() * 4) * 5;
        rows.push([d.toISOString().slice(0, 10), reg[Math.floor(rnd() * 4)], cat[Math.floor(rnd() * 4)], u, p, +(u * p * (1 - ds / 100)).toFixed(2), ds]);
      }
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sales');
      S.wb = wb;
      S.name = 'sample-sales.xlsx';
      S.src = 'sample';
      openSheet('Sales');
      afterLoad();
    };
  }

  showTab('dash');
}
