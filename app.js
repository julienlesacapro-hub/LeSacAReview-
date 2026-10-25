import { RANKS, norm, deriveKey, encryptDataset, decryptDataset, parseWorkbook, buildIndex, searchIndices, refOf, compareRef, classifySub, classifyCase } from './core.js';
import { LEXIQUE } from './lexique.js';

const $ = (s, r = document) => r.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const fmtIF = (v) => (v == null ? '—' : v.toLocaleString('fr-FR', { maximumFractionDigits: 3 }));
const rkClass = (c) => 'rk-' + c.toLowerCase().replace('+', 'plus');
const PAGE = 40;
const SUB_LABEL = { free: 'Gratuite', paid: 'Payante', mixed: 'Gratuite · OA payant', unknown: 'À vérifier' };
const CASE_LABEL = { yes: 'Oui', limited: 'Limité', no: 'Non', unknown: 'À vérifier' };
const NLM = (id) => `https://www.ncbi.nlm.nih.gov/nlmcatalog/${encodeURIComponent(id)}`;
const REPO_UPLOAD = 'https://github.com/julienlesacapro-hub/LeSacAReview-/upload/main';

// ---------- Stockage local (jamais bloquant) ----------
const ls = {
  get: (k, d) => { try { const v = localStorage.getItem('lsr.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem('lsr.' + k, JSON.stringify(v)); } catch { /* ignoré */ } },
};
const idb = (mode, fn) =>
  new Promise((res, rej) => {
    const o = indexedDB.open('lsr', 1);
    o.onupgradeneeded = () => o.result.createObjectStore('kv');
    o.onerror = () => rej(o.error);
    o.onsuccess = () => {
      const tx = o.result.transaction('kv', mode);
      const r = fn(tx.objectStore('kv'));
      tx.oncomplete = () => res(r.result);
      tx.onerror = () => rej(tx.error);
    };
  });
const keyGet = () => idb('readonly', (s) => s.get('key')).catch(() => null);
const keySet = (k) => idb('readwrite', (s) => s.put(k, 'key')).catch(() => {});
const keyDel = () => idb('readwrite', (s) => s.delete('key')).catch(() => {});

// ---------- État ----------
const emptyF = () => ({ year: 'last', ranks: new Set(), ifMin: '', ifMax: '', sub: new Set(), link: false, st: new Set(), ed: new Set(), dom: new Set(), th: new Set(), cr: new Set(), vf: '', x: {} });
const S = { data: null, index: null, key: null, buf: null, q: '', scope: 'all', exact: ls.get('exact', false), all: ls.get('all', true), sort: ls.get('sort', 'rank'), f: emptyF(), rows: [], shown: 0 };

// ---------- Démarrage ----------
const show = (id) => ['app', 'gate', 'empty'].forEach((x) => ($('#' + x).hidden = x !== id));

async function boot() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  let buf = null;
  try {
    const r = await fetch('data.enc', { cache: 'no-cache' });
    if (r.ok) buf = await r.arrayBuffer();
  } catch { /* hors ligne sans cache */ }
  if (!buf) return showEmpty();
  S.buf = buf;
  const key = await keyGet();
  if (key) {
    try {
      S.key = key;
      return load(await decryptDataset(buf, key));
    } catch (e) {
      S.key = null;
      if (e.message === 'badkey') await keyDel();
      if (e.message === 'format') return showEmpty();
    }
  }
  show('gate');
  $('#meta').textContent = 'Accès protégé';
  $('#gate-pass').focus();
}

function showEmpty(msg) {
  $('#meta').textContent = 'Aucune donnée';
  if (msg) $('#empty-msg').textContent = msg;
  show('empty');
}

$('#gate-form').addEventListener('submit', async (e) => {
  e.preventDefault();
  const err = $('#gate-error');
  err.hidden = true;
  const btn = e.submitter;
  btn.disabled = true;
  btn.textContent = 'Vérification…';
  try {
    const key = await deriveKey($('#gate-pass').value);
    const data = await decryptDataset(S.buf, key);
    S.key = key;
    if ($('#gate-remember').checked) await keySet(key);
    $('#gate-pass').value = '';
    load(data);
  } catch (x) {
    err.textContent = x.message === 'badkey' ? "Code d'accès incorrect." : 'Fichier de données illisible : ' + x.message;
    err.hidden = false;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Déverrouiller';
  }
});

$('#btn-lock').addEventListener('click', async () => {
  await keyDel();
  location.reload();
});

function load(data) {
  S.data = data;
  S.index = buildIndex(data, LEXIQUE);
  S.f = emptyF();
  const y = data.years;
  const d = new Date(data.generatedAt);
  $('#meta').textContent = `${data.journals.length.toLocaleString('fr-FR')} revues · ${y[0]}–${y[y.length - 1]} · mis à jour le ${d.toLocaleDateString('fr-FR')}`;
  $('#meta').title = data.label || '';
  $('#btn-lock').hidden = false;
  buildFilters();
  $('#exact').checked = S.exact;
  $('#allwords').checked = S.all;
  $('#sort').value = S.sort;
  show('app');
  run();
}

// ---------- Recherche + filtres ----------
let timer;
const schedule = () => { clearTimeout(timer); timer = setTimeout(run, 140); };

function matchesFilters(j, ref) {
  const f = S.f;
  if (f.ranks.size && !f.ranks.has(ref.cat)) return false;
  if (f.ifMin !== '' && !(ref.if != null && ref.if >= f.ifMin)) return false;
  if (f.ifMax !== '' && !(ref.if != null && ref.if <= f.ifMax)) return false;
  if (f.sub.size && !f.sub.has(classifySub(j.so))) return false;
  if (f.link && !j.li) return false;
  if (f.st.size && !f.st.has(j.st)) return false;
  if (f.ed.size && !j.ed.split(/\s*,\s*/).some((e) => f.ed.has(e))) return false;
  if (f.dom.size && !j.dm.some((d) => f.dom.has(d))) return false;
  if (f.th.size && !j.th.some((t) => f.th.has(t))) return false;
  if (f.cr.size && !f.cr.has(classifyCase(j.cr))) return false;
  if (f.vf && j.vf !== f.vf) return false;
  for (const [k, v] of Object.entries(f.x)) {
    if (!v) continue;
    const cur = (j.x && j.x[k]) || '';
    if (S.data.xmode[k] === 'select' ? cur !== v : !norm(cur).includes(norm(v))) return false;
  }
  return true;
}

function run() {
  const { data, index, f } = S;
  const hits = searchIndices(index, { q: S.q, scope: S.scope, exact: S.exact, mode: S.all ? 'and' : 'or' });
  const pool = hits ? hits.map((i) => data.journals[i]) : data.journals;
  const rows = [];
  for (const j of pool) {
    const ref = refOf(j, data.years, f.year);
    if (matchesFilters(j, ref)) rows.push({ j, ref });
  }
  if (S.sort === 'rank') rows.sort(compareRef);
  else if (S.sort === 'if') rows.sort((a, b) => (b.ref.if ?? -1) - (a.ref.if ?? -1) || compareRef(a, b));
  else rows.sort((a, b) => a.j.t.localeCompare(b.j.t, 'en', { sensitivity: 'base' }));
  S.rows = rows;
  S.shown = 0;
  $('#list').innerHTML = '';
  $('#count').innerHTML = `<b>${rows.length.toLocaleString('fr-FR')}</b> revue${rows.length > 1 ? 's' : ''} <small>sur ${data.journals.length.toLocaleString('fr-FR')}</small>`;
  if (!rows.length) {
    $('#list').innerHTML = `<li class="noresult"><b>Aucune revue trouvée</b>${S.exact ? 'Décochez « Termes exacts » pour accepter synonymes et formes proches, ou ' : ''}Essayez un autre mot, ou retirez des filtres.</li>`;
    $('#more').innerHTML = '';
  } else more();
  renderChips();
  $('#q-clear').hidden = !S.q;
}

function more() {
  const end = Math.min(S.rows.length, S.shown + PAGE);
  $('#list').insertAdjacentHTML('beforeend', S.rows.slice(S.shown, end).map((r, k) => card(r, S.shown + k)).join(''));
  S.shown = end;
  $('#more').innerHTML = end < S.rows.length ? `<button class="btn ghost" id="more-btn">Afficher la suite (${(S.rows.length - end).toLocaleString('fr-FR')})</button>` : '';
  const b = $('#more-btn');
  if (b) { b.onclick = more; obs.observe(b); }
}
const obs = new IntersectionObserver((es) => es.some((e) => e.isIntersecting) && more(), { rootMargin: '600px' });

// ---------- Cartes ----------
const EXT = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/></svg>';

function goButton(j) {
  if (j.li) return `<a class="btn link" href="${esc(j.li)}" target="_blank" rel="noopener noreferrer" title="${esc(j.li)}">Site de la revue ${EXT}</a>`;
  if (j.i) return `<a class="btn nlm" href="${NLM(j.i)}" target="_blank" rel="noopener noreferrer" title="Pas de site renseigné : fiche du catalogue NLM">Fiche NLM ${EXT}</a>`;
  return '';
}

function subBadge(j) {
  const c = classifySub(j.so);
  return `<span class="badge ${c}" title="${esc(j.so || 'Non renseigné')}">${SUB_LABEL[c]}</span>`;
}

function rankBlock(ref, years, cls = '') {
  const last = years[years.length - 1];
  const old = ref.year != null && ref.year !== last;
  return `<div class="rank"><div class="letter ${ref.cat === 'NC' ? 'nc' : ''}" aria-label="Rang ${ref.cat}">${ref.cat}</div>
    <div class="if">IF ${fmtIF(ref.if)}${ref.year ? `<small class="yr ${old ? 'old' : ''}" ${old ? `title="Pas de donnée ${last} : dernière année disponible"` : ''}>${ref.year}</small>` : '<small>aucune donnée</small>'}</div></div>`;
}

function card({ j, ref }, n) {
  const doms = [...j.dm, ...j.th].slice(0, 4).join(' · ');
  return `<li class="card ${rkClass(ref.cat)}" data-n="${n}">
    ${rankBlock(ref, S.data.years)}
    <div class="main"><h3><button class="tbtn" type="button">${esc(j.t)}</button></h3>
      ${doms ? `<p class="sub">${esc(doms)}</p>` : ''}
      <div class="badges">${subBadge(j)}${j.st ? `<span class="tag" title="${j.st === 'CI' ? 'Indexée MEDLINE' : j.st === 'NCI' ? 'Non indexée MEDLINE actuellement' : ''}">${esc(j.st)}</span>` : ''}${j.ed ? `<span class="tag">${esc(j.ed)}</span>` : ''}</div></div>
    <div class="go">${goButton(j)}</div>
  </li>`;
}

$('#list').addEventListener('click', (e) => {
  if (e.target.closest('a')) return;
  const li = e.target.closest('.card');
  if (li) openDetail(S.rows[+li.dataset.n]);
});

function openDetail({ j, ref }) {
  const yrs = S.data.years;
  const kv = (k, v, raw) => (v ? `<dt>${esc(k)}</dt><dd>${raw ? v : esc(v)}</dd>` : '');
  const urlify = (t) => esc(t).replace(/(https?:\/\/[^\s;]+)/g, '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>');
  const hist = `<table class="hist"><thead><tr><th></th>${yrs.map((y) => `<th>${y}</th>`).join('')}</tr></thead><tbody>
    <tr><th>Rang</th>${yrs.map((y) => { const c = j.y[y] && j.y[y][1]; return `<td>${c ? `<span class="mini ${rkClass(c)}">${c}</span>` : '–'}</td>`; }).join('')}</tr>
    <tr><th>IF</th>${yrs.map((y) => `<td>${j.y[y] && j.y[y][0] != null ? fmtIF(j.y[y][0]) : '–'}</td>`).join('')}</tr></tbody></table>`;
  const estim = /estimation|vérifier/i.test(j.vf + j.apce + j.so) && !/^Vérifié/.test(j.vf);
  $('#detail').innerHTML = `
    <form method="dialog" class="sheet-close"><button class="icon-btn" aria-label="Fermer">×</button></form>
    <div class="d-head ${rkClass(ref.cat)}"><div class="letter">${ref.cat}</div><div><h2 id="d-title">${esc(j.t)}</h2>
      <div class="sub">IF ${fmtIF(ref.if)}${ref.year ? ' · ' + ref.year : ''}</div></div></div>
    <div class="badges" style="margin-top:12px">${subBadge(j)}${j.st ? `<span class="tag">${esc(j.st)}</span>` : ''}${j.ed ? `<span class="tag">${esc(j.ed)}</span>` : ''}</div>
    <div class="d-actions go">${goButton(j)}${j.li && j.i ? `<a class="btn nlm" href="${NLM(j.i)}" target="_blank" rel="noopener noreferrer">Fiche NLM ${EXT}</a>` : ''}</div>
    <h3>Historique</h3>${hist}
    ${estim ? '<p class="warnbox">Certaines informations de soumission / APC sont des estimations ou restent à vérifier : consultez le site de la revue avant de soumettre.</p>' : ''}
    <h3>Soumission &amp; publication</h3>
    <dl class="kv">${kv('Gratuite / payante', j.so)}${kv('APC vérifié', j.apc && j.apc !== '—' ? j.apc : '')}${kv('Ordre de grandeur APC', j.apce && j.apce !== '—' ? j.apce : '')}${kv('Case reports acceptés', j.cr)}</dl>
    <h3>Fiche</h3>
    <dl class="kv">${kv('Domaine(s)', j.dm.join(' · '))}${kv('Thématique(s)', j.th.join(' · '))}${kv('Indexation MEDLINE', j.st === 'CI' ? 'CI – indexée' : j.st === 'NCI' ? 'NCI – non indexée actuellement' : j.st)}${kv('Édition JCR', j.ed)}${kv('Disciplines WoS', j.dc)}${kv('ISSN', j.is)}${kv('eISSN', j.es)}${kv('NLMid', j.i)}${Object.entries(j.x || {}).map(([k, v]) => kv(k, v)).join('')}${kv('Vérification', j.vf)}${kv('Remarques', j.rm)}${kv('Sources', urlify(j.src), true)}</dl>`;
  $('#detail').showModal();
}

// ---------- Filtres ----------
function pills(name, items, set, cls = '') {
  return `<div class="pills">${items.map(([v, label, color]) => `<label class="pill ${cls}" ${color ? `style="--c:var(--${color})"` : ''}><input type="checkbox" data-f="${name}" value="${esc(v)}" ${set.has(v) ? 'checked' : ''}><span>${esc(label ?? v)}</span></label>`).join('')}</div>`;
}
const distinct = (fn) => [...new Set(S.data.journals.flatMap(fn).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'fr'));

function buildFilters() {
  const { data, f } = S;
  const J = data.journals;
  const domCount = new Map();
  J.forEach((j) => j.dm.forEach((d) => domCount.set(d, (domCount.get(d) || 0) + 1)));
  const doms = [...domCount.keys()].sort((a, b) => a.localeCompare(b, 'fr'));
  const ths = distinct((j) => j.th);
  const eds = distinct((j) => j.ed.split(/\s*,\s*/));
  const sts = distinct((j) => [j.st]).filter((s) => s !== 'À vérifier');
  const vfs = distinct((j) => [j.vf]);
  data.xmode = {};
  const xs = (data.extras || []).map((label) => {
    const vals = distinct((j) => [j.x && j.x[label]]);
    data.xmode[label] = vals.length <= 30 ? 'select' : 'text';
    return { label, vals };
  }).filter((x) => x.vals.length);

  $('#filters-body').innerHTML = `
    <div class="fgroup"><h3>Année de référence</h3>
      <select id="f-year"><option value="last">Dernier rang connu (auto)</option>${[...data.years].reverse().map((y) => `<option value="${y}">${y}</option>`).join('')}</select></div>
    <div class="fgroup"><h3>Rang</h3>${pills('ranks', RANKS.map((r) => [r, r, 'rk-' + r.toLowerCase().replace('+', 'plus')]), f.ranks, 'rk')}</div>
    <div class="fgroup"><h3>Impact Factor</h3><div class="row2">
      <label class="field">Min<input id="f-ifmin" type="number" step="0.1" min="0" inputmode="decimal"></label>
      <label class="field">Max<input id="f-ifmax" type="number" step="0.1" min="0" inputmode="decimal"></label></div></div>
    <div class="fgroup"><h3>Soumission</h3>${pills('sub', Object.entries(SUB_LABEL).map(([k, v]) => [k, v]), f.sub)}
      <label class="check" style="margin-top:10px"><input type="checkbox" id="f-link"><span>Avec site officiel renseigné</span></label></div>
    <div class="fgroup"><h3>Case reports acceptés</h3>${pills('cr', Object.entries(CASE_LABEL).map(([k, v]) => [k, v]), f.cr)}</div>
    <div class="fgroup"><h3>Domaine</h3>
      <input type="search" id="f-domq" placeholder="Filtrer la liste…">
      <div class="dom-list" id="f-doms">${doms.map((d) => `<label><input type="checkbox" data-f="dom" value="${esc(d)}"><span>${esc(d)}</span><small>${domCount.get(d)}</small></label>`).join('')}</div></div>
    ${ths.length ? `<div class="fgroup"><h3>Thématique</h3>${pills('th', ths.map((t) => [t]), f.th)}</div>` : ''}
    <div class="fgroup"><h3>Indexation MEDLINE</h3>${pills('st', sts.map((s) => [s, s === 'CI' ? 'CI · indexée' : s === 'NCI' ? 'NCI · non indexée' : s]), f.st)}</div>
    ${eds.length ? `<div class="fgroup"><h3>Édition JCR</h3>${pills('ed', eds.map((t) => [t]), f.ed)}</div>` : ''}
    ${vfs.length > 1 ? `<div class="fgroup"><h3>Niveau de vérification</h3><select id="f-vf"><option value="">Tous</option>${vfs.map((v) => `<option value="${esc(v)}">${esc(v.length > 70 ? v.slice(0, 68) + '…' : v)}</option>`).join('')}</select></div>` : ''}
    ${xs.map((x) => `<div class="fgroup"><h3>${esc(x.label)}</h3>${data.xmode[x.label] === 'select' ? `<select data-x="${esc(x.label)}"><option value="">Tous</option>${x.vals.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('')}</select>` : `<input type="search" data-x="${esc(x.label)}" placeholder="Contient…">`}</div>`).join('')}`;
}

const fb = $('#filters-body');
fb.addEventListener('change', (e) => {
  const t = e.target, f = S.f;
  if (t.dataset.f) {
    const set = f[t.dataset.f];
    t.checked ? set.add(t.value) : set.delete(t.value);
  } else if (t.id === 'f-year') f.year = t.value === 'last' ? 'last' : +t.value;
  else if (t.id === 'f-link') f.link = t.checked;
  else if (t.id === 'f-vf') f.vf = t.value;
  else if (t.dataset.x != null) f.x[t.dataset.x] = t.value;
  run();
});
fb.addEventListener('input', (e) => {
  const t = e.target;
  if (t.id === 'f-ifmin' || t.id === 'f-ifmax') { S.f[t.id === 'f-ifmin' ? 'ifMin' : 'ifMax'] = t.value === '' ? '' : +t.value; schedule(); }
  else if (t.id === 'f-domq') {
    const q = norm(t.value);
    fb.querySelectorAll('#f-doms label').forEach((l) => (l.hidden = !!q && !norm(l.textContent).includes(q)));
  } else if (t.dataset.x != null && t.type === 'search') { S.f.x[t.dataset.x] = t.value; schedule(); }
});

function syncFilterUI() {
  const f = S.f;
  fb.querySelectorAll('input[data-f]').forEach((i) => (i.checked = f[i.dataset.f].has(i.value)));
  fb.querySelectorAll('select[data-x], input[data-x]').forEach((i) => (i.value = f.x[i.dataset.x] || ''));
  $('#f-year').value = String(f.year);
  $('#f-ifmin').value = f.ifMin; $('#f-ifmax').value = f.ifMax;
  $('#f-link').checked = f.link;
  if ($('#f-vf')) $('#f-vf').value = f.vf;
}

function activeChips() {
  const f = S.f, out = [];
  const add = (label, fn) => out.push({ label, fn });
  if (f.year !== 'last') add(`Année ${f.year}`, () => (f.year = 'last'));
  f.ranks.forEach((v) => add(`Rang ${v}`, () => f.ranks.delete(v)));
  if (f.ifMin !== '') add(`IF ≥ ${f.ifMin}`, () => (f.ifMin = ''));
  if (f.ifMax !== '') add(`IF ≤ ${f.ifMax}`, () => (f.ifMax = ''));
  f.sub.forEach((v) => add(SUB_LABEL[v], () => f.sub.delete(v)));
  if (f.link) add('Avec site officiel', () => (f.link = false));
  f.cr.forEach((v) => add(`Case reports : ${CASE_LABEL[v]}`, () => f.cr.delete(v)));
  f.dom.forEach((v) => add(v, () => f.dom.delete(v)));
  f.th.forEach((v) => add(v, () => f.th.delete(v)));
  f.st.forEach((v) => add(v, () => f.st.delete(v)));
  f.ed.forEach((v) => add(v, () => f.ed.delete(v)));
  if (f.vf) add('Vérification : ' + (f.vf.length > 24 ? f.vf.slice(0, 22) + '…' : f.vf), () => (f.vf = ''));
  Object.entries(f.x).forEach(([k, v]) => v && add(`${k} : ${v}`, () => (f.x[k] = '')));
  return out;
}

function renderChips() {
  const chips = activeChips();
  $('#chips').innerHTML = chips.map((c, i) => `<button class="chip" data-i="${i}" type="button">${esc(c.label)} <i aria-hidden="true">×</i></button>`).join('');
  const n = $('#filters-count');
  n.hidden = !chips.length;
  n.textContent = chips.length;
  $('#chips')._chips = chips;
}
$('#chips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip');
  if (!b) return;
  $('#chips')._chips[+b.dataset.i].fn();
  syncFilterUI();
  run();
});
$('#f-reset').addEventListener('click', () => { S.f = emptyF(); syncFilterUI(); run(); });
$('#btn-filters').addEventListener('click', () => $('#filters').classList.add('open'));
$('#f-close').addEventListener('click', () => $('#filters').classList.remove('open'));
$('#f-apply').addEventListener('click', () => { $('#filters').classList.remove('open'); window.scrollTo({ top: 0 }); });

// ---------- Saisie ----------
$('#q').addEventListener('input', (e) => { S.q = e.target.value; schedule(); });
$('#q-clear').addEventListener('click', () => { $('#q').value = ''; S.q = ''; run(); $('#q').focus(); });
$('#search-form').addEventListener('submit', (e) => { e.preventDefault(); $('#q').blur(); run(); });
document.querySelectorAll('input[name="scope"]').forEach((r) => r.addEventListener('change', () => { S.scope = r.value; run(); }));
$('#exact').addEventListener('change', (e) => { S.exact = e.target.checked; ls.set('exact', S.exact); run(); });
$('#allwords').addEventListener('change', (e) => { S.all = e.target.checked; ls.set('all', S.all); run(); });
$('#sort').addEventListener('change', (e) => { S.sort = e.target.value; ls.set('sort', S.sort); run(); });

// ---------- Mise à jour des données ----------
const U = { file: null, out: null };
const openUpdate = () => {
  $('#u-pass2').hidden = !!S.key;
  $('#u-pass').placeholder = S.key ? 'Laisser vide pour garder le code actuel' : "Choisissez un code d'accès";
  $('#u-pass2').placeholder = "Confirmer le code d'accès";
  $('#update').showModal();
};
$('#btn-update').addEventListener('click', openUpdate);
$('#empty-update').addEventListener('click', openUpdate);

const canGo = () => ($('#u-go').disabled = !U.file || (!S.key && !$('#u-pass').value));
$('#u-file').addEventListener('change', (e) => { U.file = e.target.files[0] || null; $('#u-fname').textContent = U.file ? U.file.name : 'aucun fichier'; canGo(); });
$('#u-pass').addEventListener('input', canGo);

function loadXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'vendor/xlsx.mini.min.js';
    s.onload = () => res(window.XLSX);
    s.onerror = () => rej(new Error('Lecteur Excel indisponible (hors ligne ?)'));
    document.head.append(s);
  });
}

$('#u-go').addEventListener('click', async () => {
  const out = $('#u-out'), go = $('#u-go');
  const fail = (m) => (out.innerHTML = `<div class="ko">${esc(m)}</div>`);
  go.disabled = true;
  out.innerHTML = '<p class="hint">Lecture du classeur…</p>';
  try {
    let key = S.key;
    const p = $('#u-pass').value;
    if (p) {
      if (!S.key && p !== $('#u-pass2').value) return fail('Les deux codes ne sont pas identiques.');
      if (p.length < 12) return fail("Code trop court : 12 caractères minimum (le fichier chiffré est public, un code faible peut être deviné). Une phrase de plusieurs mots convient très bien.");
      out.innerHTML = '<p class="hint">Dérivation de la clé…</p>';
      key = await deriveKey(p);
    }
    const XLSX = await loadXLSX();
    const data = parseWorkbook(XLSX, await U.file.arrayBuffer(), U.file.name);
    const bytes = await encryptDataset(data, key);
    // comparaison avec les données actuellement chargées
    let diff = '';
    if (S.data) {
      const k = (j) => j.i || norm(j.t);
      const old = new Set(S.data.journals.map(k)), cur = new Set(data.journals.map(k));
      const added = data.journals.filter((j) => !old.has(k(j))).length;
      const removed = S.data.journals.filter((j) => !cur.has(k(j))).length;
      const newYears = data.years.filter((y) => !S.data.years.includes(y));
      diff = `<li>${added} revue(s) ajoutée(s), ${removed} retirée(s) par rapport aux données actuelles${newYears.length ? `, nouvelle(s) année(s) : ${newYears.join(', ')}` : ''}</li>`;
    }
    U.out = { bytes, data, key };
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/octet-stream' }));
    out.innerHTML = `<div class="ok"><b>data.enc prêt</b> (${(bytes.length / 1024).toFixed(0)} Ko)<ul>
        <li>${data.journals.length.toLocaleString('fr-FR')} revues · années ${data.years.join(', ')}</li>${diff}
        ${data.extras.length ? `<li>Colonnes supplémentaires détectées : ${esc(data.extras.join(', '))}</li>` : ''}</ul></div>
      ${data.warnings.length ? `<div class="warnbox"><b>À vérifier dans le classeur</b><ul>${data.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul></div>` : ''}
      <a class="btn primary" id="u-dl" href="${url}" download="data.enc">1 · Télécharger data.enc</a>
      <a class="btn ghost" href="${REPO_UPLOAD}" target="_blank" rel="noopener noreferrer">2 · Déposer sur GitHub (Upload files) ${EXT}</a>
      <p class="hint">Glissez <code>data.enc</code> à la racine du dépôt (il remplace l'ancien), puis « Commit changes ». L'appli est à jour 1 à 2 minutes plus tard, sur tous vos appareils.</p>
      <button class="btn ghost" id="u-preview" type="button">Voir ces données ici, tout de suite (aperçu)</button>`;
    $('#u-preview').onclick = async () => {
      S.key = key; S.buf = bytes.buffer;
      $('#btn-lock').hidden = false;
      load(data);
      $('#update').close();
    };
  } catch (e) {
    fail('Échec : ' + e.message);
  } finally {
    canGo();
  }
});

boot();
