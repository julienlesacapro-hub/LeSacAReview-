// Cœur de l'application : lecture du classeur Excel, chiffrement, recherche.
// Aucune dépendance DOM : ce fichier est aussi exécuté sous Node par les tests.

export const RANKS = ['A+', 'A', 'B', 'C', 'D', 'E', 'NC'];
export const SCHEMA = 1;

// ---------- Texte ----------
export const norm = (s) =>
  String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/œ/g, 'oe')
    .replace(/æ/g, 'ae');

export const tokenize = (s) => norm(s).split(/[^a-z0-9]+/).filter(Boolean);

const STOP = new Set(['le', 'la', 'les', 'l', 'de', 'du', 'des', 'd', 'et', 'ou', 'un', 'une', 'en', 'au', 'aux', 'of', 'the', 'and', 'in', 'for', 'to', 'a']);

// ---------- Chiffrement (AES-GCM 256, clé dérivée par PBKDF2) ----------
// Le dépôt GitHub est public : seul un fichier chiffré (data.enc) y est publié.
const MAGIC = [0x4c, 0x53, 0x52, 0x31]; // "LSR1"
const SALT = new TextEncoder().encode('LeSacAReview-v1!'); // 16 octets, fixe : permet de garder la clé en cache entre deux mises à jour
const ITER = 600000;

export async function deriveKey(passphrase) {
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(passphrase.normalize('NFC')), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: SALT, iterations: ITER, hash: 'SHA-256' },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

async function pipe(bytes, stream) {
  const out = new Response(new Blob([bytes]).stream().pipeThrough(stream));
  return new Uint8Array(await out.arrayBuffer());
}

export async function encryptDataset(obj, key) {
  let bytes = new TextEncoder().encode(JSON.stringify(obj));
  let flag = 0;
  if (typeof CompressionStream !== 'undefined') {
    bytes = await pipe(bytes, new CompressionStream('gzip'));
    flag = 1;
  }
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes));
  const out = new Uint8Array(4 + 1 + 12 + ct.length);
  out.set(MAGIC, 0);
  out[4] = flag;
  out.set(iv, 5);
  out.set(ct, 17);
  return out;
}

export async function decryptDataset(buf, key) {
  const b = new Uint8Array(buf);
  if (b.length < 40 || MAGIC.some((m, i) => b[i] !== m)) throw new Error('format');
  const flag = b[4];
  let plain;
  try {
    plain = new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b.slice(5, 17) }, key, b.slice(17)));
  } catch {
    throw new Error('badkey');
  }
  if (flag === 1) plain = await pipe(plain, new DecompressionStream('gzip'));
  return JSON.parse(new TextDecoder().decode(plain));
}

// Mode comptes utilisateurs : même contenu, gzip sans chiffrement (l'accès est contrôlé par le serveur).
const MAGIC0 = [0x4c, 0x53, 0x52, 0x30]; // "LSR0"
export async function packDataset(obj) {
  const raw = new TextEncoder().encode(JSON.stringify(obj));
  const body = typeof CompressionStream !== 'undefined' ? await pipe(raw, new CompressionStream('gzip')) : raw;
  const out = new Uint8Array(5 + body.length);
  out.set(MAGIC0, 0);
  out[4] = typeof CompressionStream !== 'undefined' ? 1 : 0;
  out.set(body, 5);
  return out;
}
export async function unpackDataset(buf) {
  const b = new Uint8Array(buf);
  if (b.length < 8 || MAGIC0.some((m, i) => b[i] !== m)) throw new Error('format');
  const body = b[4] === 1 ? await pipe(b.slice(5), new DecompressionStream('gzip')) : b.slice(5);
  return JSON.parse(new TextDecoder().decode(body));
}

// ---------- Lecture du classeur ----------
const COLS = [
  ['nlm', ['nlmid']],
  ['issn', ['issn']],
  ['essn', ['essn']],
  ['st', ["statut d'indexation", 'statut']],
  ['t', ['titre']],
  ['ed', ['jcr edition', 'jcr']],
  ['dc', ['discipline(s) wos', 'discipline']],
  ['dm', ['domaine(s)', 'domaine']],
  ['th', ['thematique(s)', 'thematique']],
  ['li', ['lien vers la revue', 'lien']],
  ['so', ['soumission']],
  ['apc', ['apc verifie']],
  ['apce', ['ordre de grandeur apc']],
  ['cr', ['case reports']],
  ['vf', ['niveau de verification']],
  ['src', ['source(s)', 'sources']],
  ['rm', ['remarques']],
];
const IGNORED = ['dernier if connu']; // recalculé à partir des années

const num = (v) => {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
};
const str = (v) => (v == null ? '' : String(v).trim());
const list = (v, sep) => str(v).split(sep).map((s) => s.trim()).filter(Boolean);

function normalizeRank(v, warn, ctx) {
  const r = str(v).toUpperCase().replace(/\s/g, '');
  if (!r) return null;
  if (RANKS.includes(r)) return r;
  warn(`Rang inconnu « ${v} » (${ctx}) : traité comme NC`);
  return 'NC';
}

function normalizeLink(v) {
  const s = str(v);
  if (!s || /^(à vérifier|—|-)$/i.test(s)) return '';
  if (/^https?:\/\//i.test(s)) return s;
  return /^[\w-]+(\.[\w-]+)+(\/|$)/.test(s) ? 'https://' + s : '';
}

export function parseWorkbook(XLSX, arrayBuffer, fileName = '') {
  const warnings = [];
  const warn = (m) => warnings.length < 40 && warnings.push(m);
  const wb = XLSX.read(arrayBuffer, { type: 'array' });

  const sheetName =
    wb.SheetNames.find((n) => norm(n) === 'revues') ||
    wb.SheetNames.find((n) => /titre/.test(norm(JSON.stringify((XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null })[0]) || []))));
  if (!sheetName) throw new Error("Onglet « Revues » introuvable (il doit contenir une colonne « Titre »).");
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null, raw: true });

  let hdr = rows.findIndex((r, i) => i < 10 && r.some((c) => norm(c).trim() === 'titre'));
  if (hdr < 0) throw new Error('Ligne d’en-tête introuvable : la colonne « Titre » est obligatoire.');
  const head = rows[hdr];
  const sub = rows[hdr + 1] || [];
  const hasSub = sub.some((c) => norm(c).trim() === 'if');
  const start = hdr + (hasSub ? 2 : 1);

  const col = {};
  const years = [];
  const extras = [];
  head.forEach((label, c) => {
    const l = norm(label).trim();
    if (!l) return;
    if (/^(19|20)\d\d$/.test(l)) {
      const ifc = c;
      const catc = c + 1;
      if (hasSub && !/^cat/.test(norm(sub[catc]).trim())) warn(`Année ${l} : la colonne suivante n'est pas « Cat. »`);
      years.push({ y: +l, ifc, catc });
      return;
    }
    if (IGNORED.some((p) => l.startsWith(p))) return;
    for (const [key, pats] of COLS) {
      if (col[key] == null && pats.some((p) => (p.length <= 5 ? l === p : l.startsWith(p)))) {
        col[key] = c;
        return;
      }
    }
    extras.push({ label: str(label), c });
  });
  years.sort((a, b) => a.y - b.y);
  if (col.t == null) throw new Error('Colonne « Titre » introuvable.');
  if (!years.length) warn("Aucune colonne d'année (ex. 2025 avec IF / Cat.) détectée.");
  for (const [key, pats] of COLS) if (col[key] == null && ['st', 'dm', 'li', 'so'].includes(key)) warn(`Colonne « ${pats[0]} » absente : la fonction correspondante sera vide.`);

  const get = (r, k) => (col[k] == null ? null : r[col[k]]);
  const journals = [];
  const seen = new Map();
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    const title = str(get(r, 't'));
    if (!title) continue;
    const j = {
      i: str(get(r, 'nlm')),
      t: title,
      is: str(get(r, 'issn')),
      es: str(get(r, 'essn')),
      st: str(get(r, 'st')),
      ed: str(get(r, 'ed')),
      dc: str(get(r, 'dc')),
      dm: list(get(r, 'dm'), ';'),
      th: list(get(r, 'th'), ';'),
      li: normalizeLink(get(r, 'li')),
      so: str(get(r, 'so')),
      apc: str(get(r, 'apc')),
      apce: str(get(r, 'apce')),
      cr: str(get(r, 'cr')),
      vf: str(get(r, 'vf')),
      src: str(get(r, 'src')),
      rm: str(get(r, 'rm')),
      y: {},
    };
    for (const { y, ifc, catc } of years) {
      const v = num(r[ifc]);
      const c = normalizeRank(r[catc], warn, `${title}, ${y}`);
      if (v != null || c) j.y[y] = [v, c];
    }
    if (extras.length) {
      j.x = {};
      for (const { label, c } of extras) {
        const v = str(r[c]);
        if (v) j.x[label] = v;
      }
    }
    const key = j.i || norm(title);
    if (seen.has(key)) warn(`Doublon ignoré : ${title} (${j.i || 'sans NLMid'})`);
    else {
      seen.set(key, true);
      journals.push(j);
    }
  }
  if (!journals.length) throw new Error('Aucune revue lue : vérifiez le fichier.');

  // Correspondance code WoS -> libellés (onglet facultatif)
  const domains = [];
  const cs = wb.SheetNames.find((n) => norm(n) === 'codes_disciplines');
  if (cs) {
    const cr = XLSX.utils.sheet_to_json(wb.Sheets[cs], { header: 1, defval: null });
    for (const r of cr.slice(1)) if (r[0] && r[1] && r[3]) domains.push({ code: str(r[0]), en: str(r[1]), fr: str(r[3]) });
  }

  // Libellé de version (onglet Lisez-moi, ligne « Version »)
  let label = '';
  const ls = wb.SheetNames.find((n) => norm(n).startsWith('lisez'));
  if (ls) {
    const lr = XLSX.utils.sheet_to_json(wb.Sheets[ls], { header: 1, defval: null });
    const v = lr.find((r) => norm(r[0]).trim() === 'version');
    if (v) label = str(v[1]);
  }

  return {
    schema: SCHEMA,
    generatedAt: new Date().toISOString(),
    source: fileName,
    label,
    years: years.map((y) => y.y),
    extras: extras.map((e) => e.label),
    domains,
    journals,
    warnings,
  };
}

// ---------- Règles métier ----------
export function classifySub(text) {
  const t = norm(text);
  if (!t || t.includes('a verifier')) return 'unknown';
  if (t.startsWith('gratuit') && t.includes('payant')) return 'mixed';
  if (t.startsWith('payant')) return 'paid';
  if (t.startsWith('gratuit')) return 'free';
  if (t.includes('payant') && t.includes('gratuit')) return 'mixed';
  if (t.includes('payant')) return 'paid';
  if (t.includes('gratuit')) return 'free';
  return 'unknown';
}

export function classifyCase(text) {
  const t = norm(text);
  if (t.startsWith('oui')) return 'yes';
  if (t.startsWith('non')) return 'no';
  if (t.startsWith('limit')) return 'limited';
  return 'unknown';
}

/** Rang + IF retenus pour une revue. yearSel = 'last' (dernier rang connu) ou une année. */
export function refOf(j, years, yearSel) {
  if (yearSel !== 'last') {
    const v = j.y[yearSel];
    return { year: yearSel, if: v ? v[0] : null, cat: (v && v[1]) || 'NC', missing: !v || (v[0] == null && !v[1]) };
  }
  for (let k = years.length - 1; k >= 0; k--) {
    const v = j.y[years[k]];
    if (v && v[1]) return { year: years[k], if: v[0], cat: v[1], missing: false };
  }
  for (let k = years.length - 1; k >= 0; k--) {
    const v = j.y[years[k]];
    if (v && v[0] != null) return { year: years[k], if: v[0], cat: 'NC', missing: false };
  }
  return { year: null, if: null, cat: 'NC', missing: true };
}

export function compareRef(a, b) {
  const ra = RANKS.indexOf(a.ref.cat);
  const rb = RANKS.indexOf(b.ref.cat);
  if (ra !== rb) return ra - rb;
  const ia = a.ref.if ?? -1;
  const ib = b.ref.if ?? -1;
  if (ia !== ib) return ib - ia;
  return a.j.t.localeCompare(b.j.t, 'en', { sensitivity: 'base' });
}

// ---------- Recherche ----------
const stem = (w) => {
  let s = norm(w);
  if (s.length > 4) s = s.replace(/[sx]$/, '');
  return s.length >= 7 ? s.slice(0, 6) : s;
};

function lev1(a, b) {
  // vrai si distance de Levenshtein <= 1
  if (Math.abs(a.length - b.length) > 1) return false;
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  if (i === a.length && i === b.length) return true;
  if (a.length === b.length) return a.slice(i + 1) === b.slice(i + 1) || (a[i] === b[i + 1] && a[i + 1] === b[i] && a.slice(i + 2) === b.slice(i + 2));
  return a.length > b.length ? a.slice(i + 1) === b.slice(i) : b.slice(i + 1) === a.slice(i);
}

export function buildIndex(data, lexicon = []) {
  const T = new Map(); // tokens du titre
  const D = new Map(); // domaines (FR + EN), thématiques
  const O = new Map(); // ISSN, NLMid, extras
  const byLabel = new Map(); // libellé normalisé -> indices
  const en = new Map(data.domains.map((d) => [norm(d.fr), d.en]));
  const put = (m, tok, i) => {
    const a = m.get(tok);
    if (!a) m.set(tok, [i]);
    else if (a[a.length - 1] !== i) a.push(i);
  };
  data.journals.forEach((j, i) => {
    tokenize(j.t).forEach((t) => put(T, t, i));
    [...j.dm, ...j.th].forEach((lab) => {
      const k = norm(lab);
      (byLabel.get(k) || byLabel.set(k, []).get(k)).push(i);
      tokenize(lab).forEach((t) => put(D, t, i));
      tokenize(en.get(k) || '').forEach((t) => put(D, t, i));
    });
    [j.i, j.is, j.es, ...Object.values(j.x || {})].forEach((v) => tokenize(v).forEach((t) => put(O, t, i)));
  });
  return { T, D, O, byLabel, lexicon: lexicon.map((e) => ({ a: e.a.map(norm), d: (e.d || []).map(norm), t: (e.t || []).map(norm) })), n: data.journals.length };
}

function prefixHits(map, st, w) {
  const out = [];
  const short = st.length < 4;
  for (const [tok, arr] of map) if (short ? tok === st || tok === w : tok.startsWith(st)) out.push(arr);
  return out;
}

function wordSet(index, w, scope, exact) {
  const maps = scope === 'titre' ? [index.T] : scope === 'dom' ? [index.D] : [index.T, index.D, index.O];
  const res = new Set();
  const addAll = (arrs) => arrs.forEach((a) => a.forEach((i) => res.add(i)));
  if (exact) {
    maps.forEach((m) => m.get(w) && m.get(w).forEach((i) => res.add(i)));
    return res;
  }
  const st = stem(w);
  maps.forEach((m) => addAll(prefixHits(m, st, w)));
  // synonymes / traduction FR <-> EN via le lexique
  for (const e of index.lexicon) {
    if (!e.a.some((a) => a === w || (st.length >= 4 && a.startsWith(st)) || (a.length >= 4 && w.startsWith(a)))) continue;
    if (scope !== 'titre') e.d.forEach((lab) => (index.byLabel.get(lab) || []).forEach((i) => res.add(i)));
    if (scope !== 'dom')
      e.t.forEach((ts) => {
        if (ts[0] === '=') index.T.get(ts.slice(1)) && index.T.get(ts.slice(1)).forEach((i) => res.add(i)); // mot entier
        else addAll(prefixHits(index.T, ts, ts));
      });
  }
  // faute de frappe (1 caractère) si rien trouvé
  if (!res.size && st.length >= 5) {
    for (const m of maps) for (const [tok, arr] of m) if (tok.length >= st.length && lev1(st, tok.slice(0, st.length))) arr.forEach((i) => res.add(i));
  }
  return res;
}

/** Retourne la liste des indices de revues correspondant, ou null si la requête est vide. */
export function searchIndices(index, { q = '', scope = 'all', exact = false, mode = 'and' } = {}) {
  const words = tokenize(q).filter((w) => !STOP.has(w) && (w.length > 1 || /\d/.test(w)));
  if (!words.length) return null;
  const sets = words.map((w) => wordSet(index, w, scope, exact));
  let acc;
  if (mode === 'or') {
    acc = new Set();
    sets.forEach((s) => s.forEach((i) => acc.add(i)));
  } else {
    sets.sort((a, b) => a.size - b.size);
    acc = sets[0];
    for (let k = 1; k < sets.length; k++) acc = new Set([...acc].filter((i) => sets[k].has(i)));
  }
  return [...acc];
}
