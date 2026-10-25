// Test de bout en bout du cœur : lecture du classeur -> chiffrement -> déchiffrement -> recherche.
// Usage : node tests/test.mjs chemin/vers/classeur.xlsx
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { parseWorkbook, deriveKey, encryptDataset, decryptDataset, buildIndex, searchIndices, refOf, compareRef, classifySub, norm } from '../core.js';
import { LEXIQUE } from '../lexique.js';
const XLSX = createRequire(import.meta.url)('../vendor/xlsx.mini.min.js');

const file = process.argv[2];
const t0 = Date.now();
const data = parseWorkbook(XLSX, readFileSync(file));
console.log('lecture', Date.now() - t0, 'ms', data.journals.length, 'revues', data.years, data.label.slice(0, 60), data.warnings.slice(0, 3));
assert.equal(data.journals.length, 5355);
assert.deepEqual(data.years, [2021, 2022, 2023, 2024, 2025]);

const key = await deriveKey('test-passphrase');
const enc = await encryptDataset(data, key);
console.log('data.enc', (enc.length / 1024).toFixed(0), 'Ko');
const back = await decryptDataset(enc, key);
assert.equal(back.journals.length, 5355);
await assert.rejects(decryptDataset(enc, await deriveKey('mauvais')), /badkey/);

const idx = buildIndex(back, LEXIQUE);
const run = (q, o = {}) => (searchIndices(idx, { q, ...o }) || []).map((i) => back.journals[i]);
const names = (q, o) => run(q, o).map((j) => j.t);
const doms = (q, o) => [...new Set(run(q, o).flatMap((j) => j.dm))];

const trach = run('trachée');
console.log('trachée ->', trach.length, doms('trachée').join(' | '));
assert(doms('trachée').includes('ORL'));
console.log('trachea ->', run('trachea').length, '| laryngologie ->', run('laryngologie').length, '| larynx ->', run('larynx').length);
console.log('cardio fr/en:', run('cardiologie').length, run('cardiology').length, run('heart').length, run('coeur').length);
console.log('exact "ORL":', run('orl', { exact: true }).length, '| exact "laryngol":', run('laryngol', { exact: true }).length, '| exact "laryngol" approx:', run('laryngol').length);
console.log('faute "cardioligie":', run('cardioligie').length);
console.log('hyperbare:', names('hyperbare').slice(0, 6), run('plongée').length);
console.log('titre "lancet":', names('lancet', { scope: 'titre' }));
console.log('domaine "pédiatrie" titre scope:', run('pédiatrie', { scope: 'titre' }).length, 'dom scope:', run('pédiatrie', { scope: 'dom' }).length);
console.log('cancer poumon ET/OU:', run('cancer poumon').length, run('cancer poumon', { mode: 'or' }).length);
console.log('vide:', searchIndices(idx, { q: '  ' }));

// tri
const sel = run('orl').map((j) => ({ j, ref: refOf(j, back.years, 'last') })).sort(compareRef);
console.log(sel.slice(0, 5).map((x) => `${x.ref.cat} ${x.ref.if} ${x.ref.year} ${x.j.t}`));
const order = ['A+', 'A', 'B', 'C', 'D', 'E', 'NC'];
for (let k = 1; k < sel.length; k++) assert(order.indexOf(sel[k - 1].ref.cat) <= order.indexOf(sel[k].ref.cat));
const subs = {}; back.journals.forEach((j) => { const c = classifySub(j.so); subs[c] = (subs[c] || 0) + 1; });
console.log('soumission', subs);
const plos = back.journals.find((j) => j.t === 'PLoS One');
console.log(plos.li, plos.y[2025], classifySub(plos.so));
console.log('OK');
