#!/usr/bin/env python3
# © 2026 LESACA Julien – Tous droits réservés.
"""Fusionne un ou plusieurs exports SIGAPS (.txt) dans le classeur « Revues par domaine ».

    python3 tools/fusion_sigaps.py --xlsx ancien.xlsx --export export1.txt [export2.txt ...] \
            --out nouveau.xlsx [--num 4] [--date 02/11/2026]

- Revues déjà présentes (même NLMid) : valeurs SIGAPS (colonnes A→H et années) mises à jour si elles changent ;
  toutes vos colonnes ajoutées (lien, soumission, APC, case reports, remarques…) sont conservées.
- Revues « Hors export » retrouvées par leur titre : rattachées à leur ligne SIGAPS.
- Nouvelles revues : ajoutées à la fin (triées par titre), colonnes ajoutées pré‑remplies « À vérifier ».
- Nouvelle année (ex. 2026 IF / Cat.) détectée dans l'export : colonnes ajoutées à la suite.
- Colonne « Domaine(s) » recalculée pour toutes les lignes à partir des codes de disciplines WoS.
- Onglets Codes_disciplines, Synthèse, Sources, Lisez-moi mis à jour ; plages de formules/mises en forme étendues.

Le classeur produit contient des Impact Factors Clarivate : ne jamais le publier (le dépôt GitHub est public).
"""
import argparse, collections, copy, re, sys
from pathlib import Path

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.cell_range import MultiCellRange

sys.path.insert(0, str(Path(__file__).parent))
from disciplines_fr import DISCIPLINES_FR

LIMIT = 30000  # dernière ligne couverte par les formules / mises en forme (ancienne valeur : 5857)
NLM = 'https://www.ncbi.nlm.nih.gov/nlmcatalog/'
NCI_REMARK = 'NCI = non indexée MEDLINE actuellement (définition SIGAPS) → valorisation SIGAPS à confirmer avec la DRCI.'
ABSENT = "Absente de l'export SIGAPS du 15/10/2026 : discipline WoS non sélectionnée ou revue sans IF – à rechercher par titre dans SIGAPS."
CASE_RE = re.compile(r'case rep|case connect', re.I)
# repérage par le titre (prudent : seuls les motifs précis validés sur l'ancien classeur)
THEMES = [
    ('Case reports', r'case rep|case connect'),
    ('Cicatrisation', r'wound|burn'),
    ('Traumatologie', r'traumatol|trauma surg|trauma acute|arch trauma|j trauma\b|injury'),
    ('Médecine militaire', r'\bmil med|military|mil health|mil veterans'),
    ('Hyperbare & plongée', r'hyperb|diving|undersea'),
    ('Chirurgie tête & cou', r'otolaryng|otorhinolaryng|rhinol|laryngol|head neck|oral maxillofac|craniomaxillofac|cleft'),
    ('Chirurgie viscérale', r'colorectal|colon rectal|dig surg|hepatobiliary|bariatr|hernia'),
]


# ---------------------------------------------------------------- lecture des exports
def parse_export(path):
    """Lit un export SIGAPS. Accepte aussi une suite d'export SANS en-tête (lignes de revues seules, puis tableau final
    « Liste des disciplines… ») : les années sont alors déduites des autres fichiers ou du classeur."""
    lines = Path(path).read_text(encoding='utf8').splitlines()
    codes, years, announced, data = {}, None, None, []
    in_footer = False
    for l in lines:
        m = re.match(r'^(?:Discipline.*:\t)?- ([A-Z0-9]{2}) : (.+)$', l)
        if m:
            codes[m.group(1)] = m.group(2).strip()
            continue
        if l.startswith('Nombre de r'):
            announced = int(re.sub(r'\D', '', l))
        elif l.startswith('connu\t'):
            years = [int(x) for x in l.split('\t')[1:] if re.fullmatch(r'\d{4}', x.strip())]
        elif l.startswith('Liste des disciplines'):
            in_footer = True
        elif in_footer:
            p = l.split('\t')
            if len(p) == 3 and re.fullmatch(r'[A-Z0-9]{2}', p[0].strip()) and p[2].strip().isdigit():
                codes[p[0].strip()] = p[1].strip()
        elif l.strip() and '\t' in l and not l.startswith(('IF\tCat.', 'NLMid')) and re.match(r'^[0-9A-Za-z]+\t', l):
            data.append([x.strip() for x in l.split('\t')])
    return dict(path=str(path), codes=codes, years=years, data=data, announced=announced)


def finalize_export(e, years):
    """Sépare lignes complètes / incomplètes une fois les années connues."""
    width = 8 + 2 * len(years)
    e['rows'] = [p for p in e['data'] if len(p) == width]
    e['broken'] = len(e['data']) - len(e['rows'])
    e['years'] = years


NLM_KEYS = {'Author(s)', 'Title Abbreviation', 'Title(s)', 'Publication Start Year', 'Publication End Year', 'Frequency', 'Country of Publication',
            'Publisher', 'Description', 'Language', 'ISSN', 'LCCN', 'Electronic Links', 'Fully Indexed In', 'In', 'Current Indexing Status',
            'Current Subset', 'MeSH', 'Publication Type(s)', 'Notes', 'Other ID', 'NLM ID', 'Coden', 'Latest Publisher', 'Selectively Indexed in',
            'Other Subject(s)', 'Edition', 'Indexes', 'Complemented by (work)', 'Mode of access', 'Terms of Use', 'SuDocs Number'}
NLM_COLS = [('Titre complet (NLM)', 'title'), ('Langue', 'lang'), ('Pays de publication', 'country'), ('MeSH (NLM)', 'mesh')]
LINK_BAD = re.compile(r'pmc\.ncbi|ncbi\.nlm|clockss|portico|europepmc|bibpurl|doi\.org|archive\.org|handle\.net|purl\.|lockss|nih\.gov|ovidsp|ebscohost', re.I)


def parse_nlm(path):
    """Notices du catalogue NLM (export texte « résumé étendu »)."""
    recs, cur, last = [], None, None
    for l in Path(path).read_text(encoding='utf8').splitlines():
        m = re.match(r'^(\d+)\. (.*)$', l)
        if m and re.match(r'^\d+\. (Author\(s\)|Title Abbreviation|Title\(s\)|Editor)', l):
            if cur:
                recs.append(cur)
            cur, l = {}, m.group(2)
        if cur is None:
            continue
        m2 = re.match(r"^([A-Z][A-Za-z()' ]{1,40}): ?(.*)$", l)
        if m2 and m2.group(1) in NLM_KEYS:
            last = m2.group(1)
            cur[last] = (cur[last] + ' ' + m2.group(2)).strip() if last in cur else m2.group(2).strip()
        elif last and l.strip():
            cur[last] += ' ' + l.strip()
    if cur:
        recs.append(cur)
    return recs


def nlm_id(r):
    return re.sub(r'\[.*$', '', r.get('NLM ID', '')).strip().rstrip('R')


def best_link(links):
    ls = [u.strip() for u in re.split(r',\s*(?=https?://)', links or '') if u.strip()]
    ok = [u for u in ls if not LINK_BAD.search(u)]
    return (ok or ls or [''])[0]


def nlm_values(r):
    langs = []
    for x in re.split(r'\s*,\s*', r.get('Language', '')):
        if x and x not in langs:
            langs.append(x)
    return {
        'title': re.sub(r'\s+', ' ', r.get('Title(s)', '')).strip().rstrip('.').strip(),
        'lang': ', '.join(langs),
        'country': r.get('Country of Publication', '').strip(),
        'mesh': r.get('MeSH', '').strip(),
    }


def shift_block(ws, start_col, by, max_row):
    """Décale vers la droite les colonnes >= start_col (valeurs, styles, fusions, largeurs)."""
    merges = [m for m in list(ws.merged_cells.ranges) if m.min_col >= start_col]
    for m in merges:
        ws.unmerge_cells(str(m))
    maxc = ws.max_column
    widths = {c: ws.column_dimensions[L(c)].width for c in range(start_col, maxc + 1)}
    ws.move_range(f"{L(start_col)}1:{L(maxc)}{max_row}", cols=by)
    for m in merges:
        ws.merge_cells(start_row=m.min_row, start_column=m.min_col + by, end_row=m.max_row, end_column=m.max_col + by)
    for c, w in widths.items():
        ws.column_dimensions[L(c + by)].width = w


def num(v):
    v = (v or '').strip().replace(',', '.')
    if not v:
        return None
    f = float(v)
    return int(f) if f == int(f) and '.' not in v else f


# ---------------------------------------------------------------- classeur
def find_header(ws):
    for r in range(1, 6):
        if any(str(c.value or '').strip().lower() == 'titre' for c in ws[r]):
            return r
    raise SystemExit("Onglet Revues : ligne d'en-tête introuvable.")


def copy_style(src, dst):
    dst._style = copy.copy(src._style)


def extend_ranges(ws, new_last_col=None):
    """Étend mises en forme conditionnelles et listes déroulantes jusqu'à LIMIT (et à la nouvelle année)."""
    old = ws.conditional_formatting
    new = ConditionalFormattingList()
    for cf in old:
        rngs = []
        for r in str(cf.sqref).split():
            rngs.append(re.sub(r'(\d+)$', str(LIMIT), r) if ':' in r else r)
        for rule in cf.rules:
            new.add(' '.join(rngs), rule)
    ws.conditional_formatting = new
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(' '.join(re.sub(r'(\d+)$', str(LIMIT), r) if ':' in r else r for r in str(dv.sqref).split()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--xlsx', required=True)
    ap.add_argument('--export', nargs='+', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--num', type=int, default=4, help="numéro d'export (colonne Source)")
    ap.add_argument('--announced', type=int, default=0, help='total annoncé par SIGAPS si le(s) fichier(s) donné(s) n\'en contien(nen)t pas')
    ap.add_argument('--nlm', nargs='*', default=[], help='exports du catalogue NLM (.txt) : liens, titre complet, langue, pays, MeSH')
    ap.add_argument('--date', default='', help='date de l\'export (jj/mm/aaaa)')
    a = ap.parse_args()

    exports = [parse_export(p) for p in a.export]
    known = [e['years'] for e in exports if e['years']]
    if known:
        years_exp = known[0]
    else:  # suite d'export sans en-tête : années déduites de la longueur des lignes + classeur
        import collections as _c
        n = _c.Counter(len(p) for e in exports for p in e['data']).most_common(1)[0][0]
        years_exp = None
        _wb = openpyxl.load_workbook(a.xlsx, read_only=True)
        _ws = _wb['Revues']
        _hdr = [c.value for c in next(_ws.iter_rows(min_row=1, max_row=1))]
        _yrs = [int(v) for v in _hdr[18:] if v not in (None, '') and re.fullmatch(r'\d{4}', str(v))]
        _wb.close()
        if len(_yrs) * 2 + 8 == n:
            years_exp = _yrs
        else:
            years_exp = list(range(_yrs[-1] - (n - 8) // 2 + 1, _yrs[-1] + 1)) if _yrs else None
        if not years_exp:
            raise SystemExit("Export sans en-tête : impossible de déterminer les années.")
        print(f"Export sans en-tête : années déduites du classeur et des lignes : {years_exp}")
    for e in exports:
        if e['years'] and e['years'] != years_exp:
            raise SystemExit('Les exports n’ont pas les mêmes années : ' + ', '.join(map(str, e['years'])))
        finalize_export(e, years_exp)
    codes = {}
    for e in exports:
        codes.update(e['codes'])
    sig = {}  # NLMid -> ligne
    for e in exports:
        for p in e['rows']:
            sig[p[0]] = p
    n_read = len(sig)
    announced = a.announced or sum(e['announced'] or 0 for e in exports)

    wb = openpyxl.load_workbook(a.xlsx)
    ws = wb['Revues']
    h1 = find_header(ws)
    first = h1 + 2
    last = ws.max_row
    while last > first and ws.cell(last, 5).value in (None, ''):
        last -= 1

    # --- années présentes dans le classeur
    sheet_years = []
    c = 19
    while re.fullmatch(r'\d{4}', str(ws.cell(h1, c).value or '').strip()):
        sheet_years.append(int(ws.cell(h1, c).value))
        c += 2
    extra_start = c if ws.cell(h1, c).value not in (None, '') else None  # colonnes supplémentaires à droite des années
    added_years = [y for y in years_exp if y not in sheet_years]
    if added_years and sheet_years and added_years[0] != sheet_years[-1] + 1:
        raise SystemExit(f'Année(s) {added_years} non contiguë(s) avec {sheet_years[-1]}.')
    if added_years and extra_start:
        shift_block(ws, extra_start, 2 * len(added_years), last)
        extra_start += 2 * len(added_years)
    for y in added_years:  # ajout des colonnes à la suite
        c0 = 19 + 2 * len(sheet_years)
        for dr in (h1, h1 + 1):
            for k in (0, 1):
                copy_style(ws.cell(dr, c0 - 2 + k), ws.cell(dr, c0 + k))
        ws.cell(h1, c0, y)
        ws.merge_cells(start_row=h1, start_column=c0, end_row=h1, end_column=c0 + 1)
        ws.cell(h1 + 1, c0, 'IF')
        ws.cell(h1 + 1, c0 + 1, 'Cat.')
        for k in (0, 1):
            ws.column_dimensions[L(c0 + k)].width = ws.column_dimensions[L(c0 - 2 + k)].width
        for r in range(first, last + 1):
            for k in (0, 1):
                copy_style(ws.cell(r, c0 - 2 + k), ws.cell(r, c0 + k))
        # mises en forme conditionnelles / listes : mêmes règles que l'année précédente
        for cf in list(ws.conditional_formatting):
            if str(cf.sqref).split()[-1].startswith(L(c0 - 1)):
                for rule in cf.rules:
                    ws.conditional_formatting.add(f'{L(c0 + 1)}{first}:{L(c0 + 1)}{LIMIT}', copy.copy(rule))
        for dv in ws.data_validations.dataValidation:
            if str(dv.sqref).split()[-1].startswith(L(c0 - 1)):
                dv.sqref = MultiCellRange(str(dv.sqref) + f' {L(c0 + 1)}{first}:{L(c0 + 1)}{LIMIT}')
        sheet_years.append(y)
    ncol_years = len(sheet_years)
    ycol = {y: 19 + 2 * k for k, y in enumerate(sheet_years)}
    last_cat_col = L(18 + 2 * ncol_years)
    old_last_cat = L(18 + 2 * (ncol_years - len(added_years)))

    fr = {}
    cs = wb['Codes_disciplines']
    for r in cs.iter_rows(min_row=2, values_only=True):
        if r[0] and r[3]:
            fr[r[0]] = r[3]
    for k, v in DISCIPLINES_FR.items():
        fr.setdefault(k, v)
    unknown = [k for k in codes if k not in fr]
    if unknown:
        raise SystemExit('Codes sans libellé français (à ajouter dans tools/disciplines_fr.py) : ' + ', '.join(unknown))

    def domains(dc):
        out = []
        for c in dc.split():
            d = fr.get(c)
            if d and d not in out:
                out.append(d)
        return ' ; '.join(out)

    exp_label = f"Export(s) SIGAPS santarm n°{a.num}"
    stats = collections.Counter()
    by_id, by_title = {}, collections.defaultdict(list)
    for r in range(first, last + 1):
        nid = ws.cell(r, 1).value
        if nid:
            by_id[str(nid)] = r
        by_title[str(ws.cell(r, 5).value).casefold()].append(r)

    def set_sigaps(r, p):
        """Écrit les valeurs SIGAPS d'une ligne export ; retourne True si quelque chose a changé."""
        ch = False
        vals = {2: p[1] or None, 3: p[2] or None, 4: p[3], 5: p[4], 6: p[5], 7: p[6], 8: p[7]}
        for col, v in vals.items():
            cell = ws.cell(r, col)
            if (cell.value or None) != v and str(cell.value or '') != str(v or ''):
                cell.value = v
                ch = True
        for k, y in enumerate(years_exp):
            for j in (0, 1):
                raw = p[8 + 2 * k + j]
                v = num(raw) if j == 0 else (raw or None)
                cell = ws.cell(r, ycol[y] + j)
                if (cell.value if cell.value != '' else None) != v:
                    cell.value = v
                    ch = True
        return ch

    def add_source(r):
        s = str(ws.cell(r, 17).value or '')
        if f'n°{a.num}' in s:
            return
        m = re.match(r"^Export\(s\) SIGAPS santarm (n°[\d, n°]+)(.*)$", s)
        if m:
            ws.cell(r, 17).value = f"Export(s) SIGAPS santarm {m.group(1)}, n°{a.num}{m.group(2)}"
        else:
            ws.cell(r, 17).value = (s + ' ; ' if s else '') + f"{exp_label}{' (' + a.date + ')' if a.date else ''}"

    new_rows = []
    for nid, p in sig.items():
        if nid in by_id:
            r = by_id[nid]
            if set_sigaps(r, p):
                stats['mises à jour'] += 1
            add_source(r)
            stats['retrouvées (déjà présentes)'] += 1
            continue
        cand = [r for r in by_title.get(p[4].casefold(), []) if str(ws.cell(r, 16).value or '').startswith('Hors export')]
        if cand:
            r = cand[0]
            set_sigaps(r, p)
            ws.cell(r, 1).value = nid
            ws.cell(r, 1).hyperlink = NLM + nid
            ws.cell(r, 16).value = str(ws.cell(r, 16).value).replace('Hors export SIGAPS', 'SIGAPS (export)')
            old_src = str(ws.cell(r, 17).value or '')
            ws.cell(r, 17).value = f"{exp_label}{' (' + a.date + ')' if a.date else ''} ; {old_src}" if old_src else exp_label
            rm = str(ws.cell(r, 18).value or '')
            ws.cell(r, 18).value = rm.replace(ABSENT, f"Retrouvée dans l'export SIGAPS n°{a.num}{' du ' + a.date if a.date else ''} (titre identique).")
            by_id[nid] = r
            stats['« Hors export » rattachées'] += 1
            continue
        new_rows.append(p)

    # --- nouvelles revues
    new_rows.sort(key=lambda p: p[4].casefold())
    tmpl = first + 1  # ligne type « SIGAPS (export) – non renseignée »
    case_tmpl = next((r for r in range(first, last + 1) if str(ws.cell(r, 15).value or '').startswith("Oui (revue de cas cliniques")), None)
    r = last
    new_idx = []
    for p in new_rows:
        r += 1
        ncols = 18 + 2 * ncol_years
        for c in range(1, ncols + 1):
            copy_style(ws.cell(tmpl, c), ws.cell(r, c))
        ws.row_dimensions[r].height = ws.row_dimensions[tmpl].height
        ws.cell(r, 1, p[0]).hyperlink = NLM + p[0]
        set_sigaps(r, p)
        ws.cell(r, 9, domains(p[6]))
        themes = [n for n, rx in THEMES if re.search(rx, p[4], re.I)]
        ws.cell(r, 10, ' ; '.join(themes) or None)
        ws.cell(r, 12, 'À vérifier')
        ws.cell(r, 13, '—')
        ws.cell(r, 14, '—')
        case = bool(CASE_RE.search(p[4]))
        ws.cell(r, 15, "Oui (revue de cas cliniques, d'après son titre)" if case else 'À vérifier')
        ws.cell(r, 16, 'SIGAPS (export) – colonnes ajoutées non renseignées')
        ws.cell(r, 17, f"{exp_label} (fichiers fournis par l'utilisateur{', ' + a.date if a.date else ''})")
        new_idx.append(r)
    stats['nouvelles revues'] = len(new_rows)
    last_row = r

    # --- domaines recalculés + remarques automatiques
    dom_changed = 0
    rows_all = list(range(first, last_row + 1))
    n_have = sum(1 for r in rows_all if str(ws.cell(r, 16).value or '').startswith('SIGAPS'))
    incomplete = bool(announced) and n_have < announced
    for r in rows_all:
        dc = str(ws.cell(r, 7).value or '')
        if dc:
            d = domains(dc)
            if d and d != (ws.cell(r, 9).value or ''):
                ws.cell(r, 9).value = d
                dom_changed += 1

    def cats(r):
        return [ws.cell(r, ycol[y] + 1).value for y in sheet_years]

    def vals(r):
        return [ws.cell(r, ycol[y]).value for y in sheet_years]

    lastyear = sheet_years[-1]
    for r in new_idx:
        rm = []
        if ws.cell(r, 4).value == 'NCI':
            rm.append(NCI_REMARK)
        cs_ = cats(r)
        if all(c in (None, '') for c in cs_):
            rm.append(f"Aucune catégorie {sheet_years[0]}-{lastyear} (dernier IF connu : {ws.cell(r, 8).value}).")
        elif cs_[-1] in (None, ''):
            rm.append(f"Pas de catégorie {lastyear} (dernier IF connu : {ws.cell(r, 8).value}).")
        if rm:
            ws.cell(r, 18, ' '.join(rm))
    # séries IF/Cat. strictement identiques (toutes années renseignées) : probable changement de titre
    groups = collections.defaultdict(list)
    for r in rows_all:
        v = vals(r)
        if all(x not in (None, '') for x in v):
            groups[tuple(zip(v, cats(r)))].append(r)
    tc = 0
    for rs in groups.values():
        if len(rs) < 2:
            continue
        for r in rs:
            rm = str(ws.cell(r, 18).value or '')
            if 'Séries IF/Cat. identiques' in rm:
                continue
            others = [str(ws.cell(o, 5).value) for o in rs if o != r]
            note = f"Séries IF/Cat. identiques à « {' », « '.join(others)} » : probable changement de titre (à vérifier)."
            ws.cell(r, 18).value = (rm + ' ' if rm else '') + note
            tc += 1
    stats['remarques « changement de titre » ajoutées'] = tc

    # --- catalogue NLM : liens, titre complet, langue, pays, MeSH
    nlm_stats = collections.Counter()
    mism = []
    if a.nlm:
        byn = {}
        for f in a.nlm:
            for rec in parse_nlm(f):
                k = nlm_id(rec)
                if k and (k not in byn or len(rec) > len(byn[k])):
                    byn[k] = rec
        endcol = 18 + 2 * ncol_years
        cols = {str(ws.cell(h1, c).value): c for c in range(endcol + 1, ws.max_column + 1) if ws.cell(h1, c).value}
        nxt = max([endcol] + list(cols.values()))
        widths = {'title': 34, 'lang': 16, 'country': 18, 'mesh': 40}
        for label, key in NLM_COLS:
            if label not in cols:
                nxt += 1
                cols[label] = nxt
                ws.cell(h1, nxt, label)
                copy_style(ws.cell(h1, 18), ws.cell(h1, nxt))
                copy_style(ws.cell(h1 + 1, 18), ws.cell(h1 + 1, nxt))
                ws.merge_cells(start_row=h1, start_column=nxt, end_row=h1 + 1, end_column=nxt)
                ws.column_dimensions[L(nxt)].width = widths[key]
        for r in rows_all:
            rec = byn.get(str(ws.cell(r, 1).value or '').strip().rstrip('R'))
            if not rec:
                continue
            nlm_stats['notices NLM rattachées'] += 1
            v = nlm_values(rec)
            for label, key in NLM_COLS:
                if v[key]:
                    cell = ws.cell(r, cols[label])
                    cell.value = v[key]
                    copy_style(ws.cell(r, 18), cell)
            link = best_link(rec.get('Electronic Links'))
            if link and not ws.cell(r, 11).value:
                ws.cell(r, 11).value = link
                ws.cell(r, 11).hyperlink = link
                ws.cell(r, 11).font = copy.copy(ws.cell(r, 1).font)
                nlm_stats['liens ajoutés'] += 1
            src = str(ws.cell(r, 17).value or '')
            if 'NLM Catalog' not in src:
                ws.cell(r, 17).value = (src + ' ; ' if src else '') + 'NLM Catalog'
            st = ws.cell(r, 4).value
            ns = 'CI' if rec.get('Current Indexing Status', '').startswith('Currently') else 'NCI'
            if st in ('CI', 'NCI') and st != ns:
                mism.append((ws.cell(r, 5).value, st, ns))
                note = f"Statut MEDLINE : SIGAPS = {st}, NLM Catalog = {ns} (à vérifier)."
                rm = str(ws.cell(r, 18).value or '')
                if 'NLM Catalog =' not in rm:
                    ws.cell(r, 18).value = (rm + ' ' if rm else '') + note
        nlm_stats['notices NLM lues (distinctes)'] = len(byn)

    # --- mise en forme / plages
    ws.auto_filter.ref = f"A{h1 + 1}:R{last_row}"
    extend_ranges(ws)

    # --- Codes_disciplines
    exist = {str(r[0].value): r[0].row for r in cs.iter_rows(min_row=2) if r[0].value}
    tmpl_row = 2
    nxt = cs.max_row
    for code, en in sorted(codes.items()):
        if code in exist:
            row = exist[code]
        else:
            nxt += 1
            row = nxt
            for c in range(1, 7):
                copy_style(cs.cell(tmpl_row, c), cs.cell(row, c))
            cs.cell(row, 1, code)
            cs.cell(row, 2, en)
        cs.cell(row, 3, 'Oui')
        if not cs.cell(row, 4).value:
            cs.cell(row, 4, fr[code])
        pres = str(cs.cell(row, 5).value or '')
        if str(a.num) not in [x.strip() for x in pres.split(',')]:
            cs.cell(row, 5, (pres + ', ' if pres else '') + str(a.num))
        cs.cell(row, 6, f'=COUNTIF(Revues!$G$3:$G${LIMIT},"*{code}*")')

    # --- formules : plages étendues (et dernière année)
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    v = c.value.replace('5857', str(LIMIT))
                    if added_years:
                        v = v.replace(f'${old_last_cat}$', f'${last_cat_col}$')
                    c.value = v

    # --- Synthèse : domaines ajoutés (bloc en bas)
    sy = wb['Synthèse']
    present = {str(sy.cell(r, 1).value) for r in range(1, sy.max_row + 1)}
    used = []
    for p in sig.values():
        for c in p[6].split():
            d = fr.get(c)
            if d and d not in used:
                used.append(d)
    todo = sorted(d for d in set(fr[c] for c in codes) if d not in present)
    if todo:
        r0 = sy.max_row + 3
        sy.cell(r0, 1, f"Autres domaines (sélection élargie, export n°{a.num})")
        copy_style(sy.cell(3, 1), sy.cell(r0, 1))
        for c in range(1, 11):
            copy_style(sy.cell(4, c), sy.cell(r0 + 1, c))
            sy.cell(r0 + 1, c).value = sy.cell(4, c).value
        for k, d in enumerate(todo):
            rr = r0 + 2 + k
            for c in range(1, 11):
                copy_style(sy.cell(5, c), sy.cell(rr, c))
                v = sy.cell(5, c).value
                sy.cell(rr, c).value = (v.replace('$A5', f'$A{rr}').replace('5857', str(LIMIT)) if isinstance(v, str) and v.startswith('=') else v)
            sy.cell(rr, 1).value = d
        if added_years:
            for c in range(1, 11):
                v = sy.cell(4, c).value
                if isinstance(v, str):
                    sy.cell(r0 + 1, c).value = v.replace(str(sheet_years[-2]), str(sheet_years[-1]))
    if added_years:  # intitulés « Cat. 2025 » -> « Cat. 2026 » dans toute la synthèse
        for row in sy.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith('=') and f'Cat. {sheet_years[-2]}' in c.value:
                    c.value = c.value.replace(f'Cat. {sheet_years[-2]}', f'Cat. {sheet_years[-1]}')
    stats['domaines ajoutés à la synthèse'] = len(todo)

    # --- Sources / Lisez-moi / Couverture
    so = wb['Sources']
    rr = so.max_row + 1
    for c in (1, 2):
        copy_style(so.cell(2, c), so.cell(rr, c))
    so.cell(rr, 1, f"Export SIGAPS santarm n°{a.num} – {len(codes)} disciplines WoS ({', '.join(Path(e['path']).name for e in exports)})")
    so.cell(rr, 2, f"Données SIGAPS (IF, catégories, disciplines, statut MEDLINE). {n_read} revues lues dans ce(s) fichier(s) ; {n_have} revues SIGAPS dans le classeur pour {announced} annoncées."
            + (" EXPORT INCOMPLET : fichier tronqué, à refaire." if incomplete else ''))
    if not incomplete:
        for r in range(2, so.max_row + 1):
            v = str(so.cell(r, 2).value or '')
            if 'EXPORT INCOMPLET : fichier tronqué, à refaire.' in v:
                so.cell(r, 2).value = v.replace('EXPORT INCOMPLET : fichier tronqué, à refaire.', f"Fichier tronqué, complété par l'export n°{a.num}.")
    lm = wb['Lisez-moi']
    for r in range(1, lm.max_row + 1):
        k = str(lm.cell(r, 1).value or '')
        if k == 'Version':
            mv = re.match(r'v(\d+)', str(lm.cell(r, 2).value or ''))
            prev = int(mv.group(1)) if mv else 6
            lm.cell(r, 2).value = (f"v{prev + 1} – {a.date or 'mise à jour'} – v{prev} + export SIGAPS n°{a.num} ({len(codes)} disciplines WoS) : "
                                   f"{stats['nouvelles revues']} revues ajoutées"
                                   + (f" ; catalogue NLM : {nlm_stats['notices NLM rattachées']} revues enrichies, {nlm_stats['liens ajoutés']} liens ajoutés" if a.nlm else '')
                                   + ('' if incomplete else f" ; SIGAPS complet ({n_have} revues)"))
        elif k == 'Exports intégrés':
            t = str(lm.cell(r, 2).value or '')
            if not incomplete:
                t = t.replace(', export tronqué : revues manquantes', f", complété par l'export n°{a.num}")
            lm.cell(r, 2).value = t + (f" ; n°{a.num} : {len(codes)} disciplines WoS ({n_read} revues lues ; {n_have} revues SIGAPS au total pour {announced} annoncées"
                                       + (", export tronqué : revues manquantes" if incomplete else '') + ").")
        elif k == 'Exhaustivité':
            t = str(lm.cell(r, 2).value or '')
            if not incomplete:
                t = t.replace(' → INCOMPLET.', f" → complété par l'export n°{a.num}.")
            lm.cell(r, 2).value = t + (f" Export n°{a.num} : {n_read} lignes lues ; {n_have} revues SIGAPS dans le classeur pour {announced} annoncées"
                                       + (" → INCOMPLET." if incomplete else " : complet."))
    cv = wb['Couverture_demande']
    for r in range(2, cv.max_row + 1):
        if 'NE non sélectionnée' in str(cv.cell(r, 4).value or ''):
            cv.cell(r, 4).value = str(cv.cell(r, 4).value) + f" [NE sélectionnée dans l'export n°{a.num}.]"

    if a.nlm:
        rr = so.max_row + 1
        for c in (1, 2):
            copy_style(so.cell(rr - 1, c), so.cell(rr, c))
        so.cell(rr, 1, 'Catalogue NLM (' + ', '.join(Path(f).name for f in a.nlm) + ')')
        so.cell(rr, 2, "Liens vers les revues (colonne Lien, jamais écrasés s'ils existent), titre complet, langue, pays, MeSH. "
                       f"{nlm_stats['notices NLM rattachées']} revues rattachées par NLMid.")
        for r in range(1, lm.max_row + 1):
            if str(lm.cell(r, 1).value or '') == 'Colonnes ajoutées':
                lm.cell(r, 2).value = str(lm.cell(r, 2).value or '') + " Colonnes « Titre complet (NLM) », « Langue », « Pays de publication », « MeSH (NLM) » et liens issus du catalogue NLM (à droite des années)."

    wb.save(a.out)

    # --- bilan
    print(f"Export(s) : {n_read} revues lues ({sum(e['broken'] for e in exports)} ligne(s) incomplète(s) ignorée(s)) ; revues SIGAPS dans le classeur : {n_have} pour {announced} annoncées")
    for k, v in stats.items():
        print(f'  {k} : {v}')
    print(f'  domaines recalculés : {dom_changed}')
    for k, v in nlm_stats.items():
        print(f'  NLM : {k} : {v}')
    if mism:
        print(f'  NLM : {len(mism)} statut(s) MEDLINE différent(s) de SIGAPS : ' + '; '.join(f'{t} (SIGAPS {a_}, NLM {b_})' for t, a_, b_ in mism))
    if added_years:
        print(f'  nouvelle(s) année(s) : {added_years}')
    print(f"Classeur : {last_row - first + 1} revues -> {a.out}")
    if incomplete:
        print(f"ATTENTION : export incomplet ({announced - n_have} revues manquantes). Refaire l'export en plusieurs fichiers puis relancer avec tous les fichiers.")


if __name__ == '__main__':
    main()
