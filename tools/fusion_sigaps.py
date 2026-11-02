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
    lines = Path(path).read_text(encoding='utf8').splitlines()
    codes, start, years, announced = {}, None, [], None
    for i, l in enumerate(lines):
        m = re.match(r'^(?:Discipline.*:\t)?- ([A-Z0-9]{2}) : (.+)$', l)
        if m and start is None:
            codes[m.group(1)] = m.group(2).strip()
        if l.startswith('Nombre de r'):
            announced = int(re.sub(r'\D', '', l))
        if l.startswith('connu\t'):
            years = [int(x) for x in l.split('\t')[1:] if re.fullmatch(r'\d{4}', x.strip())]
        if l.startswith('IF\tCat.'):
            start = i + 1
            break
    if start is None or not years:
        raise SystemExit(f"{path} : en-tête SIGAPS non reconnu (lignes « connu\\t2021… » et « IF\\tCat. » attendues).")
    rows, broken = [], 0
    for l in lines[start:]:
        if not l.strip():
            continue
        p = [x.strip() for x in l.split('\t')]
        if len(p) != 8 + 2 * len(years):
            broken += 1
            continue
        rows.append(p)
    return dict(path=str(path), codes=codes, years=years, rows=rows, broken=broken, announced=announced)


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
    ap.add_argument('--date', default='', help='date de l\'export (jj/mm/aaaa)')
    a = ap.parse_args()

    exports = [parse_export(p) for p in a.export]
    years_exp = exports[0]['years']
    for e in exports:
        if e['years'] != years_exp:
            raise SystemExit('Les exports n’ont pas les mêmes années : ' + ', '.join(map(str, e['years'])))
    codes = {}
    for e in exports:
        codes.update(e['codes'])
    sig = {}  # NLMid -> ligne
    for e in exports:
        for p in e['rows']:
            sig[p[0]] = p
    n_read = len(sig)
    announced = sum(e['announced'] or 0 for e in exports)

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
    while ws.cell(h1, c).value not in (None, ''):
        sheet_years.append(int(ws.cell(h1, c).value))
        c += 2
    added_years = [y for y in years_exp if y not in sheet_years]
    if added_years and sheet_years and added_years[0] != sheet_years[-1] + 1:
        raise SystemExit(f'Année(s) {added_years} non contiguë(s) avec {sheet_years[-1]}.')
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
    so.cell(rr, 2, f"Données SIGAPS (IF, catégories, disciplines, statut MEDLINE). {n_read} revues lues sur {announced} annoncées."
            + (" EXPORT INCOMPLET : fichier tronqué, à refaire." if n_read < announced else ''))
    lm = wb['Lisez-moi']
    for r in range(1, lm.max_row + 1):
        k = str(lm.cell(r, 1).value or '')
        if k == 'Version':
            lm.cell(r, 2).value = (f"v7 – {a.date or 'mise à jour'} – v6 + export SIGAPS n°{a.num} ({len(codes)} disciplines WoS) : "
                                   f"{stats['nouvelles revues']} revues ajoutées")
        elif k == 'Exports intégrés':
            lm.cell(r, 2).value = str(lm.cell(r, 2).value or '') + (f" ; n°{a.num} : {len(codes)} disciplines WoS ({n_read} revues lues sur {announced} annoncées"
                                                                    + (", export tronqué : revues manquantes" if n_read < announced else '') + ").")
        elif k == 'Exhaustivité':
            lm.cell(r, 2).value = str(lm.cell(r, 2).value or '') + (f" Export n°{a.num} : {n_read} lignes lues pour {announced} annoncées"
                                                                    + (" → INCOMPLET." if n_read < announced else " : complet."))
    cv = wb['Couverture_demande']
    for r in range(2, cv.max_row + 1):
        if 'NE non sélectionnée' in str(cv.cell(r, 4).value or ''):
            cv.cell(r, 4).value = str(cv.cell(r, 4).value) + f" [v7 : NE sélectionnée dans l'export n°{a.num}.]"

    wb.save(a.out)

    # --- bilan
    print(f"Export(s) : {n_read} revues lues sur {announced} annoncées ({sum(e['broken'] for e in exports)} ligne(s) incomplète(s) ignorée(s))")
    for k, v in stats.items():
        print(f'  {k} : {v}')
    print(f'  domaines recalculés : {dom_changed}')
    if added_years:
        print(f'  nouvelle(s) année(s) : {added_years}')
    print(f"Classeur : {last_row - first + 1} revues -> {a.out}")
    if n_read < announced:
        print(f"ATTENTION : export incomplet ({announced - n_read} revues manquantes). Refaire l'export en plusieurs fichiers puis relancer avec tous les fichiers.")


if __name__ == '__main__':
    main()
