#!/usr/bin/env python3
"""Aufbereitung der Hessischen Kreiszahlen für das Dashboard des Bildungsbüros Lahn-Dill-Kreis.

Aufruf (aus dem Hauptordner des Repositorys):
    python3 werkzeuge/kreiszahlen.py daten   # liest quellen/*.xlsx, schreibt daten.json und werkzeuge/ausgabe/*.csv + bericht.txt
    python3 werkzeuge/kreiszahlen.py seite   # baut index.html (und gemeinden.html) aus werkzeuge/vorlage.html, kacheln*.json, bezugsgroessen*.json und d3.min.js

Grundsätze:
- Referenz für Tabellen- und Spaltennummern (mid = "<Tabelle>_<Spalte>") ist die Ausgabe Herbst 2026 (REFERENZ).
  Jede Ausgabe – ältere wie künftige – wird über Tabellentitel und Spaltenköpfe auf diese Merkmale abgebildet.
  Neue Tabellen oder Spalten, die es in der Referenz nicht gibt, fallen weg und stehen im Bericht.
- Merkmalskatalog und Regionen: werkzeuge/katalog/merkmale.csv und regionen.csv.
- Geheim gehaltene oder fehlende Werte (".", "x", "/", "•" …) bleiben leer, "—" wird 0. Dezimalwerte auf 4 Stellen.
- Dateinamen in quellen/: HKr_<JJ>-1hj*.xlsx bzw. HKr_<JJ>_1*.xlsx = Frühjahr 20JJ, HKr_<JJ>-2hj*.xlsx = Herbst 20JJ.
"""
import collections, csv, difflib, glob, json, os, re, sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(ROOT, 'werkzeuge')
QUELLEN = os.path.join(ROOT, 'quellen')
REFERENZ = 'HKr_26-2hj.xlsx'
MAX_TAB = 66  # Tabellen 67/68 sind Verzeichnisse, keine Kreiswerte

# ---------------------------------------------------------------- Einlesen
def clean(s):
    if s is None: return ''
    s = str(s).replace('\xad', '')
    s = re.sub(r'-\s*\n\s*', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    for _ in range(3): s = re.sub(r'((?:19|20)\d{2})\d\)', r'\1', s)
    s = re.sub(r'\s*\d+\)\s*$', '', s)
    s = re.sub(r'(\D)\d\)', r'\1', s)
    return s.strip()

MARKER = {'—', '–', '-', '/', '.', 'x', '•'}

def merge_fragments(path):
    # "7 › bis unter › 18" (über drei Kopfzeilen verteilt) → "7 bis unter 18"
    out = []
    i = 0
    while i < len(path):
        if i + 2 < len(path) and path[i + 1] == 'bis unter':
            out.append(f'{path[i]} bis unter {path[i + 2]}'); i += 3
        else:
            out.append(path[i]); i += 1
    return out

def markers(v):
    """Fußnotenzeichen in einer Kopf- oder Titelzelle: "Hochschulreife2)" → [2], "20251)2)" → [1, 2]."""
    if v is None: return []
    t = str(v)
    out = []
    t2 = t
    while True:  # Jahreszahl mit angehängter Fußnote, z. B. 20251)
        m = re.search(r'((?:19|20)\d{2})(\d)\)', t2)
        if not m: break
        out.append(int(m.group(2))); t2 = t2[:m.start()] + m.group(1) + t2[m.end():]
    out += [int(x) for x in re.findall(r'(?<![\d.])(\d)\)', t2)]
    return sorted(set(out))

def notes_text(rows, start):
    """Fußnoten und Quellenangaben unter einer Tabelle (Zeilen nach den Daten)."""
    fn, quelle = {}, []
    for r in rows[start:]:
        txt = ' '.join(str(c).strip() for c in r if isinstance(c, str) and c.strip())
        txt = re.sub(r'\s+', ' ', txt).strip()
        if not txt or set(txt) <= set('_ '): continue
        if txt.startswith('Quelle'): quelle.append(txt); continue
        parts = re.split(r'(?:^|\s*—\s*)(\d+)\)\s*', txt)
        if len(parts) > 1:
            for i in range(1, len(parts) - 1, 2):
                fn[int(parts[i])] = parts[i + 1].strip().rstrip('—').strip()
        elif fn:
            last = max(fn); fn[last] = (fn[last] + ' ' + txt).strip()
        else:
            quelle.append(txt)
    return fn, quelle

def parse(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for name in wb.sheetnames:
        if not name.isdigit(): continue
        ws = wb[name]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        title = re.sub(r'^\d+\.\s*', '', clean(rows[0][0]) if rows and rows[0] else '')
        title_fn = markers(rows[0][0]) if rows and rows[0] else []
        hdr0 = [clean(x) for x in rows[1]]
        has_ags = 'Gebiets' in (hdr0[0] or '') or 'schlüssel' in (hdr0[0] or '')
        namecol = 1 if has_ags else 0
        first = None
        for i in range(2, min(len(rows), 15)):
            r = rows[i]
            if has_ags:
                if r[0] is not None and re.fullmatch(r'\d{6}', str(r[0]).strip()): first = i; break
            elif r[0] and any(isinstance(x, (int, float)) or (isinstance(x, str) and x.strip() in MARKER) for x in r[1:]): first = i; break
        if first is None:
            out[int(name)] = dict(title=title, cols={}, rows=[], has_ags=has_ags); continue
        grid = {(i, j): v for i in range(1, first) for j, v in enumerate(rows[i])}
        for mr in ws.merged_cells.ranges:
            if mr.min_row - 1 >= first or mr.max_row - 1 < 1: continue
            v = ws.cell(mr.min_row, mr.min_col).value
            for j in range(mr.min_col - 1, mr.max_col): grid[(mr.min_row - 1, j)] = v
        ncol = max(len(r) for r in rows[1:first + 1])
        off = 0 if has_ags else 1
        cols, colfn = {}, {}
        for j in range(namecol + 1, ncol):
            path, mk = [], set()
            for i in range(1, first):
                t = clean(grid.get((i, j)))
                mk.update(markers(grid.get((i, j))))
                if t and (not path or path[-1] != t): path.append(t)
            if path and path != ['Zurück zum Inhalt']:
                cols[j + 1 + off] = merge_fragments(path)
                if mk: colfn[j + 1 + off] = sorted(mk)
        data, last = [], first
        for k, r in enumerate(rows[first:], first):
            if r and r[0] is not None and (re.fullmatch(r'\d{6}', str(r[0]).strip()) or (not has_ags and any(isinstance(x, (int, float)) for x in r[1:]))): last = k
        for r in rows[first:]:
            if r[0] is None: continue
            nm = clean(r[namecol]) if len(r) > namecol else ''
            if has_ags:
                k = str(r[0]).strip()
                if not re.fullmatch(r'\d{6}', k): continue
            else:
                k = None
                if not nm or not any(x is not None and x != '' for x in r[1:]): continue
                if str(r[0]).lstrip().startswith(('Quelle', '_')) or re.match(r'\s*\d+\)', str(r[0])): continue
            data.append(dict(ags=k, name=nm, vals={j + 1 + off: r[j] for j in range(namecol + 1, len(r)) if j + 1 + off in cols}))
        fn, quelle = notes_text(rows, last + 1)
        out[int(name)] = dict(title=title, cols=cols, rows=data, has_ags=has_ags, title_fn=title_fn, colfn=colfn, fn=fn, quelle=quelle)
    # Zeichenerklärung und allgemeine Hinweise (Blatt „Impressum“)
    zeichen, hinweise = [], []
    if 'Impressum' in wb.sheetnames:
        mode = None
        for r in wb['Impressum'].iter_rows(values_only=True):
            cells = [str(c).strip() for c in r if c is not None and str(c).strip()]
            if not cells: continue
            if cells[0].startswith('Zeichenerklärung'): mode = 'z'; continue
            if cells[0].startswith('Hinweise'): mode = 'h'; continue
            if mode == 'z':
                if len(cells) >= 3 and cells[1] == '=': zeichen.append([cells[0], ' '.join(cells[2:])])
                elif zeichen: zeichen[-1][1] += ' ' + ' '.join(cells)
            elif mode == 'h':
                hinweise += [re.sub(r'\s+', ' ', x).strip() for x in ' '.join(cells).split('\n') if x.strip()]
    out['_allgemein'] = dict(zeichen=zeichen, hinweise=hinweise)
    return out

def ausgabe_label(fn):
    m = re.match(r'HKr_(\d{2})[_-](\d)', os.path.basename(fn))
    if not m: raise SystemExit(f'Dateiname nicht erkannt: {fn} (erwartet HKr_JJ-1hj… oder HKr_JJ-2hj…)')
    jahr, hj = 2000 + int(m.group(1)), int(m.group(2))
    return (jahr, hj), ('Frühjahr ' if hj == 1 else 'Herbst ') + str(jahr)

# ---------------------------------------------------------------- Abgleich mit der Referenz
MON = 'Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember'
def unfoot(t): return re.sub(r'((?:19|20)\d{2})\d\)', r'\1', t)
def norm_title(t):
    raw = t
    t = unfoot(t)
    t = re.sub(r'\(in Hektar\)|\(Angaben in [^)]*\)|\(Stand:[^)]*\)?', ' ', t)
    t = re.sub(r'\s-\s.*$', ' ', t)
    t = re.sub(r'\b(im|am)\s+(1\.|2\.)\s*Halbjahr\s+(\d{4})?', ' ', t)
    t = re.sub(r'\b(im Jahresdurchschnitt|im Jahr|am|im|zum|Ende des Schuljahres)?\s*(\d{1,2}\.\s*)?(%s)?\s*(19|20)\d{2}(/\d{2})?\b' % MON, ' ', t)
    t = re.sub(r'\bund(\s+und)*\s*(?=$|nach)', ' ', t)
    t = re.sub(r'\bin Hessen\b', ' ', t)
    t = re.sub(r'[^\wäöüß]+', ' ', t).lower()
    s = re.sub(r'\s+', ' ', t).strip()
    s = re.sub(r'(?<=[a-zäöüß])\d+\b', '', s)
    # bekannte Umbenennungen zwischen den Ausgaben
    s = s.replace('finanzen der gemeinden und gemeindeverbände', 'finanzen der kreisfreien städte und gesamtkreise')
    s = s.replace(' im 1 halbjahr', '').replace(' am 31 dezember', '').replace(' in tausend', '')
    s = s.replace('steuerbaren umsätzen', 'umsätzen').replace('nach grundstücksarten', 'nach ausgewählten grundstücksarten')
    s = s.replace('ohne gebäude und ohne inventar', 'ohne weitere gegenstände rechte und pflichten reine geldleistung')
    if re.search(r'und im Jahr', raw): s += ' +halbjahr+jahr'
    return s.strip()

def norm_path(path):
    s = ' › '.join(x.replace('%', ' prozent ') for x in path if re.search(r'[\w%]', x))
    s = unfoot(s)
    s = re.sub(r'\d{1,2}\.\d{1,2}\.(19|20)\d{2}', 'DATUM', s)
    s = re.sub(r'\b(19|20)\d{2}(/\d{2})?\b', 'JAHR', s)
    s = re.sub(r'\b(und zwar|davon|darunter)\b', ' ', s)
    s = re.sub(r'[^\wäöüß›]+', ' ', s).lower()
    # gleichbedeutende Formulierungen zwischen den Ausgaben
    s = re.sub(r'empfängerinnen und empfänger|empfänger innen', 'empfänger', s)
    s = re.sub(r'je beschäftigte n|je beschäftigter person|je tätiger person', 'je person', s)
    return re.sub(r'\s+', ' ', s).strip()

STOP = {'insgesamt', 'davon', 'darunter', 'und', 'zwar', 'in', 'im', 'am', 'der', 'die', 'das', 'mit', 'je', 'zusammen', 'waren', 'alt'}
def toks(p):
    x = norm_path(p) if isinstance(p, list) else p
    return [w for w in re.split(r'[\s›]+', x) if w and w not in STOP]
def sim(a, b):
    ta, tb = toks(a), toks(b)
    full = difflib.SequenceMatcher(None, ' '.join(sorted(ta)), ' '.join(sorted(tb))).ratio()
    a = [x for x in a if toks([x])] or a; b = [x for x in b if toks([x])] or b
    la, lb = toks(a[-1:]), toks(b[-1:])
    leaf = difflib.SequenceMatcher(None, ' '.join(sorted(la)), ' '.join(sorted(lb))).ratio() if (la or lb) else 1.0
    if la and lb and set(la) == set(lb): leaf = 1.0
    if ta and tb and sorted(ta) == sorted(tb): return 1.0
    return 0.55 * full + 0.45 * leaf

FIXDATE = re.compile(r'\b(\d{1,2}\.\d{1,2}\.)((?:19|20)\d{2})\b')
def fixed_dates(path):
    """Feste Bezugsdaten wie Zensusstichtage (nicht 31.12. oder 30.06.), die sich zwischen Ausgaben nicht verschieben dürfen."""
    return {a + b for a, b in FIXDATE.findall(' '.join(path)) if a not in ('31.12.', '30.06.', '30.6.', '1.1.', '01.01.')}

def strip_prefix(cols):
    """Entfernt eine erste Kopfzeile, die (fast) alle Spalten einer Tabelle teilen, z. B. "Bodenfläche › …"."""
    if len(cols) < 3: return cols
    firsts = collections.Counter(p[0] for p in cols.values() if len(p) > 1)
    if not firsts: return cols
    top, n = firsts.most_common(1)[0]
    if n < 0.7 * len(cols): return cols
    return {c: (p[1:] if len(p) > 1 and p[0] == top else p) for c, p in cols.items()}

def align(oc, cc, thr=0.62, strict=0.8):
    """Zuordnung der Spalten: erst ordnungserhaltend (Needleman-Wunsch über die Ähnlichkeit der Spaltenköpfe),
    dann für umgestellte Spalten ohne Reihenfolge, aber nur bei hoher Ähnlichkeit."""
    fd_o = {c: fixed_dates(p) for c, p in oc.items()}; fd_c = {c: fixed_dates(p) for c, p in cc.items()}
    so, scc = strip_prefix(oc), strip_prefix(cc)
    def S_(a, b):
        if fd_c[b] and fd_c[b] != fd_o[a]: return 0.0
        return max(sim(oc[a], cc[b]), sim(so[a], scc[b]))
    A, Bc = sorted(oc), sorted(cc)
    n, m = len(A), len(Bc)
    S = [[S_(A[i], Bc[j]) for j in range(m)] for i in range(n)]
    D = [[0.0] * (m + 1) for _ in range(n + 1)]; bt = [[None] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            best, how = D[i - 1][j], 'u'
            if D[i][j - 1] > best: best, how = D[i][j - 1], 'l'
            if S[i - 1][j - 1] >= thr and D[i - 1][j - 1] + S[i - 1][j - 1] > best: best, how = D[i - 1][j - 1] + S[i - 1][j - 1], 'd'
            D[i][j] = best; bt[i][j] = how
    i, j = n, m; res, scores = {}, {}
    while i > 0 and j > 0:
        h = bt[i][j]
        if h == 'd': res[A[i - 1]] = Bc[j - 1]; scores[A[i - 1]] = S[i - 1][j - 1]; i -= 1; j -= 1
        elif h == 'u': i -= 1
        else: j -= 1
    # zweiter Durchgang: umgestellte Spalten
    pairs = sorted(((S[i][j], A[i], Bc[j]) for i in range(n) for j in range(m) if A[i] not in res and Bc[j] not in res.values() and S[i][j] >= strict), reverse=True)
    for sc_, a, b in pairs:
        if a in res or b in res.values(): continue
        res[a] = b; scores[a] = sc_
    return res, scores

# Geprüfte Einzelfälle, die die automatische Zuordnung nicht erkennt: Merkmal → Muster des Spaltenkopfs in älteren Ausgaben
OVERRIDES = {
    '46_3': r'^Empfängerhaushalte am 31\. Dezember \d{4} insgesamt › Anzahl der Haushalte$',
    '48_4': r'^(Empfänger/-innen nach Sitz des Trägers insgesamt|Empfängerinnen und Empfänger nach Sitz des Trägers › insgesamt) › Anzahl$',
}

def conv(v):
    if v is None: return ''
    if isinstance(v, (int, float)):
        f = float(v); return str(int(f)) if f.is_integer() else ('%.4f' % f).rstrip('0').rstrip('.')
    s = str(v).strip().replace('‬', '').replace('‪', '')
    if s in ('—', '-', '–'): return '0'
    try:
        f = float(s.replace(' ', '').replace(',', '.')); return str(int(f)) if f.is_integer() else ('%.4f' % f).rstrip('0').rstrip('.')
    except ValueError:
        return ''

# ---------------------------------------------------------------- Berichtszeitraum
MONL = MON.split('|')
def period(text):
    t = re.sub(r'\s+', ' ', text)
    m = re.search(r'Schuljahr(?:es)? ((?:19|20)\d{2}/\d{2})', t)
    if m: return ('Schuljahr ' + m.group(1), m.group(1), 'Schuljahr')
    m = re.search(r'([12])\. Halbjahr ((?:19|20)\d{2})', t)
    if m: return (f'{m.group(1)}. Halbjahr {m.group(2)}', m.group(2), f'{m.group(1)}. Halbjahr')
    m = re.search(r'(\d{1,2})\. (%s) ((?:19|20)\d{2})' % MON, t)
    if m: return (f'{int(m.group(1))}. {m.group(2)} {m.group(3)}', m.group(3), f'Stichtag {int(m.group(1))}. {m.group(2)}')
    m = re.search(r'(\d{1,2})\.(\d{1,2})\.\s?((?:19|20)\d{2})', t)
    if m:
        mo = MONL[int(m.group(2)) - 1]
        return (f'{int(m.group(1))}. {mo} {m.group(3)}', m.group(3), f'Stichtag {int(m.group(1))}. {mo}')
    m = re.search(r'Jahresdurchschnitt ((?:19|20)\d{2})', t)
    if m: return ('Jahresdurchschnitt ' + m.group(1), m.group(1), 'Jahresdurchschnitt')
    ys = re.findall(r'\b((?:19|20)\d{2})\b', t)
    if ys: return (ys[-1], ys[-1], 'Jahr')
    return None

def bezug_for(mid, title, path, table_stichtag=None):
    if re.search(r'[Ww]ährend des Jahres', path): table_stichtag = None  # Bewegungen im Jahr, kein Stichtag
    p = period(re.split(r'gegenüber', path)[0]) or table_stichtag or period(title)
    # Tabellen 9/10 sind Bewegungen im Jahr bzw. Halbjahr, auch wenn der Titel einen Stichtag nennt
    if p and mid.split('_')[0] in ('9', '10') and p[2].startswith('Stichtag'):
        p = (p[1], p[1], 'Jahr') if '31. Dezember' in p[2] else ('1. Halbjahr ' + p[1], p[1], '1. Halbjahr')
    return p

# ---------------------------------------------------------------- Befehl: daten
def cmd_daten():
    cat = list(csv.DictReader(open(os.path.join(TOOLS, 'katalog', 'merkmale.csv'), encoding='utf-8')))
    reg = sorted(csv.DictReader(open(os.path.join(TOOLS, 'katalog', 'regionen.csv'), encoding='utf-8')), key=lambda r: int(r['reihenfolge']))
    mids = {r['mid'] for r in cat}
    files = sorted(glob.glob(os.path.join(QUELLEN, 'HKr_*.xlsx')), key=lambda f: ausgabe_label(f)[0])
    if not any(os.path.basename(f) == REFERENZ for f in files): raise SystemExit('Referenzausgabe fehlt in quellen/: ' + REFERENZ)
    parsed = {os.path.basename(f): parse(f) for f in files}
    allgemein = {fn: d.pop('_allgemein', {}) for fn, d in parsed.items()}
    newest = os.path.basename(files[-1])
    notes = {}  # Fußnoten und Quellen der neuesten Ausgabe je Tabelle der Referenz
    cur = parsed[REFERENZ]
    inv = {norm_title(d['title']): t for t, d in cur.items() if t <= MAX_TAB}
    valid = {r['ags'] for r in reg}
    name2ags = {}
    for d in cur.values():
        for r in d['rows']:
            if r['ags']: name2ags.setdefault(re.sub(r'\W+', '', r['name'].lower()), r['ags'])
    for r in reg: name2ags.setdefault(re.sub(r'\W+', '', r['region'].lower()), r['ags'])

    werte, bezug, bericht = [], [], []
    labels = []
    for f in files:
        fn = os.path.basename(f); _, lab = ausgabe_label(f); labels.append(lab)
        bericht.append(f'== {lab} ({fn})')
        for t, d in sorted(parsed[fn].items()):
            ct = (t if t <= MAX_TAB else None) if fn == REFERENZ else inv.get(norm_title(d['title']))
            if ct is None:
                if t <= MAX_TAB or fn != REFERENZ: bericht.append(f'   Tabelle {t} ohne Gegenstück in der Referenz, entfällt: {d["title"][:100]}')
                continue
            oc = d['cols']; cc = {c: p for c, p in cur[ct]['cols'].items() if f'{ct}_{c}' in mids}
            cmap, sc = align(oc, cc) if fn != REFERENZ else ({c: c for c in oc if c in cc}, {c: 1.0 for c in oc})
            if fn != REFERENZ:
                for mid_, pat in OVERRIDES.items():
                    t_, c_ = map(int, mid_.split('_'))
                    if t_ != ct or c_ in cmap.values(): continue
                    hit = [c for c in oc if c not in cmap and re.match(pat, ' › '.join(oc[c]))]
                    if len(hit) == 1: cmap[hit[0]] = c_; sc[hit[0]] = 1.0
            un_o = [' › '.join(oc[c]) for c in sorted(oc) if c not in cmap]
            un_c = [f'{ct}_{c}' for c in sorted(cc) if c not in cmap.values()]
            weak = [(' › '.join(oc[a]), ' › '.join(cc[b]), round(sc[a], 2)) for a, b in cmap.items() if sc[a] < 0.8]
            if un_o: bericht.append(f'   T{t}→T{ct}: Spalten ohne Gegenstück: {un_o}')
            if un_c: bericht.append(f'   T{t}→T{ct}: Merkmale ohne Wert in dieser Ausgabe: {un_c}')
            for w in weak: bericht.append(f'   T{t}→T{ct}: unsichere Zuordnung {w}')
            if fn == newest:
                spalten = {}
                for c, m in cmap.items():
                    nrs = sorted(set(d.get('colfn', {}).get(c, [])))
                    if nrs: spalten[f'{ct}_{m}'] = nrs
                notes[ct] = dict(titel=d.get('title_fn', []), fn={str(k): v for k, v in sorted(d.get('fn', {}).items())}, quelle=d.get('quelle', []), spalten=spalten)
                missing = [n for n in set(d.get('title_fn', [])) | {x for v in spalten.values() for x in v} if n not in d.get('fn', {})]
                if missing: bericht.append(f'   T{t}→T{ct}: Fußnotenzeichen ohne Text: {sorted(missing)}')
            for r in d['rows']:
                ags = r['ags'] or name2ags.get(re.sub(r'\W+', '', r['name'].lower()))
                if not ags or ags not in valid: continue
                for c, v in r['vals'].items():
                    if c in cmap: werte.append((lab, f'{ct}_{cmap[c]}', ags, conv(v)))
            tst = None
            for pth in oc.values():
                q = period(re.split(r'gegenüber', ' › '.join(pth))[0])
                if q and q[2].startswith('Stichtag'): tst = q; break
            for c, m in cmap.items():
                p = bezug_for(f'{ct}_{m}', d['title'], ' › '.join(oc[c]), tst)
                if p: bezug.append(dict(ausgabe=lab, mid=f'{ct}_{m}', bezug=p[0], jahr=p[1], art=p[2]))
                else: bericht.append(f'   {ct}_{m}: kein Berichtszeitraum erkennbar')

    # Plausibilität: Sprünge des Landeswerts zur gleichen Ausgabe des Vorjahres
    land = collections.defaultdict(dict)
    for lab, mid, ags, v in werte:
        if ags == '000000' and v != '': land[mid][lab] = float(v)
    catn = {r['mid']: r['merkmal'] for r in cat}
    bericht.append('== Auffällige Sprünge des Landeswerts (Faktor > 2,5 zur gleichen Ausgabe des Vorjahres)')
    for mid, d in land.items():
        for season in ('Frühjahr', 'Herbst'):
            seq = [(l, d[l]) for l in labels if l.startswith(season) and l in d]
            for (l1, a), (l2, b) in zip(seq, seq[1:]):
                if a and b and (b / a > 2.5 or b / a < 0.4) and max(abs(a), abs(b)) > 5:
                    bericht.append(f'   {mid} {l1}: {a:g} → {l2}: {b:g}  ({catn.get(mid, "")[:70]})')

    ordn = {lab: i for i, lab in enumerate(labels)}
    catpos = {r['mid']: i for i, r in enumerate(cat)}
    regpos = {r['ags']: i for i, r in enumerate(reg)}
    werte.sort(key=lambda r: (ordn[r[0]], catpos.get(r[1], 10 ** 6), regpos.get(r[2], 10 ** 6)))
    bezug.sort(key=lambda r: (ordn[r['ausgabe']], catpos.get(r['mid'], 10 ** 6)))
    dupl = [k for k, n in collections.Counter((a, m, g) for a, m, g, _ in werte).items() if n > 1]
    if dupl: raise SystemExit(f'Doppelte Werte, Abbruch: {dupl[:5]}')

    os.makedirs(os.path.join(TOOLS, 'ausgabe'), exist_ok=True)
    with open(os.path.join(TOOLS, 'ausgabe', 'werte.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh, lineterminator='\n'); w.writerow(['ausgabe', 'mid', 'ags', 'wert']); w.writerows(werte)
    with open(os.path.join(TOOLS, 'ausgabe', 'bezug.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=['ausgabe', 'mid', 'bezug', 'jahr', 'art'], lineterminator='\n'); w.writeheader(); w.writerows(bezug)
    open(os.path.join(TOOLS, 'ausgabe', 'bericht.txt'), 'w', encoding='utf-8').write('\n'.join(bericht) + '\n')

    # daten.json für die Seite
    ri = {r['ags']: i for i, r in enumerate(reg)}
    def num(v):
        if v == '': return None
        f = float(v); return int(f) if f.is_integer() else f
    W = {lab: {} for lab in labels}
    for lab, mid, ags, v in werte:
        W[lab].setdefault(mid, [None] * len(reg))[ri[ags]] = num(v)
    for lab in labels: W[lab] = {m: a for m, a in W[lab].items() if any(x is not None for x in a)}
    BZ = {lab: {} for lab in labels}
    for r in bezug: BZ[r['ausgabe']][r['mid']] = [r['bezug'], r['jahr'], r['art']]
    themen, tabs = [], {}
    for r in cat:
        if r['thema'] not in themen: themen.append(r['thema'])
        tabs.setdefault(r['tab'], [themen.index(r['thema']), r['tabelle']])
    D = dict(ausgaben=labels, regionen=[[r['ags'], r['region'], r['ebene']] for r in reg], themen=themen, merkmale=[[r['mid'], r['merkmal'], r.get('einheit', '')] for r in cat],
             werte=[W[l] for l in labels], bezug=[BZ[l] for l in labels],
             quelle=f'Hessisches Statistisches Landesamt, Hessische Kreiszahlen, Excel-Ausgaben {labels[0]} bis {labels[-1]}', tabellen=tabs,
             hinweise={str(k): v for k, v in sorted(notes.items())}, erlaeuterungen=dict(ausgabe=labels[-1], **allgemein.get(newest, {})))
    open(os.path.join(ROOT, 'daten.json'), 'w', encoding='utf-8').write(json.dumps(D, ensure_ascii=False, separators=(',', ':')))
    # Erläuterungen als Tabelle für das Claude-Dashboard
    rows = []
    for k, v in sorted(notes.items()):
        for n, txt in v['fn'].items():
            mids = '*' if int(n) in v['titel'] else ','.join(m for m, nrs in v['spalten'].items() if int(n) in nrs)
            rows.append(dict(id=f'{k}-{n}', tab=str(k), art='fussnote', nr=n, text=txt, mids=mids))
        for i, q in enumerate(v['quelle'], 1):
            rows.append(dict(id=f'{k}-q{i}', tab=str(k), art='quelle', nr='', text=q, mids='*'))
    E = allgemein.get(newest, {})
    for i, h in enumerate(E.get('hinweise', []), 1): rows.append(dict(id=f'h{i}', tab='', art='hinweis', nr=str(i), text=h, mids=''))
    for i, (z, t) in enumerate(E.get('zeichen', []), 1): rows.append(dict(id=f'z{i}', tab='', art='zeichen', nr=z, text=t, mids=''))
    json.dump(dict(ausgabe=labels[-1], rows=rows), open(os.path.join(TOOLS, 'ausgabe', 'hinweise.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
    print(f'{len(labels)} Ausgaben, {len(werte)} Werte, {len(bezug)} Berichtszeiträume. Bericht: werkzeuge/ausgabe/bericht.txt')

# ---------------------------------------------------------------- Befehl: seite
# Zwei Seiten aus derselben Vorlage: Kreiszahlen (index.html) und Gemeindezahlen (gemeinden.html).
SCHALTER = [('kreis', 'index.html', './', 'Kreiszahlen'), ('gemeinde', 'gemeinden.html', 'gemeinden.html', 'Gemeindezahlen')]
SEITEN = {
    'kreis': dict(
        datei='index.html', daten='daten.json', kacheln='kacheln.json', basis='bezugsgroessen.json', spanne=lambda D: f'{D["ausgaben"][0]} bis {D["ausgaben"][-1]}',
        TITEL='Kreiszahlen Lahn-Dill-Kreis',
        BESCHREIBUNG='Bildungs- und Strukturdaten des Lahn-Dill-Kreises im Vergleich der hessischen Kreise, nach den Hessischen Kreiszahlen des Statistischen Landesamts.',
        LEDE='Hessische Kreiszahlen des Statistischen Landesamts, Ausgabe',
        TAKE='<span class="rn" data-slot="fokus">—</span>: <b class="kv" data-k="abi" data-slot="fokus">—</b> der Schulentlassenen verlassen die Schule mit allgemeiner Hochschulreife, <b class="kv" data-k="ohneHSA" data-slot="fokus">—</b> ohne Hauptschulabschluss (<span class="rn" data-slot="vgl">—</span>: <span class="kv" data-k="abi" data-slot="vgl">—</span> bzw. <span class="kv" data-k="ohneHSA" data-slot="vgl">—</span>).',
        KACHEL_NOTE='Kennzahlen zum oben gewählten Thema, Werte der neuesten Ausgabe. Rang 1 = höchster Wert unter den Landkreisen und kreisfreien Städten.',
        VERGLEICH_H2='Jedes Merkmal im Kreisvergleich',
        VERGLEICH_NOTE='Alle rund 800 Merkmale aus 66 Tabellen. Wählen Sie Thema, Tabelle und Merkmal. Ein Klick auf einen Balken macht den Kreis zur Fokus-Region.',
        BALKEN_NOTE='Landkreise und kreisfreien Städte, absteigend sortiert. Gestrichelt: Land Hessen und eine Vergleichsregion, die kein Kreis ist, sofern ihr Wert in den Maßstab der Kreise passt.',
        DETAIL_NOTE='Alle Merkmale der gewählten Tabelle für Fokus- und Vergleichsregion, neueste Ausgabe. Hessen-Vergleich: bei Anteilen, Quoten und Durchschnitten der Index (Hessen = 100), bei Anzahlen und Summen der Anteil am Landeswert; bei Veränderungsraten, Salden und Anteilen am Land kein Vergleich (—). Klick auf eine Zeile zeigt das Merkmal oben im Kreisvergleich.',
        ERL_NOTE='Methodische Hinweise des Hessischen Statistischen Landesamts aus der Ausgabe <span id="erl-ausgabe">—</span> der Hessischen Kreiszahlen sowie Erläuterungen zu den Berechnungen in diesem Dashboard. Die Fußnoten zu einzelnen Tabellen und Merkmalen stehen jeweils direkt unter dem Kreisvergleich und der Detailtabelle.',
        BERECHNUNGEN=[
            'Grundlage sind die Excel-Ausgaben der Hessischen Kreiszahlen ab Frühjahr 2023. Ältere Ausgaben sind über Tabellentitel und Spaltenköpfe den Merkmalen der Ausgabe Herbst 2026 zugeordnet; Tabellen, die es dort nicht mehr gibt, sind nicht enthalten.',
            'Kreisvergleich, Kacheln und Detailtabelle zeigen die neueste Ausgabe. Der Zeitverlauf zeigt je Berichtszeitraum den Wert der neuesten Ausgabe, die ihn enthält; Halbjahres- und Jahreswerte werden nicht gemischt.',
            'Kacheln: Anteile und Quoten sind aus den genannten Tabellenspalten berechnet (Zähler geteilt durch Nenner). Rang 1 ist der höchste Wert unter den Landkreisen und kreisfreien Städten; gleiche Werte erhalten den gleichen Rang.',
            'Hanau ist seit 1. Januar 2026 kreisfrei. Weisen die Tabellen einer Kennzahl Hanau unterschiedlich aus, wird der Main-Kinzig-Kreis einschließlich Hanau gerechnet und so gekennzeichnet.',
            'Prozent-Modus im Kreisvergleich: Anteil des Merkmals an einer geprüften Summe derselben Tabelle (z. B. Abiturientinnen und Abiturienten an allen Schulentlassenen) oder am Landeswert.',
            'Hessen-Vergleich in der Detailtabelle: bei Anteilen, Quoten und Durchschnitten der Index (Hessen = 100), bei Anzahlen und Summen der Anteil am Landeswert.',
        ],
        QUELLE='Quelle: Hessisches Statistisches Landesamt, Hessische Kreiszahlen (Excel-Ausgaben {{SPANNE}}). Kennzahlen und Darstellung: eigene Berechnungen. Geheim gehaltene oder fehlende Werte werden nicht angezeigt.',
        cfg=dict(daten='daten.json', fokus='532000', vergleich='000000', mid='17_15', einheiten=['Landkreis', 'Kreisfreie Stadt'],
                 gruppen=[['Landkreise', ['Landkreis']], ['Kreisfreie Städte', ['Kreisfreie Stadt', 'Sonderstatusstadt']], ['Regierungsbezirke', ['Regierungsbezirk']], ['Land und Zusammenfassungen', ['Land', 'Summe', 'Verband']]],
                 ref='000000', refName='Hessen', refVon='Hessen', refOption='Landeswert Hessen (Anteil des Kreises)', linien=[['000000', 'Hessen']],
                 keinRang='Kein Rang (Region ist kein Kreis)', einheitDat='Kreisen', keineWerte='Für dieses Merkmal liegen keine Werte auf Kreisebene vor.',
                 rateTabs=[13, 23, 25, 33], noIdxTabs=[21], ohneBasisTabs=[10], nurRef=[],
                 hinweise=[[[7, 11, 12, 13], 'Bevölkerungszahlen bis zum Stand Ende 2023 beruhen auf dem Zensus 2011, ab Stand 2024 auf dem Zensus 2022. Werte davor und danach sind nicht direkt vergleichbar.'],
                           [[52], 'Laut Statistischem Landesamt sind die Ergebnisse ab 2025 nicht mit den Vorjahren vergleichbar (bereinigter Abgleich mit den Melderegistern).']])),
    'gemeinde': dict(
        datei='gemeinden.html', daten='gemeinden.json', kacheln='kacheln_gemeinden.json', basis='bezugsgroessen_gemeinden.json', spanne=lambda D: f'{D["ausgaben"][0][-4:]} bis {D["ausgaben"][-1][-4:]}',
        TITEL='Gemeindezahlen Lahn-Dill-Kreis',
        BESCHREIBUNG='Strukturdaten der 23 Städte und Gemeinden des Lahn-Dill-Kreises im Vergleich, nach der Hessischen Gemeindestatistik des Statistischen Landesamts.',
        LEDE='Hessische Gemeindestatistik des Statistischen Landesamts für die 23 Städte und Gemeinden des Lahn-Dill-Kreises,',
        TAKE='<span class="rn" data-slot="fokus">—</span>: <b class="kv" data-k="bev" data-slot="fokus">—</b> Einwohnerinnen und Einwohner, Veränderung zum Vorjahr <b class="kv" data-k="bevVer" data-slot="fokus">—</b> (<span class="rn" data-slot="vgl">—</span>: <span class="kv" data-k="bev" data-slot="vgl">—</span> bzw. <span class="kv" data-k="bevVer" data-slot="vgl">—</span>).',
        KACHEL_NOTE='Kennzahlen zum oben gewählten Thema, Werte der neuesten Ausgabe. Rang 1 = höchster Wert unter den 23 Städten und Gemeinden des Lahn-Dill-Kreises.',
        VERGLEICH_H2='Jedes Merkmal im Gemeindevergleich',
        VERGLEICH_NOTE='Alle rund 250 Merkmale aus 21 Tabellen. Wählen Sie Thema, Tabelle und Merkmal. Ein Klick auf einen Balken macht die Gemeinde zur Fokus-Region.',
        BALKEN_NOTE='Die 23 Städte und Gemeinden, absteigend sortiert. Gestrichelt: Lahn-Dill-Kreis und Land Hessen, sofern ihr Wert in den Maßstab der Gemeinden passt (bei Anteilen, Quoten und Durchschnitten).',
        DETAIL_NOTE='Alle Merkmale der gewählten Tabelle für Fokus- und Vergleichsregion, neueste Ausgabe. Kreis-Vergleich: bei Anteilen, Quoten und Durchschnitten der Index (Lahn-Dill-Kreis = 100), bei Anzahlen und Summen der Anteil am Kreiswert; bei Veränderungsraten und Salden kein Vergleich (—). Klick auf eine Zeile zeigt das Merkmal oben im Gemeindevergleich.',
        ERL_NOTE='Methodische Hinweise und Begriffe des Hessischen Statistischen Landesamts aus der <span id="erl-ausgabe">—</span> der Hessischen Gemeindestatistik sowie Erläuterungen zu den Berechnungen in diesem Dashboard. Fußnoten und Begriffe zu einzelnen Tabellen stehen jeweils direkt unter dem Gemeindevergleich und der Detailtabelle.',
        BERECHNUNGEN=[
            'Grundlage sind die Excel-Ausgaben 2021 bis 2025 der Hessischen Gemeindestatistik, daraus nur der Lahn-Dill-Kreis, seine 23 Städte und Gemeinden und das Land Hessen. Ältere Ausgaben sind über Tabellentitel und Spaltenköpfe den Merkmalen der Ausgabe 2025 zugeordnet; Tabellen, die es dort nicht mehr gibt (Beschäftigte und Pendler bis Ausgabe 2023), sind nicht enthalten.',
            'Die Ausgabe 2021 stellte mehrere Tabellen auf einem Blatt zusammen, die Ausgaben 2021 und 2022 die Gemeindefinanzen in einer einzigen Tabelle. Aus diesen sind nur Spalten mit gleichem Inhalt wie heute übernommen, jede einzeln geprüft.',
            'Gemeindevergleich, Kacheln und Detailtabelle zeigen die neueste Ausgabe. Der Zeitverlauf zeigt je Berichtszeitraum den Wert der neuesten Ausgabe, die ihn enthält.',
            'Kacheln: Anteile, Quoten und Werte je Einwohner/-in sind aus den genannten Tabellenspalten derselben Ausgabe berechnet (Zähler geteilt durch Nenner). Rang 1 ist der höchste Wert unter den 23 Städten und Gemeinden; gleiche Werte erhalten den gleichen Rang.',
            'In den Tabellen zu Gemeindefinanzen und Gemeindepersonal enthält der Wert des Lahn-Dill-Kreises auch die Kreisverwaltung. Er ist dort größer als die Summe der Gemeinden.',
            'Prozent-Modus im Gemeindevergleich: Anteil des Merkmals an einer geprüften Summe derselben Tabelle (z. B. Waldfläche an der Bodenfläche) oder am Wert des Lahn-Dill-Kreises.',
            'Kreis-Vergleich in der Detailtabelle: bei Anteilen, Quoten und Durchschnitten der Index (Lahn-Dill-Kreis = 100), bei Anzahlen und Summen der Anteil am Kreiswert.',
        ],
        QUELLE='Quelle: Hessisches Statistisches Landesamt, Hessische Gemeindestatistik (Excel-Ausgaben {{SPANNE}}). Kennzahlen und Darstellung: eigene Berechnungen. Geheim gehaltene oder fehlende Werte werden nicht angezeigt.',
        cfg=dict(daten='gemeinden.json', fokus='532023', vergleich='532000', mid='1_4', einheiten=['Stadt', 'Gemeinde'],
                 gruppen=[['Städte', ['Stadt']], ['Gemeinden', ['Gemeinde']], ['Kreis und Land', ['Landkreis', 'Land']]],
                 ref='532000', refName='Kreis', refVon='Kreis', refOption='Kreiswert Lahn-Dill-Kreis (Anteil der Gemeinde)', linien=[['532000', 'Lahn-Dill-Kreis'], ['000000', 'Hessen']],
                 keinRang='Kein Rang (Region ist keine Gemeinde des Kreises)', einheitDat='Städten und Gemeinden', keineWerte='Für dieses Merkmal liegen keine Werte für die Gemeinden vor.',
                 rateTabs=[], noIdxTabs=[], ohneBasisTabs=[], nurRef=['15_7', '15_8', '15_9'],  # Verunglückte sind Personen, keine Unfälle
                 hinweise=[[[1, 2, 3], 'Bevölkerungszahlen bis zum Stand Ende 2023 (Ausgaben 2021 bis 2024) beruhen auf dem Zensus 2011, ab Stand Ende 2024 (Ausgabe 2025) auf dem Zensus 2022. Werte davor und danach sind nicht direkt vergleichbar.'],
                           [[13], 'Der Wohnungsbestand ist bis Ende 2023 aus dem Zensus 2011, ab Ende 2024 aus dem Zensus 2022 fortgeschrieben. Werte davor und danach sind nicht direkt vergleichbar.'],
                           [[5, 6, 7], 'Die Ausgabe 2021 enthält die Agrarstrukturerhebung 2016, die späteren Ausgaben die Landwirtschaftszählung 2020; laut Statistischem Landesamt sind beide voll vergleichbar.'],
                           [[16, 17, 18, 19], 'Der Wert des Lahn-Dill-Kreises enthält in dieser Tabelle auch die Kreisverwaltung.'],
                           [[16, 17, 18], 'Bis zur Ausgabe 2022 standen die Gemeindefinanzen in einer gemeinsamen Tabelle; für diese Jahre sind nur Merkmale mit gleichem Inhalt vorhanden.']])),
}

def build_page(mode):
    S = SEITEN[mode]
    t = open(os.path.join(TOOLS, 'vorlage.html'), encoding='utf-8').read()
    kacheln = open(os.path.join(TOOLS, S['kacheln']), encoding='utf-8').read().strip()
    d3 = open(os.path.join(TOOLS, 'd3.min.js'), encoding='utf-8').read()
    D = json.load(open(os.path.join(ROOT, S['daten']), encoding='utf-8'))
    spanne = S['spanne'](D)
    basis = open(os.path.join(TOOLS, S['basis']), encoding='utf-8').read().strip()
    schalter = ''.join(f'<a href="{href}"' + (' aria-current="page"' if m == mode else '') + f'>{lab}</a>' for m, _, href, lab in SCHALTER)
    for k in ('/*D3*/', '/*KACHELN*/', '/*BASIS*/', '/*CFG*/'): assert t.count(k) == 1, k
    out = t.replace('/*BASIS*/', basis, 1).replace('/*KACHELN*/', kacheln, 1).replace('/*CFG*/', json.dumps(S['cfg'], ensure_ascii=False), 1)
    out = out.replace('{{SCHALTER}}', schalter).replace('{{BERECHNUNGEN}}', '\n'.join(f'          <li>{x}</li>' for x in S['BERECHNUNGEN']))
    for k in ('TITEL', 'BESCHREIBUNG', 'LEDE', 'TAKE', 'KACHEL_NOTE', 'VERGLEICH_H2', 'VERGLEICH_NOTE', 'BALKEN_NOTE', 'DETAIL_NOTE', 'ERL_NOTE', 'QUELLE'):
        out = out.replace('{{' + k + '}}', S[k])
    out = out.replace('{{SPANNE}}', spanne)
    assert '{{' not in out, re.findall(r'\{\{\w+\}\}', out)
    out = out.replace('/*D3*/', d3, 1)
    open(os.path.join(ROOT, S['datei']), 'w', encoding='utf-8').write(out)
    print(S['datei'] + ' gebaut (' + spanne + ')')

def cmd_seite():
    build_page('kreis')
    if os.path.exists(os.path.join(ROOT, 'gemeinden.json')): build_page('gemeinde')

if __name__ == '__main__':
    {'daten': cmd_daten, 'seite': cmd_seite}.get(sys.argv[1] if len(sys.argv) > 1 else '', lambda: sys.exit(__doc__))()
