#!/usr/bin/env python3
"""Aufbereitung der Hessischen Gemeindestatistik (nur Lahn-Dill-Kreis und seine 23 Städte und Gemeinden)
für das Dashboard des Bildungsbüros Lahn-Dill-Kreis. Gleiche Systematik wie werkzeuge/kreiszahlen.py.

Aufruf (aus dem Hauptordner des Repositorys):
    python3 werkzeuge/gemeindezahlen.py katalog   # schreibt werkzeuge/katalog/gemeinden_merkmale.csv aus der Referenzausgabe neu
    python3 werkzeuge/gemeindezahlen.py daten     # liest quellen/HGSt_j*.xlsx, schreibt gemeinden.json und werkzeuge/ausgabe/gemeinden_*
    python3 werkzeuge/gemeindezahlen.py seite     # baut gemeinden.html aus werkzeuge/vorlage.html

Grundsätze:
- Referenz für Tabellen- und Spaltennummern (mid = "<Tabelle>_<Spalte>") ist die Ausgabe 2025 (REFERENZ).
  Ältere Ausgaben werden über Tabellentitel und Spaltenköpfe auf diese Merkmale abgebildet. Hatten ältere Ausgaben
  andere Tabellenzuschnitte (2021: ein Blatt je Themenbereich; 2022: Finanzen in einer Tabelle), werden ihre Spalten
  gegen alle Tabellen der Referenz aus GRUPPEN abgeglichen. Tabellen ohne Gegenstück (z. B. Beschäftigte und
  Pendler, nur bis Ausgabe 2023) fallen weg und stehen im Bericht.
- Regionen: Land Hessen und Lahn-Dill-Kreis als Bezugswerte, dazu die 23 Städte und Gemeinden (AGS 532001–532023).
  Die Kreisverwaltung (532500), die in den Finanztabellen eigens steht, ist nicht enthalten.
- Fehlende oder geheime Werte bleiben leer, "—" wird 0, Dezimalwerte auf 4 Stellen (wie bei den Kreiszahlen).
- Dateinamen in quellen/: HGSt_j<JJJJ>*.xlsx = Ausgabe JJJJ.
"""
import collections, csv, glob, json, os, re, sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kreiszahlen as K  # noqa: E402

ROOT, TOOLS, QUELLEN = K.ROOT, K.TOOLS, K.QUELLEN
REFERENZ = 'HGSt_j2025.xlsx'
MAX_TAB = 21  # Tabelle 22 ist das Verzeichnis der Verbände
KATALOG = os.path.join(TOOLS, 'katalog', 'gemeinden_merkmale.csv')
REGIONEN = os.path.join(TOOLS, 'katalog', 'gemeinden_regionen.csv')

# Ältere (Teil-)Tabellen, deren Titel keinem Titel der Referenz entspricht: Muster für den Titel → Tabellen der Referenz.
# Die Spalten werden dann gegen alle genannten Tabellen abgeglichen.
GRUPPEN = [
    (r'^Fläche am ', [1]),
    (r'^Natürliche Bevölkerungsbewegung|^Wanderungsbewegung', [4]),
    (r'^Betriebe insgesamt$', [7]),  # 2021: Kopfspalte der Tabelle „Ausgewählte Merkmale“
    (r'^Wohngebäude am ', [13]),
]
# Finanzen der Gemeinden bis Ausgabe 2022 (eine Tabelle statt der heutigen Tabellen 16–18): nur Spalten mit
# gleichem Inhalt, einzeln geprüft. Muster für den Spaltenkopf (ohne Titelzeile) → Merkmal der Referenz.
FINANZEN_ALT = [
    (r'^Einzahlungen › (?:Einzahlungen )?insgesamt › Laufende Verwaltungstätigkeit › 1 000 Euro$', '16_6'),
    (r'^Einzahlungen › (?:Einzahlungen )?insgesamt › Investitions- und Finanzierungs-? ?tätigkeit › 1 000 Euro$', '16_7'),
    (r'^Einzahlungen › Schlüsselzuweisungen vom Land › 1 000 Euro$', '16_8'),
    (r'^Einzahlungen › Steuern und steuerähnliche Abgaben insgesamt › 1 000 Euro$', '16_9'),
    (r'^Auszahlungen › (?:Auszahlungen )?insgesamt › Laufende Verwaltungstätigkeit › 1 000 Euro$', '16_12'),
    (r'^Auszahlungen › (?:Auszahlungen )?insgesamt › Investitions-? ?und Finanzierungs-? ?tätigkeit › 1 000 Euro$', '16_13'),
    (r'^Auszahlungen › Baumaßnahmen › 1 000 Euro$', '16_14'),
    (r'^Auszahlungen › Personal- und Versorgungsauszahlungen › insgesamt › 1 000 Euro$', '16_15'),
    (r'^Auszahlungen › Personal- und Versorgungsauszahlungen › (?:der laufenden Verwaltungstätigkeit in Prozent|in Prozent der laufenden Verwaltungstätigkeit)$', '16_16'),
    (r'^Realsteuerhebesätze › Grundsteuer A › (?:in )?%$', '17_8'),
    (r'^Realsteuerhebesätze › Grundsteuer B › (?:in )?%$', '17_9'),
    (r'^Realsteuerhebesätze › Gewerbesteuer › (?:in )?%$', '17_10'),
    (r'^Kassenkredite am 31\.12\.(?:\d{4})? › 1 000 Euro$', '18_8'),
    (r'^Finanzvermögen am 31\.12\.(?:\d{4})? › 1 000 Euro$', '18_10'),
]
# Geprüfte Einzelfälle, die der automatische Abgleich nicht erkennt (meist 2021: Kopf ohne Tabellentitel).
# (Muster Tabellentitel, Muster Spaltenkopf, Merkmal der Referenz); gilt nur für sonst nicht zugeordnete Spalten.
OVERRIDES = [
    (r'^Fläche am |^Bevölkerung', r'^Fläche am 31\.12\.\d{4} in km2$', '1_3'),
    (r'^Landwirtschaftliche Betriebe nach Größenklassen', r'^LF der Betriebe in Hektar$', '5_11'),
    (r'^Bodennutzung', r'^LF insgesamt$', '6_3'),
    (r'^Wohnungen in Wohn- und Nichtwohngebäuden', r'^insgesamt$', '13_3'),
    (r'^Wohnungen in Wohn- und Nichtwohngebäuden', r'^Räume$', '13_10'),
    (r'^Wohngebäude am ', r'^insgesamt$', '13_12'),
    (r'^Bauhauptgewerbe', r'^Geleistete Arbeitsstunden der Arbeiterinnen und Arbeiter in 1 000$', '10_5'),
    (r'^Baugenehmigungen', r'^(Geplante )?Wohnungen insgesamt$', '11_16'),
    (r'^Baufertigstellungen', r'^Fertig gestellte Wohnungen insgesamt$', '12_16'),
]
STICHTAG_31_12 = {f'8_{i}' for i in range(4, 20)} | {'18_4', '18_5', '18_8', '18_9', '18_10'}
JAHRESWERT = {'16_17'}  # Finanzierungssaldo: Ein- minus Auszahlungen des Jahres (Fußnote der Quelle)
FINANZ_TITEL = r'^Finanzen der Gemeinden und Gemeindeverbände'

def ntitle(t):
    t = re.sub(r'\(Fortschreibungsergebnisse[^)]*\)?', ' ', t)
    t = re.sub(r'\(Stand.*$', ' ', t)
    t = re.sub(r'\b\d{1,2}\.\d{1,2}\.((?:19|20)\d{2})\b', r'\1', t)
    t = re.sub(r'\bim Juni\b', 'im', t)
    s = K.norm_title(t)
    return re.sub(r'\s+', ' ', s.replace(' in hektar', '')).strip()

def ausgabe_label(fn):
    m = re.match(r'HGSt_j(\d{4})', os.path.basename(fn))
    if not m: raise SystemExit(f'Dateiname nicht erkannt: {fn} (erwartet HGSt_jJJJJ….xlsx)')
    return int(m.group(1)), 'Ausgabe ' + m.group(1)

def data_sheets(wb):
    names = wb.sheetnames
    if any(n.isdigit() for n in names): return [(int(n), n) for n in names if n.isdigit()]
    i0 = next(i for i, n in enumerate(names) if n.startswith('Inhalt'))
    return [(k, n) for k, n in enumerate(names[i0 + 1:], 1)]

def parse(path):
    """Wie kreiszahlen.parse, zusätzlich für Ausgaben mit benannten Blättern, ohne Titelzeile (2021)
    und mit einer Zeile Spaltennummern unter dem Tabellenkopf."""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for no, name in data_sheets(wb):
        ws = wb[name]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows: continue
        c0 = K.clean(rows[0][0])
        notitle = 'Gebiets' in c0 or 'schlüssel' in c0
        h = 0 if notitle else 1  # erste Kopfzeile
        title = name if notitle else re.sub(r'^\d+\.\s*', '', c0)
        title_fn = [] if notitle else K.markers(rows[0][0])
        hdr0 = [K.clean(x) for x in rows[h]]
        if not ('Gebiets' in (hdr0[0] or '') or 'schlüssel' in (hdr0[0] or '')):
            out[no] = dict(title=title, sheet=name, cols={}, rows=[]); continue
        first = next((i for i in range(h + 1, min(len(rows), 15)) if rows[i][0] is not None and re.fullmatch(r'\d{6}', str(rows[i][0]).strip())), None)
        if first is None:
            out[no] = dict(title=title, sheet=name, cols={}, rows=[]); continue
        numrow = {i for i in range(h, first) if all(isinstance(x, (int, float)) or x is None for x in rows[i][2:]) and any(isinstance(x, (int, float)) for x in rows[i][2:])}
        grid = {(i, j): v for i in range(h, first) for j, v in enumerate(rows[i])}
        for mr in ws.merged_cells.ranges:
            if mr.min_row - 1 >= first or mr.max_row - 1 < h: continue
            v = ws.cell(mr.min_row, mr.min_col).value
            for j in range(mr.min_col - 1, mr.max_col): grid[(mr.min_row - 1, j)] = v
        ncol = max(len(r) for r in rows[h:first + 1])
        cols, colfn = {}, {}
        for j in range(2, ncol):
            path, mk = [], set()
            for i in range(h, first):
                if i in numrow: continue
                t = K.clean(grid.get((i, j)))
                mk.update(K.markers(grid.get((i, j))))
                if t and (not path or path[-1] != t): path.append(t)
            if path and path != ['Zurück zum Inhalt']:
                cols[j + 1] = K.merge_fragments(path)
                if mk: colfn[j + 1] = sorted(mk)
        data, last = [], first
        for k, r in enumerate(rows[first:], first):
            if r and r[0] is not None and re.fullmatch(r'\d{6}', str(r[0]).strip()): last = k
        for r in rows[first:]:
            if r[0] is None: continue
            k = str(r[0]).strip()
            if not re.fullmatch(r'\d{6}', k): continue
            data.append(dict(ags=k, name=K.clean(r[1]), vals={j + 1: r[j] for j in range(2, len(r)) if j + 1 in cols}))
        fn, quelle = K.notes_text(rows, last + 1)
        if not notitle:
            out[no] = dict(title=title, sheet=name, cols=cols, rows=data, title_fn=title_fn, colfn=colfn, fn=fn, quelle=quelle); continue
        # Ausgabe 2021: mehrere Tabellen nebeneinander auf einem Blatt, erste Kopfzeile = Tabellentitel
        heads = []
        for c in sorted(cols):
            if cols[c][0] not in heads: heads.append(cols[c][0])
        for k, hd in enumerate(heads, 1):
            if hd.startswith('Zurück zu'): continue
            sub = {c: (p[1:] or p) for c, p in cols.items() if p[0] == hd}
            out[no * 100 + k] = dict(title=hd, sheet=name, cols=sub, rows=[dict(r, vals={c: v for c, v in r['vals'].items() if c in sub}) for r in data],
                                     title_fn=[], colfn={c: v for c, v in colfn.items() if c in sub}, fn=fn, quelle=quelle)
    # Zeichenerklärung und Hinweise (Blatt „Impressum“), Themengliederung (Blatt „Inhalt“)
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
    themen = {}
    if 'Inhalt' in wb.sheetnames:
        cur = None
        for r in wb['Inhalt'].iter_rows(values_only=True):
            cells = [c for c in r if c is not None and str(c).strip()]
            if len(cells) == 1 and isinstance(cells[0], str): cur = cells[0].strip()
            elif len(cells) >= 2 and str(cells[0]).strip().isdigit() and cur: themen[int(str(cells[0]).strip())] = cur
    methodik = []  # Blatt „Erläuterungen“: Abschnitte mit fett gesetzter Überschrift
    if 'Erläuterungen' in wb.sheetnames:
        prev_bold = False
        for r in wb['Erläuterungen'].iter_rows():
            cells = [c for c in r if c.value is not None and str(c.value).strip()]
            if not cells: continue
            txt = re.sub(r'[ \t\xa0]+', ' ', ' '.join(str(c.value) for c in cells).replace('\n', ' ')).strip()
            if set(txt) <= set('_ '): continue
            bold = all(c.font is not None and c.font.b for c in cells)
            if bold:
                if prev_bold and methodik and not methodik[-1]['absaetze']: methodik[-1]['titel'].append(txt)
                else: methodik.append(dict(titel=[txt], absaetze=[]))
            elif methodik:
                methodik[-1]['absaetze'].append(txt)
            prev_bold = bold
        methodik = [m for m in methodik if m['absaetze'] and not m['titel'][0].startswith('Regionalverband')]
    out['_allgemein'] = dict(zeichen=zeichen, hinweise=hinweise, themen=themen, methodik=methodik)
    return out

def files():
    fs = sorted(glob.glob(os.path.join(QUELLEN, 'HGSt_j*.xlsx')), key=lambda f: ausgabe_label(f)[0])
    if not any(os.path.basename(f) == REFERENZ for f in fs): raise SystemExit('Referenzausgabe fehlt in quellen/: ' + REFERENZ)
    return fs

# ---------------------------------------------------------------- Befehl: katalog
LABEL_FIX = [(r'(\w)- (?!(?:bzw|und|oder)\b)(?=[a-zäöüß])', r'\1'), (r'(\w)-(?=(?:schaftlich|gewerbe|kredite)\b)', r'\1'), (r'/ -', '/-')]
def label(t):
    for a, b in LABEL_FIX: t = re.sub(a, b, t)
    return t

def cmd_katalog():
    d = parse(os.path.join(QUELLEN, REFERENZ))
    th = d['_allgemein']['themen']
    with open(KATALOG, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh, lineterminator='\n'); w.writerow(['mid', 'tab', 'thema', 'tabelle', 'merkmal', 'einheit'])
        n = 0
        for t in sorted(k for k in d if isinstance(k, int) and k <= MAX_TAB):
            for c, p in sorted(d[t]['cols'].items()):
                w.writerow([f'{t}_{c}', t, th.get(t, ''), re.sub(r'\s*\(Stand.*$', '', d[t]['title']), label(' › '.join(p)), 'Hektar' if '(in Hektar)' in d[t]['title'] else '']); n += 1
    print(f'{n} Merkmale → {KATALOG}')

# ---------------------------------------------------------------- Befehl: daten
def cmd_daten():
    cat = list(csv.DictReader(open(KATALOG, encoding='utf-8')))
    reg = sorted(csv.DictReader(open(REGIONEN, encoding='utf-8')), key=lambda r: int(r['reihenfolge']))
    mids = {r['mid'] for r in cat}
    fs = files()
    parsed = {os.path.basename(f): parse(f) for f in fs}
    allgemein = {fn: d.pop('_allgemein', {}) for fn, d in parsed.items()}
    newest = os.path.basename(fs[-1])
    cur = parsed[REFERENZ]
    inv = {ntitle(d['title']): t for t, d in cur.items() if t <= MAX_TAB}
    valid = {r['ags'] for r in reg}
    werte, bezug, bericht, labels, notes = [], [], [], [], {}
    ohne = []
    for f in fs:
        fn = os.path.basename(f); _, lab = ausgabe_label(f); labels.append(lab)
        bericht.append(f'== {lab} ({fn})')
        for t, d in sorted(parsed[fn].items()):
            if fn == REFERENZ:
                cts = [t] if t <= MAX_TAB else []
            elif re.search(FINANZ_TITEL, d['title']):
                cts = [16, 17, 18]
            else:
                ct = inv.get(ntitle(d['title']))
                cts = [ct] if ct else next((g for pat, g in GRUPPEN if re.search(pat, d['title'])), [])
            if not cts:
                if t <= MAX_TAB or fn != REFERENZ: bericht.append(f'   Blatt „{d["sheet"]}“ ohne Gegenstück in der Referenz, entfällt: {d["title"][:100]}')
                continue
            oc = d['cols']
            cc = {(ct, c): p for ct in cts for c, p in cur[ct]['cols'].items() if f'{ct}_{c}' in mids}
            if fn == REFERENZ:
                cmap, sc = {c: (t, c) for c in oc if (t, c) in cc}, {c: 1.0 for c in oc}
            elif cts == [16, 17, 18]:
                cmap, sc = {}, {}
                for c, pth in oc.items():
                    hit = [m for pat, m in FINANZEN_ALT if re.match(pat, ' › '.join(pth))]
                    if len(hit) == 1:
                        t_, c_ = map(int, hit[0].split('_')); cmap[c] = (t_, c_); sc[c] = 1.0
            else:
                cmap, sc = K.align(oc, cc)
            if fn != REFERENZ:
                for tp, pp, mid_ in OVERRIDES:
                    key = tuple(map(int, mid_.split('_')))
                    if key not in cc or key in cmap.values() or not re.search(tp, d['title']): continue
                    hit = [c for c in oc if c not in cmap and re.match(pp, ' › '.join(oc[c]))]
                    if len(hit) == 1: cmap[hit[0]] = key; sc[hit[0]] = 1.0
            tag = f'„{d["sheet"]}“→T{"+".join(map(str, cts))}'
            un_o = [' › '.join(oc[c]) for c in sorted(oc) if c not in cmap]
            un_c = [f'{a}_{b}' for a, b in sorted(cc) if (a, b) not in cmap.values()]
            if un_o: bericht.append(f'   {tag}: Spalten ohne Gegenstück: {un_o}')
            if un_c: bericht.append(f'   {tag}: Merkmale ohne Wert in dieser Ausgabe: {un_c}')
            for a, b in cmap.items():
                if sc[a] < 0.8 or fn != REFERENZ and (len(cts) > 1 or d['title'] != cur[b[0]]['title'] and sc[a] < 1):
                    bericht.append(f'   {tag}: {"unsichere " if sc[a] < 0.8 else ""}Zuordnung {" › ".join(oc[a])!r} → {b[0]}_{b[1]} {" › ".join(cc[b])!r} ({sc[a]:.2f})')
            if fn == newest:
                for ct in cts:
                    spalten = {f'{ct}_{m}': sorted(set(d.get('colfn', {}).get(c, []))) for c, (t2, m) in cmap.items() if t2 == ct and d.get('colfn', {}).get(c)}
                    notes[ct] = dict(titel=d.get('title_fn', []), fn={str(k): v for k, v in sorted(d.get('fn', {}).items())}, quelle=d.get('quelle', []), spalten=spalten)
                    missing = [n for n in set(d.get('title_fn', [])) | {x for v in spalten.values() for x in v} if n not in d.get('fn', {})]
                    if missing: bericht.append(f'   {tag}: Fußnotenzeichen ohne Text: {sorted(missing)}')
            for r in d['rows']:
                if r['ags'] not in valid: continue
                for c, v in r['vals'].items():
                    if c in cmap: werte.append((lab, f'{cmap[c][0]}_{cmap[c][1]}', r['ags'], K.conv(v)))
            for c, (ct, m) in cmap.items():
                fnt = ' '.join(d.get('fn', {}).get(n, '') for n in d.get('colfn', {}).get(c, []))
                p = bezug_for(' › '.join(oc[c]), d['title'], fnt)
                if p: bezug.append(dict(ausgabe=lab, mid=f'{ct}_{m}', bezug=p[0], jahr=p[1], art=p[2], pfad=' › '.join(oc[c])))
                else: ohne.append((lab, ct, m))

    # Ohne Zeitangabe im Kopf und Titel (2021: Teiltabelle „Betriebe insgesamt“): Zeitraum der übrigen Merkmale derselben Tabelle und Ausgabe,
    # sofern dort eindeutig
    for lab, ct, m in ohne:
        cand = {(r['bezug'], r['jahr'], r['art']) for r in bezug if r['ausgabe'] == lab and r['mid'].split('_')[0] == str(ct)}
        if len(cand) == 1:
            b_, j_, a_ = cand.pop(); bezug.append(dict(ausgabe=lab, mid=f'{ct}_{m}', bezug=b_, jahr=j_, art=a_, pfad=''))
        else: bericht.append(f'   {ct}_{m} {lab}: kein Berichtszeitraum erkennbar')

    # Geprüfte Einzelfälle: Bestände am Jahresende (Kopf nennt teils kein Datum) bzw. Jahressumme trotz „am 31.12.“ im Kopf
    for r in bezug:
        if r['mid'] in STICHTAG_31_12 and r['art'] == 'Jahr': r.update(bezug=f"31. Dezember {r['jahr']}", art='Stichtag 31. Dezember')
        if r['mid'] in JAHRESWERT and r['art'] == 'Stichtag 31. Dezember': r.update(bezug=r['jahr'], art='Jahr')

    # Art des Berichtszeitraums an die Referenz angleichen, wenn ältere Köpfe denselben Zeitraum anders benennen
    # (z. B. "Fläche am 31.12.2020" statt "Fläche im Jahr 2024", "im September" statt "Zum 30. September 2024")
    ref_art = {r['mid']: r for r in bezug if r['ausgabe'] == ausgabe_label(REFERENZ)[1]}
    for r in bezug:
        ra = ref_art.get(r['mid'])
        if not ra or ra['art'] == r['art']: continue
        if ra['art'] == 'Jahr' and r['art'] == 'Stichtag 31. Dezember':
            r.update(bezug=r['jahr'], art='Jahr')
        elif ra['art'] == 'Stichtag 31. Dezember' and r['art'] == 'Jahr' and re.search(r'am 31\.12\.|31\. Dezember', r['pfad']):
            r.update(bezug=f"31. Dezember {r['jahr']}", art=ra['art'])
        elif ra['art'].startswith('Stichtag ') and r['art'] == 'Jahr' and ra['art'].split()[-1] in r['pfad']:
            r.update(bezug=f"{ra['art'][9:]} {r['jahr']}", art=ra['art'])
        else:
            bericht.append(f"   {r['mid']} {r['ausgabe']}: Berichtszeitraum {r['bezug']} ({r['art']}) weicht von der Art der Referenz ab ({ra['art']}); erscheint nicht im Zeitverlauf")
    for r in bezug: r.pop('pfad', None)

    ordn = {lab: i for i, lab in enumerate(labels)}
    catpos = {r['mid']: i for i, r in enumerate(cat)}
    regpos = {r['ags']: i for i, r in enumerate(reg)}
    werte.sort(key=lambda r: (ordn[r[0]], catpos.get(r[1], 10 ** 6), regpos.get(r[2], 10 ** 6)))
    bezug.sort(key=lambda r: (ordn[r['ausgabe']], catpos.get(r['mid'], 10 ** 6)))
    dupl = [k for k, n in collections.Counter((a, m, g) for a, m, g, _ in werte).items() if n > 1]
    if dupl: raise SystemExit(f'Doppelte Werte, Abbruch: {dupl[:5]}')

    # Plausibilität: Sprünge des Kreiswerts zur Vorjahresausgabe
    kr = collections.defaultdict(dict)
    for lab, mid, ags, v in werte:
        if ags == '532000' and v != '': kr[mid][lab] = float(v)
    catn = {r['mid']: r['merkmal'] for r in cat}
    bericht.append('== Auffällige Sprünge des Kreiswerts (Faktor > 2,5 zur Vorjahresausgabe)')
    for mid, dd in kr.items():
        seq = [(l, dd[l]) for l in labels if l in dd]
        for (l1, a), (l2, b) in zip(seq, seq[1:]):
            if a and b and (b / a > 2.5 or b / a < 0.4) and max(abs(a), abs(b)) > 5:
                bericht.append(f'   {mid} {l1}: {a:g} → {l2}: {b:g}  ({catn.get(mid, "")[:70]})')

    out = os.path.join(TOOLS, 'ausgabe')
    os.makedirs(out, exist_ok=True)
    with open(os.path.join(out, 'gemeinden_werte.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh, lineterminator='\n'); w.writerow(['ausgabe', 'mid', 'ags', 'wert']); w.writerows(werte)
    with open(os.path.join(out, 'gemeinden_bezug.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=['ausgabe', 'mid', 'bezug', 'jahr', 'art'], lineterminator='\n'); w.writeheader(); w.writerows(bezug)
    open(os.path.join(out, 'gemeinden_bericht.txt'), 'w', encoding='utf-8').write('\n'.join(bericht) + '\n')

    ri = {r['ags']: i for i, r in enumerate(reg)}
    def num(v):
        if v == '': return None
        x = float(v); return int(x) if x.is_integer() else x
    W = {lab: {} for lab in labels}
    for lab, mid, ags, v in werte: W[lab].setdefault(mid, [None] * len(reg))[ri[ags]] = num(v)
    for lab in labels: W[lab] = {m: a for m, a in W[lab].items() if any(x is not None for x in a)}
    BZ = {lab: {} for lab in labels}
    for r in bezug: BZ[r['ausgabe']][r['mid']] = [r['bezug'], r['jahr'], r['art']]
    themen, tabs = [], {}
    for r in cat:
        if r['thema'] not in themen: themen.append(r['thema'])
        tabs.setdefault(r['tab'], [themen.index(r['thema']), r['tabelle']])
    E = allgemein.get(newest, {})
    D = dict(ausgaben=labels, regionen=[[r['ags'], r['region'], r['ebene']] for r in reg], themen=themen, merkmale=[[r['mid'], r['merkmal'], r.get('einheit', '')] for r in cat],
             werte=[W[l] for l in labels], bezug=[BZ[l] for l in labels],
             quelle=f'Hessisches Statistisches Landesamt, Hessische Gemeindestatistik, Excel-Ausgaben {labels[0][-4:]} bis {labels[-1][-4:]}', tabellen=tabs,
             hinweise={str(k): v for k, v in sorted(notes.items())},
             erlaeuterungen=dict(ausgabe=labels[-1], zeichen=E.get('zeichen', []), hinweise=E.get('hinweise', []),
                                 methodik=[dict(m, tabs=sorted({inv[ntitle(t)] for t in m['titel'] if ntitle(t) in inv}) or ([5, 6, 7] if re.search(r'Landwirtschaft|Erfassungsgrenzen', m['titel'][0]) else [])) for m in E.get('methodik', [])]))
    open(os.path.join(ROOT, 'gemeinden.json'), 'w', encoding='utf-8').write(json.dumps(D, ensure_ascii=False, separators=(',', ':')))
    rows = []
    for k, v in sorted(notes.items()):
        for n, txt in v['fn'].items():
            ms = '*' if int(n) in v['titel'] else ','.join(m for m, nrs in v['spalten'].items() if int(n) in nrs)
            rows.append(dict(id=f'{k}-{n}', tab=str(k), art='fussnote', nr=n, text=txt, mids=ms))
        for i, q in enumerate(v['quelle'], 1): rows.append(dict(id=f'{k}-q{i}', tab=str(k), art='quelle', nr='', text=q, mids='*'))
    for i, h in enumerate(E.get('hinweise', []), 1): rows.append(dict(id=f'h{i}', tab='', art='hinweis', nr=str(i), text=h, mids=''))
    for i, (z, t) in enumerate(E.get('zeichen', []), 1): rows.append(dict(id=f'z{i}', tab='', art='zeichen', nr=z, text=t, mids=''))
    for i, m in enumerate(D['erlaeuterungen']['methodik'], 1):
        rows.append(dict(id=f'm{i}', tab=','.join(map(str, m['tabs'])), art='methodik', nr=str(i), text=' / '.join(m['titel']) + '\n' + '\n'.join(m['absaetze']), mids=''))
    json.dump(dict(ausgabe=labels[-1], rows=rows), open(os.path.join(out, 'gemeinden_hinweise.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
    print(f'{len(labels)} Ausgaben, {len(werte)} Werte, {len(bezug)} Berichtszeiträume. Bericht: werkzeuge/ausgabe/gemeinden_bericht.txt')

STAND = re.compile(r'\(?\s*Stand:?\s*[^)›]*\)?')
def unstand(t): return STAND.sub(' ', t)

def bezug_for(path, title, fn_text=''):
    """Berichtszeitraum: aus dem Spaltenkopf, sonst aus dem Tabellentitel (ohne Stand der Daten).
    Einwohnerzahlen als Bezugsgröße in Finanztabellen („Einwohner/-innen (30.06.2022)“) bestimmen nicht den Zeitraum der Tabelle."""
    p = K.period(unstand(re.split(r'gegenüber', path)[0]))
    m = re.search(r'Stichtag (\d{1,2}\.\d{1,2}\.(?:19|20)\d{2})', fn_text)  # z. B. Fußnote „Einwohnerzahlen … zum Stichtag 30.06.2024“
    if not p and m: p = K.period(m.group(1))
    if not p: p = K.period(K.unfoot(unstand(title)))
    m = re.search(r'\bim Juni ((?:19|20)\d{2})\b', title)  # Bauhauptgewerbe: Monatswerte für Juni
    if m and not K.period(unstand(path)) and not re.search(r'des Vorjahres', path): p = ('Juni ' + m.group(1), m.group(1), 'Monat Juni')
    if p and re.search(r'des Vorjahres', path) and p[2] == 'Jahr':
        y = str(int(p[1]) - 1); p = (y, y, 'Jahr')
    return p

# ---------------------------------------------------------------- Befehl: seite
def cmd_seite():
    K.build_page('gemeinden')

if __name__ == '__main__':
    {'katalog': cmd_katalog, 'daten': cmd_daten, 'seite': cmd_seite}.get(sys.argv[1] if len(sys.argv) > 1 else '', lambda: sys.exit(__doc__))()
