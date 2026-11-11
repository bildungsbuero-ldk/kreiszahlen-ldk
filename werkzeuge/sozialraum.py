#!/usr/bin/env python3
"""Sozialraum-Dashboard Lahn-Dill-Kreis: Daten prüfen und passwortgeschützte Seite bauen.

Daten liegen im Klartext NUR lokal in intern/ (per .gitignore ausgeschlossen). Im Repository liegt davon nur das
verschlüsselte Archiv werkzeuge/sozialraum/daten.verschluesselt. Die Seite sozialraum/index.html ist ebenfalls
verschlüsselt. Passwort aus der Umgebungsvariable SOZIALRAUM_PASSWORT oder der Datei werkzeuge/sozialraum/.passwort.

    python3 werkzeuge/sozialraum.py entpacken  # Archiv mit dem Passwort nach intern/ entpacken (vor dem Bearbeiten)
    python3 werkzeuge/sozialraum.py pruefen    # alle CSV-Dateien in intern/daten/ prüfen, Bericht nach intern/bericht.txt
    python3 werkzeuge/sozialraum.py seite      # prüfen, berechnen, Seite bauen und verschlüsseln, Archiv aktualisieren
    python3 werkzeuge/sozialraum.py hsl        # Gemeindestatistik (Excel in quellen/sozialraum/) -> intern/daten/hsl_gemeindestatistik.csv

Benötigt das Paket cryptography; für `hsl` zusätzlich openpyxl.
"""
import csv
import glob
import json
import math
import os
import re
import statistics
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(ROOT, 'werkzeuge')
SRC = os.path.join(TOOLS, 'sozialraum')
INTERN = os.path.join(ROOT, 'intern')
DATA = os.path.join(INTERN, 'daten')
ARCHIV = os.path.join(ROOT, 'werkzeuge', 'sozialraum', 'daten.verschluesselt')
OUT_PAGE = os.path.join(ROOT, 'sozialraum', 'index.html')
REPORT = os.path.join(INTERN, 'bericht.txt')

GEMEINDEN = ['Aßlar', 'Bischoffen', 'Braunfels', 'Breitscheid', 'Dietzhölztal', 'Dillenburg', 'Driedorf',
             'Ehringshausen', 'Eschenburg', 'Greifenstein', 'Haiger', 'Herborn', 'Hohenahr', 'Hüttenberg',
             'Lahnau', 'Leun', 'Mittenaar', 'Schöffengrund', 'Siegbach', 'Sinn', 'Solms', 'Waldsolms', 'Wetzlar']
VERGLEICH = ['Lahn-Dill-Kreis', 'Hessen']
KATEGORIEN = ['Psychiatrie (Klinik/Ambulanz)', 'Sozialpsychiatrie', 'Erziehungs-/Familienberatung', 'Suchtberatung',
              'Krisen/Gewaltschutz', 'Selbsthilfe', 'Schulpsychologie', 'Frühe Hilfen', 'Familienzentrum',
              'Schulsozialarbeit', 'Jugendzentrum/Jugendpflege', 'Quartier/Mehrgenerationen', 'Migration/Integration',
              'Existenzsicherung (Schuldner, Tafel)', 'Sonstiges']
SONDERORTE = ['Lahn-Dill-Kreis (kein Ort angegeben)', 'Gießen (außerhalb LDK)', 'Weilburg (außerhalb LDK)']
VOR_ORT = {'Familienzentrum', 'Schulsozialarbeit', 'Jugendzentrum/Jugendpflege', 'Quartier/Mehrgenerationen'}
SCHULARTEN = ['Grundschule', 'Weiterführende Schule', 'Förderschule', 'Berufliche Schule', 'nicht zugeordnet']
GANZTAG = ['Pakt für den Ganztag', 'Profil 1', 'Profil 2', 'Profil 3', 'Betreuung', 'unbekannt']
GANZTAG_ZAEHLT = {'Pakt für den Ganztag', 'Profil 1', 'Profil 2', 'Profil 3'}
# Bedarfsindex: Kennzahl, Richtung (+1 = höherer Wert bedeutet mehr Bedarf)
INDEX = [('kinderarmut', 1), ('sgb2', 1), ('alo_svb', 1), ('kita_mh', 1), ('kinder_u15_nichtdt', 1), ('einkuenfte_je_stpfl', -1)]
INDEX_MIN = 4

# Erwartete Spalten und Prüfregeln je Datei: (Spalte, Typ, Pflicht, (min, max))
SCHEMA = {
    'hsl_gemeindestatistik.csv': {'zeilen': 'gemeinden+vergleich', 'spalten': [
        ('ags', 'text', True, None), ('gemeinde', 'gemeinde', True, None), ('einwohner', 'int', True, (1000, 7000000)),
        ('flaeche', 'num', True, (1, 30000)), ('dichte', 'num', True, (10, 5000)), ('veraend_vj', 'num', True, (-10, 10)),
        ('u6', 'num', True, (0, 20)), ('u6_15', 'num', True, (0, 25)), ('ab65', 'num', True, (5, 50)),
        ('nichtdt', 'num', True, (0, 60)), ('kinder_u15_nichtdt', 'num', True, (0, 60)),
        ('geburten_je1000', 'num', True, (0, 30)), ('wanderungssaldo_je1000', 'num', True, (-60, 60)),
        ('einkuenfte_je_stpfl', 'num', True, (15000, 150000)), ('whg_klein', 'num', True, (0, 60))]},
    'wegweiser_kommune_2023.csv': {'zeilen': 'gemeinden+vergleich', 'spalten': [
        ('gemeinde', 'gemeinde', True, None)] + [(c, 'num', False, (0, 100)) for c in
        ['sgb2', 'alg2', 'kinderarmut', 'jugendarmut', 'altersarmut', 'alo_svb', 'alo_svb_u25', 'efh_anteil',
         'auslaender', 'auslaender_u15', 'kita3_mh', 'kita_mh', 'sgb2_ausl', 'alo_bev']] + [('wohnflaeche', 'num', False, (20, 100))]},
    'kita_quoten_2023.csv': {'zeilen': 'gemeinden+vergleich', 'spalten': [
        ('gemeinde', 'gemeinde', True, None), ('kita_quote_u3', 'num', False, (0, 100)), ('kita_quote_3_5', 'num', False, (0, 130))]},
    'kita_mh_kreis.csv': {'zeilen': 'gemeinden', 'spalten': [
        ('gemeinde', 'gemeinde', True, None), ('kitas', 'int', False, (1, 60)), ('kinder', 'int', False, (10, 5000)),
        ('migrantenkinder', 'int', False, (0, 5000)), ('kita_mh_kreis', 'num', False, (0, 100))]},
    'aerzte_kommune.csv': {'zeilen': 'gemeinden', 'spalten': [
        ('gemeinde', 'gemeinde', True, None), ('hausaerzte', 'int', True, (0, 200)), ('kinderaerzte', 'int', True, (0, 50)),
        ('mittelbereich_hausaerzte', 'enum:Haiger/Dillenburg|Herborn|Wetzlar', True, None)]},
    'versorgungsgrade_kv.csv': {'zeilen': 'frei', 'pflicht': ('planungsbereich', ['Mittelbereich Haiger/Dillenburg', 'Mittelbereich Herborn', 'Mittelbereich Wetzlar']), 'pflicht2': ('arztgruppe', ['Hausärzte', 'Psychotherapeuten (gesamt)', 'Kinder- und Jugendpsychiater']), 'spalten': [
        ('planungsbereich', 'text', True, None), ('arztgruppe', 'text', True, None),
        ('versorgungsgrad_prozent', 'num', True, (0, 400)), ('offene_sitze', 'num', False, (0, 100)),
        ('anzahl_aerzte', 'num', False, (0, 2000)), ('stand', 'datum', True, None)]},
    'hze_jugendamt_2023.csv': {'zeilen': 'frei', 'pflicht': ('bezirk', ['Kreisjugendamt Lahn-Dill (ohne Wetzlar)', 'Jugendamt Stadt Wetzlar', 'Hessen']), 'spalten': [
        ('bezirk', 'text', True, None)] + [(c, 'int', True, (0, 10000000)) for c in
        ['hilfen_31_12', 'begonnen', 'beendet', 'mit_transferleistungen', 'auslaendische_herkunft',
         'nach_gefaehrdungseinschaetzung', 'einwohner_31_12_2023']]},
    'schulen_liste.csv': {'zeilen': 'frei', 'spalten': [
        ('kommune', 'gemeinde', True, None), ('name', 'text', True, None), ('schulart', 'enum:' + '|'.join(SCHULARTEN), True, None),
        ('schulform', 'text', True, None), ('ganztag_kurz', 'enum:' + '|'.join(GANZTAG), True, None), ('quelle_url', 'url', False, None)]},
    'angebote_liste.csv': {'zeilen': 'frei', 'spalten': [
        ('kategorie', 'enum:' + '|'.join(KATEGORIEN), True, None), ('name', 'text', True, None), ('traeger', 'text', False, None),
        ('kommune', 'standort', True, None), ('weitere_kommunen', 'gemeindeliste', False, None),
        ('reichweite', 'enum:vor Ort|Fachdienst (regional/kreisweit)', True, None), ('zielgruppe', 'text', False, None),
        ('url', 'url', False, None)]},
}


def lesen(datei):
    with open(os.path.join(DATA, datei), encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def zahl(v):
    if v is None:
        return None
    v = str(v).strip()
    if v == '':
        return None
    return float(v)


def pruefen():
    fehler, hinweise, info = [], [], []
    vorhanden = sorted(os.path.basename(p) for p in glob.glob(os.path.join(DATA, '*.csv')))
    for datei in vorhanden:
        if datei not in SCHEMA and datei != 'kennzahlen_kommunen.csv':
            hinweise.append(f'{datei}: unbekannte Datei, wird nicht verwendet')
    for datei, regel in SCHEMA.items():
        if datei not in vorhanden:
            fehler.append(f'{datei}: fehlt')
            continue
        zeilen = lesen(datei)
        kopf = list(zeilen[0].keys()) if zeilen else []
        for spalte, typ, pflicht, bereich in regel['spalten']:
            if spalte not in kopf:
                fehler.append(f'{datei}: Spalte „{spalte}“ fehlt')
        for nr, z in enumerate(zeilen, start=2):
            for spalte, typ, pflicht, bereich in regel['spalten']:
                if spalte not in z:
                    continue
                roh = z[spalte] or ''
                v = roh.strip()
                ort = f'{datei}, Zeile {nr}, {spalte}'
                if roh != v:
                    fehler.append(f'{ort}: „{roh}“ hat Leerzeichen am Anfang oder Ende')
                if v == '':
                    if pflicht:
                        fehler.append(f'{ort}: leer')
                    continue
                if typ in ('num', 'int'):
                    try:
                        x = float(v)
                    except ValueError:
                        fehler.append(f'{ort}: „{v}“ ist keine Zahl (Dezimalpunkt verwenden)')
                        continue
                    if typ == 'int' and not float(x).is_integer():
                        fehler.append(f'{ort}: „{v}“ ist keine ganze Zahl')
                    if bereich and not (bereich[0] <= x <= bereich[1]):
                        fehler.append(f'{ort}: {v} liegt außerhalb des erlaubten Bereichs {bereich[0]}–{bereich[1]} (Tippfehler?)')
                elif typ == 'gemeinde':
                    erlaubt = GEMEINDEN + (VERGLEICH if regel['zeilen'] != 'gemeinden' and datei != 'schulen_liste.csv' else [])
                    if v not in erlaubt:
                        fehler.append(f'{ort}: „{v}“ ist keine Kommune des Lahn-Dill-Kreises (Schreibweise prüfen)')
                elif typ == 'standort':
                    if v not in GEMEINDEN + SONDERORTE:
                        fehler.append(f'{ort}: „{v}“ ist keine Kommune des Lahn-Dill-Kreises; erlaubt sind die 23 Kommunen und: ' + ', '.join(SONDERORTE))
                elif typ == 'gemeindeliste':
                    for teil in v.split('; '):
                        if teil not in GEMEINDEN:
                            fehler.append(f'{ort}: „{teil}“ ist keine Kommune (mehrere mit „; “ trennen)')
                elif typ.startswith('enum:'):
                    if v not in typ[5:].split('|'):
                        fehler.append(f'{ort}: „{v}“ nicht erlaubt; erlaubt: {typ[5:].replace("|", ", ")}')
                elif typ == 'url':
                    if not re.match(r'^https?://', v):
                        fehler.append(f'{ort}: „{v}“ ist kein Link (muss mit http:// oder https:// beginnen)')
                elif typ == 'datum':
                    if not re.match(r'^\d{4}-\d{2}-\d{2}$', v):
                        fehler.append(f'{ort}: „{v}“ ist kein Datum im Format JJJJ-MM-TT')
        if regel['zeilen'] in ('gemeinden', 'gemeinden+vergleich'):
            soll = GEMEINDEN + (VERGLEICH if regel['zeilen'] == 'gemeinden+vergleich' else [])
            ist = [z.get('gemeinde', '') for z in zeilen]
            for g in soll:
                if ist.count(g) == 0:
                    fehler.append(f'{datei}: Zeile für „{g}“ fehlt')
                elif ist.count(g) > 1:
                    fehler.append(f'{datei}: „{g}“ kommt {ist.count(g)}-mal vor')
        for key in ('pflicht', 'pflicht2'):
            if key in regel:
                spalte, werte = regel[key]
                ist = [z.get(spalte, '') for z in zeilen]
                for w in werte:
                    if w not in ist:
                        fehler.append(f'{datei}: Zeile mit {spalte} = „{w}“ fehlt (wird von der Seite benötigt, Schreibweise genau einhalten)')
        schluessel = 'name' if datei in ('angebote_liste.csv', 'schulen_liste.csv') else None
        if schluessel:
            namen = [(z.get('name'), z.get('kommune')) for z in zeilen]
            doppelt = {n for n in namen if namen.count(n) > 1}
            for n in sorted(doppelt):
                fehler.append(f'{datei}: „{n[0]}“ ({n[1]}) kommt mehrfach vor')
            if datei == 'angebote_liste.csv':
                nur = [z['name'] for z in zeilen]
                for n in sorted({n for n in nur if nur.count(n) > 1}):
                    fehler.append(f'{datei}: Name „{n}“ kommt mehrfach vor (Namen müssen eindeutig sein)')
        info.append(f'{datei}: {len(zeilen)} Zeilen')
    return fehler, hinweise, info


def bericht_schreiben(fehler, hinweise, info, extra=()):
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    teile = ['Sozialraum-Dashboard: Prüfbericht', '', 'Dateien:'] + ['  ' + i for i in info]
    teile += ['', f'Fehler ({len(fehler)}):'] + (['  ' + f for f in fehler] or ['  keine'])
    teile += ['', f'Hinweise ({len(hinweise)}):'] + (['  ' + h for h in hinweise] or ['  keine'])
    if extra:
        teile += [''] + list(extra)
    text = '\n'.join(teile) + '\n'
    with open(REPORT, 'w', encoding='utf-8') as f:
        f.write(text)
    print(text)


def rund(x, n=1):
    return None if x is None else round(x, n)


def vor_ort(a):
    return a['kategorie'] in VOR_ORT and a['reichweite'] == 'vor Ort'


def berechnen():
    hsl = {z['gemeinde']: z for z in lesen('hsl_gemeindestatistik.csv')}
    ww = {z['gemeinde']: z for z in lesen('wegweiser_kommune_2023.csv')}
    kita = {z['gemeinde']: z for z in lesen('kita_quoten_2023.csv')}
    kita_db = {z['gemeinde']: z for z in lesen('kita_mh_kreis.csv')}
    aerzte = {z['gemeinde']: z for z in lesen('aerzte_kommune.csv')}
    vg = lesen('versorgungsgrade_kv.csv')
    schulen = lesen('schulen_liste.csv')
    angebote = lesen('angebote_liste.csv')
    hze = lesen('hze_jugendamt_2023.csv')
    ha_vg = {z['planungsbereich'].replace('Mittelbereich ', ''): z for z in vg if z['arztgruppe'] == 'Hausärzte'}

    def zeile(g):
        h, w, k = hsl[g], ww.get(g, {}), kita.get(g, {})
        r = {'gemeinde': g, 'einwohner': int(zahl(h['einwohner']))}
        for c in ['nichtdt', 'kinder_u15_nichtdt', 'einkuenfte_je_stpfl', 'ab65', 'u6', 'geburten_je1000',
                  'wanderungssaldo_je1000', 'veraend_vj', 'whg_klein']:
            r[c] = zahl(h[c])
        for c in ['sgb2', 'kinderarmut', 'jugendarmut', 'altersarmut', 'alo_svb', 'alo_svb_u25', 'kita_mh', 'kita3_mh',
                  'sgb2_ausl', 'efh_anteil']:
            r[c] = zahl(w.get(c))
        r['kita_u3'] = zahl(k.get('kita_quote_u3'))
        r['kita_3_5'] = zahl(k.get('kita_quote_3_5'))
        r['kita_mh_kreis'] = zahl(kita_db.get(g, {}).get('kita_mh_kreis'))
        return r

    gem = [zeile(g) for g in GEMEINDEN]
    vergleich = [zeile(g) for g in VERGLEICH]

    stats = {}
    for c, s in INDEX:
        werte = [r[c] for r in gem if r[c] is not None]
        stats[c] = (statistics.mean(werte), statistics.stdev(werte))
    for r in gem:
        z = [s * (r[c] - stats[c][0]) / stats[c][1] for c, s in INDEX if r[c] is not None]
        r['n_ind'] = len(z)
        r['_index'] = sum(z) / len(z) if len(z) >= INDEX_MIN else None
        r['bedarfsindex'] = round(r['_index'], 2) if r['_index'] is not None else None
    rang = sorted([r for r in gem if r['bedarfsindex'] is not None], key=lambda r: -r['bedarfsindex'])
    for i, r in enumerate(rang, start=1):
        r['rang'] = i
    for r in gem:
        r.setdefault('rang', None)

    def in_g(a, g):
        return a['kommune'] == g or g in [t for t in (a.get('weitere_kommunen') or '').split('; ') if t]

    for r in gem:
        g = r['gemeinde']
        sch = [s for s in schulen if s['kommune'] == g]
        mine = [a for a in angebote if in_g(a, g)]
        a = aerzte.get(g, {})
        r['schulen'] = len(sch)
        r['grundschulen'] = sum(1 for s in sch if s['schulart'] == 'Grundschule')
        r['weiterfuehrend'] = sum(1 for s in sch if s['schulart'] == 'Weiterführende Schule')
        r['ganztag'] = sum(1 for s in sch if s['ganztag_kurz'] in GANZTAG_ZAEHLT)
        r['hausaerzte'] = int(zahl(a['hausaerzte'])) if a else None
        r['kinderaerzte'] = int(zahl(a['kinderaerzte'])) if a else None
        r['mittelbereich'] = a.get('mittelbereich_hausaerzte', '')
        mb = ha_vg.get(r['mittelbereich'])
        r['hausarzt_vg'] = zahl(mb['versorgungsgrad_prozent']) if mb else None
        r['familienzentren'] = sum(1 for x in mine if x['kategorie'] == 'Familienzentrum' and vor_ort(x))
        r['schulsozialarbeit'] = sum(1 for x in mine if x['kategorie'] == 'Schulsozialarbeit' and vor_ort(x))
        r['jugend'] = sum(1 for x in mine if x['kategorie'] == 'Jugendzentrum/Jugendpflege' and vor_ort(x))
        r['vorort'] = sum(1 for x in mine if vor_ort(x))
        r['_dichte'] = r['vorort'] / r['einwohner'] * 10000
        r['vorort_je10k'] = round(r['_dichte'], 1)
        r['fachdienste'] = sum(1 for x in angebote if x['kommune'] == g and not vor_ort(x))
        r['psychiatrie'] = sum(1 for x in angebote if x['kommune'] == g and x['kategorie'] in ('Psychiatrie (Klinik/Ambulanz)', 'Sozialpsychiatrie'))

    hze_rows = []
    for z in hze:
        h = int(zahl(z['hilfen_31_12']))
        ew = int(zahl(z['einwohner_31_12_2023']))
        hze_rows.append({'bezirk': z['bezirk'], 'hilfen': h, 'je_1000_ew': round(h / ew * 1000, 2),
                         'anteil_transfer': round(int(zahl(z['mit_transferleistungen'])) / h * 100, 1)})

    kreis = {
        'schulen': len(schulen), 'ganztag': sum(1 for s in schulen if s['ganztag_kurz'] in GANZTAG_ZAEHLT),
        'familienzentren': sum(1 for a in angebote if a['kategorie'] == 'Familienzentrum' and vor_ort(a)),
        'schulsozialarbeit': sum(1 for a in angebote if a['kategorie'] == 'Schulsozialarbeit' and vor_ort(a)),
        'jugend': sum(1 for a in angebote if a['kategorie'] == 'Jugendzentrum/Jugendpflege' and vor_ort(a)),
        'angebote': len(angebote),
    }
    median_vorort = statistics.median(r['_dichte'] for r in gem if r['_index'] is not None)
    for r in gem:
        # Einstufung mit ungerundeten Werten
        r['hoher_bedarf_wenig_angebot'] = bool(r['_index'] is not None and r['_index'] > 0 and r['_dichte'] < median_vorort)
    for r in gem:
        r.pop('_index'); r.pop('_dichte')
    median_vorort = round(median_vorort, 2)
    return {
        'gemeinden': gem, 'vergleich': vergleich, 'hze': hze_rows, 'versorgungsgrade': vg, 'kreis': kreis,
        'median_vorort': median_vorort,
        'angebote': [{k: a.get(k, '') for k in ['kategorie', 'name', 'traeger', 'kommune', 'weitere_kommunen', 'reichweite', 'zielgruppe', 'url']} for a in angebote],
        'schulen_liste': [{k: s.get(k, '') for k in ['kommune', 'name', 'schulart', 'schulform', 'ganztag_kurz']} for s in schulen],
        'index_def': {'kennzahlen': [c for c, _ in INDEX], 'richtung': {c: s for c, s in INDEX}, 'min': INDEX_MIN,
                      'mittel': {c: round(stats[c][0], 3) for c, _ in INDEX}, 'sd': {c: round(stats[c][1], 3) for c, _ in INDEX}},
    }


def kennzahlen_csv(d):
    spalten = ['gemeinde', 'einwohner', 'bedarfsindex', 'rang', 'n_ind', 'sgb2', 'kinderarmut', 'jugendarmut', 'altersarmut',
               'alo_svb', 'alo_svb_u25', 'kita_mh', 'kita3_mh', 'kita_mh_kreis', 'sgb2_ausl', 'nichtdt', 'kinder_u15_nichtdt',
               'einkuenfte_je_stpfl', 'ab65', 'u6', 'geburten_je1000', 'wanderungssaldo_je1000', 'veraend_vj', 'whg_klein',
               'efh_anteil', 'kita_u3', 'kita_3_5', 'schulen', 'grundschulen', 'weiterfuehrend', 'ganztag', 'hausaerzte',
               'kinderaerzte', 'mittelbereich', 'hausarzt_vg', 'familienzentren', 'schulsozialarbeit', 'jugend', 'vorort',
               'vorort_je10k', 'hoher_bedarf_wenig_angebot', 'fachdienste', 'psychiatrie']
    with open(os.path.join(DATA, 'kennzahlen_kommunen.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f)
        w.writerow(spalten)
        for r in d['gemeinden'] + d['vergleich']:
            w.writerow(['' if r.get(c) is None else r.get(c) for c in spalten])


def oeffentlich_pruefen():
    """Bricht ab, wenn Daten im Klartext an einer Stelle liegen, die ins öffentliche Repository gelangen kann."""
    import subprocess
    ign = subprocess.run(['git', 'check-ignore', '-q', 'intern/daten/test.csv'], cwd=ROOT).returncode
    if ign != 0:
        sys.exit('intern/ ist nicht in .gitignore ausgeschlossen. Abbruch, damit keine Daten öffentlich werden.')
    verfolgt = subprocess.run(['git', 'ls-files', '--', '*.csv', '*.xlsx', '*.xls', '*.ods', 'intern', 'sozialraum'],
                              cwd=ROOT, capture_output=True, text=True).stdout.split()
    verboten = [v for v in verfolgt if v.startswith('intern/') or (v.startswith('sozialraum/') and v != 'sozialraum/index.html')
                or (v.endswith('.csv') and not v.startswith('werkzeuge/') and not v.startswith('quellen/'))]
    if verboten:
        sys.exit('Diese Dateien sind im Repository und könnten Daten im Klartext enthalten: ' + ', '.join(verboten) +
                 '\nBitte mit „git rm --cached <datei>“ herausnehmen und nach intern/ legen.')


def seite():
    oeffentlich_pruefen()
    passwort = passwort_lesen()
    if not os.path.isdir(DATA):
        entpacken(passwort)
    fehler, hinweise, info = pruefen()
    if fehler:
        bericht_schreiben(fehler, hinweise, info, ['Seite NICHT gebaut, erst Fehler beheben.'])
        sys.exit(1)
    d = berechnen()
    kennzahlen_csv(d)
    d['methodik'] = json.load(open(os.path.join(SRC, 'methodik.json'), encoding='utf-8'))
    d['dateien'] = {os.path.basename(p): open(p, encoding='utf-8-sig').read() for p in sorted(glob.glob(os.path.join(DATA, '*.csv')))}
    vorlage = open(os.path.join(SRC, 'vorlage.html'), encoding='utf-8').read()
    d3 = open(os.path.join(TOOLS, 'd3.min.js'), encoding='utf-8').read()
    daten = json.dumps(d, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    html = vorlage.replace('/*__D3__*/', d3).replace('/*__DATEN__*/null', daten)
    if '/*__D3__*/' in html or '/*__DATEN__*/' in html:
        sys.exit('Platzhalter in der Vorlage nicht gefunden')
    with open(OUT_PAGE, 'w', encoding='utf-8') as f:
        f.write(verschluesseln(html, passwort))
    archiv_status = packen(passwort)
    rang = [f"  {r['rang']:>2}. {r['gemeinde']:<14} {r['bedarfsindex']:+.2f}" for r in sorted(
        [r for r in d['gemeinden'] if r['rang']], key=lambda r: r['rang'])]
    bericht_schreiben(fehler, hinweise, info, ['Bedarfsindex (Rang, Kommune, Wert):'] + rang +
                      ['', f'Seite gebaut und verschlüsselt: sozialraum/index.html ({len(html) // 1024} KB)', archiv_status])


PASSWORT_DATEI = os.path.join(SRC, '.passwort')
PBKDF2_RUNDEN = 600000


def passwort_lesen():
    pw = os.environ.get('SOZIALRAUM_PASSWORT', '').strip()
    if not pw and os.path.exists(PASSWORT_DATEI):
        pw = open(PASSWORT_DATEI, encoding='utf-8').read().strip()
    if len(pw) < 8:
        sys.exit('Kein Passwort (mind. 8 Zeichen): SOZIALRAUM_PASSWORT setzen oder werkzeuge/sozialraum/.passwort anlegen.')
    return pw


def _schluessel(passwort, salz, runden):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salz, iterations=runden).derive(passwort.encode())


def _zu(klartext, passwort):
    """AES-256-GCM, Schlüssel per PBKDF2-SHA256. Ergebnis als JSON-Text (base64)."""
    import base64
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salz, iv = os.urandom(16), os.urandom(12)
    geheim = AESGCM(_schluessel(passwort, salz, PBKDF2_RUNDEN)).encrypt(iv, klartext, None)
    b64 = lambda b: base64.b64encode(b).decode()
    return json.dumps({'salz': b64(salz), 'iv': b64(iv), 'runden': PBKDF2_RUNDEN, 'daten': b64(geheim)})


def _auf(paket, passwort):
    import base64
    from cryptography.exceptions import InvalidTag
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    p = json.loads(paket)
    d = lambda k: base64.b64decode(p[k])
    try:
        return AESGCM(_schluessel(passwort, d('salz'), p['runden'])).decrypt(d('iv'), d('daten'), None)
    except InvalidTag:
        sys.exit('Falsches Passwort für das Datenarchiv.')


def _archiv_inhalt():
    """Alle Dateien unter intern/ außer dem Prüfbericht, als sortierte Liste (Pfad, Bytes)."""
    teile = []
    for wurzel, _, dateien in os.walk(INTERN):
        for n in dateien:
            pfad = os.path.join(wurzel, n)
            rel = os.path.relpath(pfad, INTERN).replace(os.sep, '/')
            if rel != 'bericht.txt':
                teile.append((rel, open(pfad, 'rb').read()))
    return sorted(teile)


def packen(passwort):
    import io
    import zipfile
    inhalt = _archiv_inhalt()
    if os.path.exists(ARCHIV):
        alt = zipfile.ZipFile(io.BytesIO(_auf(open(ARCHIV, encoding='utf-8').read(), passwort)))
        if sorted((n, alt.read(n)) for n in alt.namelist()) == inhalt:
            return 'Datenarchiv unverändert.'
    puffer = io.BytesIO()
    with zipfile.ZipFile(puffer, 'w', zipfile.ZIP_DEFLATED) as z:
        for rel, daten in inhalt:
            z.writestr(zipfile.ZipInfo(rel, (2020, 1, 1, 0, 0, 0)), daten)
    with open(ARCHIV, 'w', encoding='utf-8') as f:
        f.write(_zu(puffer.getvalue(), passwort))
    return f'Datenarchiv aktualisiert: werkzeuge/sozialraum/daten.verschluesselt ({len(inhalt)} Dateien).'


def entpacken(passwort=None):
    import io
    import zipfile
    passwort = passwort or passwort_lesen()
    if not os.path.exists(ARCHIV):
        sys.exit('Kein Datenarchiv werkzeuge/sozialraum/daten.verschluesselt gefunden.')
    z = zipfile.ZipFile(io.BytesIO(_auf(open(ARCHIV, encoding='utf-8').read(), passwort)))
    for n in z.namelist():
        ziel = os.path.join(INTERN, *n.split('/'))
        os.makedirs(os.path.dirname(ziel), exist_ok=True)
        with open(ziel, 'wb') as f:
            f.write(z.read(n))
    print(f'{len(z.namelist())} Dateien nach intern/ entpackt.')


def verschluesseln(html, passwort):
    """Seite verschlüsseln; entschlüsselt wird im Browser (login.html)."""
    paket = _zu(html.encode('utf-8'), passwort)
    login = open(os.path.join(SRC, 'login.html'), encoding='utf-8').read()
    if '/*__PAKET__*/null' not in login:
        sys.exit('Platzhalter in login.html nicht gefunden')
    return login.replace('/*__PAKET__*/null', paket)


def hsl():
    import openpyxl
    dateien = sorted(glob.glob(os.path.join(ROOT, 'quellen', 'sozialraum', 'HGSt_j*.xlsx')))
    if not dateien:
        sys.exit('Keine Datei quellen/sozialraum/HGSt_j<JJJJ>.xlsx gefunden')
    pfad = dateien[-1]
    wb = openpyxl.load_workbook(pfad, read_only=True, data_only=True)

    def blatt(name):
        return {z[0]: z for z in wb[name].iter_rows(values_only=True)
                if z and isinstance(z[0], str) and (z[0].startswith('532') or z[0] == '000000')}

    b1, b3, b4, b13, b20 = (blatt(n) for n in ('1', '3', '4', '13', '20'))
    kopf = ['ags', 'gemeinde', 'einwohner', 'flaeche', 'dichte', 'veraend_vj', 'u6', 'u6_15', 'ab65', 'nichtdt',
            'kinder_u15_nichtdt', 'geburten_je1000', 'wanderungssaldo_je1000', 'einkuenfte_je_stpfl', 'whg_klein']
    zeilen = []
    for ags, r in b1.items():
        name = 'Hessen' if ags == '000000' else r[1].strip().split(',')[0]
        a3, a4, w, t = b3[ags], b4[ags], b13[ags], b20[ags]
        ew = r[3]
        zeilen.append([ags, name, ew, r[2], round(r[11]), round(r[9], 2), round(r[13], 1), round(r[15], 1), round(r[19], 1),
                       round(a3[9], 1), round((a3[10] + a3[12]) / (r[12] + r[14]) * 100, 1),
                       round(a4[2] / ew * 1000, 1), round(a4[15] / ew * 1000, 1), t[4], round((w[3] + w[4]) / w[2] * 100, 1)])
    os.makedirs(DATA, exist_ok=True)
    ziel = os.path.join(DATA, 'hsl_gemeindestatistik.csv')
    with open(ziel, 'w', encoding='utf-8', newline='') as f:
        wr = csv.writer(f)
        wr.writerow(kopf)
        wr.writerows(zeilen)
    print(f'{os.path.basename(pfad)} -> intern/daten/hsl_gemeindestatistik.csv ({len(zeilen)} Zeilen)')


if __name__ == '__main__':
    befehl = sys.argv[1] if len(sys.argv) > 1 else ''
    if befehl == 'hsl':
        hsl()
    elif befehl == 'entpacken':
        entpacken()
    elif befehl == 'pruefen':
        f, h, i = pruefen()
        bericht_schreiben(f, h, i)
        sys.exit(1 if f else 0)
    elif befehl == 'seite':
        seite()
    else:
        print(__doc__)
        sys.exit(2)
