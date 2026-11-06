# Sozialraum-Übersicht pflegen

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/sozialraum/

## Aufbau
- `sozialraum/index.html` – die Seite (gebaut, nicht von Hand bearbeiten)
- `sozialraum/daten/` – alle Daten als CSV (UTF-8, Komma, Dezimalpunkt). Diese Dateien sind die Grundlage der Seite.
  `kennzahlen_kommunen.csv` wird vom Werkzeug erzeugt.
- `sozialraum/vorlagen/` – Vorlagen für Einreichungen
- `sozialraum/einreichungen/` – freigegebene Einreichungen im Original
- `quellen/sozialraum/` – Originaldateien amtlicher Quellen (Gemeindestatistik, Statistischer Bericht K V 8)
- `werkzeuge/sozialraum.py` – Prüfen, Berechnen, Seitenbau
- `werkzeuge/sozialraum/methodik.json` – Definitionen, Quellen und Lücken. Daraus entstehen die Info-Popups und die Seite „Methodik und Quellen“.
  Jede neue Kennzahl braucht hier einen Eintrag.
- `werkzeuge/sozialraum/vorlage.html` – Vorlage der Seite (nutzt `werkzeuge/d3.min.js`)

## Seite neu bauen
```
python3 werkzeuge/sozialraum.py pruefen   # nur prüfen
python3 werkzeuge/sozialraum.py seite     # prüfen und bauen
```
Das Werkzeug prüft Spalten, Kommunennamen, Zahlenbereiche, Kategorien und doppelte Einträge. Bei Fehlern baut es
nichts. Bericht: `werkzeuge/ausgabe/sozialraum_bericht.txt` (mit Rangfolge des Bedarfsindex). Danach Seite lokal öffnen
(`python3 -m http.server` im Hauptordner, dann `/sozialraum/`) und alle vier Reiter ansehen.

## Eingereichte Daten verarbeiten
Einreichungen kommen über das GitHub-Formular „Daten für die Sozialraum-Übersicht einreichen“ (Label `datenlieferung`).

1. **Sichten**: Was wird geliefert, für welche Kommunen, Quelle, Stand, Veröffentlichung erlaubt? Personenbezogene
   Daten oder Werte unter 3 Fällen nicht übernehmen; im Zweifel nachfragen.
2. **Rücksprache**: Vorschlag für die Einarbeitung (welche Datei, welche Spalte, neue Kennzahl ja/nein, Änderung am
   Bedarfsindex ja/nein) im Formular-Thread notieren und vom Bildungsbüro freigeben lassen. Ohne Freigabe nichts einarbeiten.
3. **Einarbeiten** auf einem eigenen Branch:
   - Original nach `sozialraum/einreichungen/JJJJ-MM-TT_kurzbeschreibung.<endung>` legen.
   - Angebote: Zeile(n) in `angebote_liste.csv`; erlaubte Kategorien und `reichweite` stehen in `werkzeuge/sozialraum.py`.
   - Neue Kennzahl: eigene CSV-Datei mit Spalte `gemeinde` (alle 23 Kommunen, leere Werte erlaubt), Prüfregel in `SCHEMA`
     in `werkzeuge/sozialraum.py`, Berechnung in `berechnen()`, Eintrag in `methodik.json` (Definition, Jahr, Quelle),
     Anzeige in `vorlage.html` (Auswahlliste `EINZEL`, Tabellen).
   - Bedarfsindex nur nach ausdrücklicher Entscheidung ändern (`INDEX` in `werkzeuge/sozialraum.py`) und den Abschnitt
     in der Methodik anpassen.
   - `stand` in `methodik.json` aktualisieren.
4. **Bauen und prüfen** (siehe oben), Stichproben gegen die Einreichung vergleichen.
5. **Pull Request** mit Verweis auf das Formular („Schließt #…“). Nach dem Zusammenführen veröffentlicht GitHub Pages
   die Seite automatisch. Formular mit kurzer Rückmeldung schließen.

## Regelmäßige Aktualisierung
- **Gemeindestatistik** (jährlich, Herbst): neue `HGSt_j<JJJJ>.xlsx` aus der Statistischen Bibliothek
  (https://www.statistischebibliothek.de, Suche „Hessische Gemeindestatistik“) nach `quellen/sozialraum/` legen,
  `python3 werkzeuge/sozialraum.py hsl` ausführen. Prüfen, ob die Blätter 1, 3, 4, 13 und 20 noch denselben Aufbau haben.
- **Wegweiser Kommune** (jährlich): Export-Schnittstelle
  `https://www.wegweiser-kommune.de/data-api/rest/export/soziale-lage+integration+<kommune>+<jahr>+tabelle.csv?charset=UTF-8`,
  Kommunen-Kürzel z. B. `wetzlar`, `herborn-ldk`, `breitscheid-ldk`, Kreis `lahn-dill-kreis-lk`, Land `hessen`.
  Kita-Quoten: Indikatoren `unter-3-jaehrige-in-tageseinrichtungen` und `3-bis-5-jaehrige-in-tageseinrichtungen`.
- **KV Hessen** (halbjährlich): Kartendienste unter https://www.kvhessen.de/fokus-gesundheit/ (ArcGIS-Abfragen
  `hausaerzte`, `allgemeine_fachaerzte`, `spezialisierte_fachaerzte`).
- **Hilfen zur Erziehung**: Statistischer Bericht K V 8 enthält seit 2024 nur Landeswerte; Werte nach Jugendamt
  oder Kommune bei der Jugendhilfeplanung anfragen.
