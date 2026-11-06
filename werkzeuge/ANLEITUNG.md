# Neue Ausgabe der Hessischen Kreiszahlen einspielen

Die Kreiszahlen erscheinen zweimal im Jahr (Frühjahrsausgabe meist Ende März, Herbstausgabe meist Ende Oktober).
Quelle: https://www.statistischebibliothek.de/mir/receive/HESerie_mods_00000870

## 1. Excel-Datei holen
1. Auf der Quellenseite die neueste Ausgabe öffnen (Liste „Inhalt“, z. B. „2027 Frühjahrsausgabe“).
2. Auf der Ausgabenseite liegen zwei Dateipakete („Derivate“): eines mit dem PDF, eines (Kennung „nxa2“) mit der
   Excel-Datei. Über `https://www.statistischebibliothek.de/mir/servlets/MCRFileNodeServlet/<derivat-id>/` lassen sich
   die Dateien auflisten, z. B. Herbst 2026: `HEHeft_derivate_00016729/HKr_26-2hj.xlsx`.
3. Die `.xlsx` unverändert nach `quellen/` legen und so benennen, dass das Werkzeug die Ausgabe erkennt:
   `HKr_<JJ>-1hj.xlsx` für die Frühjahrsausgabe, `HKr_<JJ>-2hj.xlsx` für die Herbstausgabe (JJ = zweistelliges Jahr).
   Gibt es nur ein PDF ohne Excel, nichts einspielen und das melden.

## 2. Daten und Seite neu erzeugen
```
pip install openpyxl --break-system-packages   # falls nicht vorhanden
python3 werkzeuge/kreiszahlen.py daten
python3 werkzeuge/kreiszahlen.py seite
```
`daten` schreibt `daten.json`, `werkzeuge/ausgabe/werte.csv`, `werkzeuge/ausgabe/bezug.csv` und `werkzeuge/ausgabe/bericht.txt`.

## 3. Prüfen, bevor etwas veröffentlicht wird
- Alle Werte der früheren Ausgaben müssen unverändert bleiben (vorherige `werkzeuge/ausgabe/werte.csv` aus Git mit der neuen vergleichen).
- `bericht.txt` lesen, besonders den Abschnitt der neuen Ausgabe: entfallene Tabellen, Spalten ohne Gegenstück,
  unsichere Zuordnungen und auffällige Sprünge. Jede unsichere Zuordnung gegen die Excel-Datei prüfen.
- Stichprobe: mindestens 30 Werte der neuen Ausgabe (darunter Tabellen 7, 11, 15–18, 53 für den Lahn-Dill-Kreis 532000
  und Land Hessen 000000) direkt mit den Zellen der Excel-Datei vergleichen.
- Seite lokal öffnen (`python3 -m http.server` im Hauptordner) und prüfen: oben steht die neue Ausgabe, Kacheln aller
  Themen zeigen Zahlen (kein „—“, kein NaN), Zeitverlauf hat einen Punkt mehr, keine Fehler in der Konsole.
- Hat das Statistische Landesamt Tabellen umnummeriert oder neu geschnitten, so dass Kacheln (werkzeuge/kacheln.json)
  keine Werte mehr bekommen: nicht raten, sondern melden.

## 4. Veröffentlichen
- Commit mit Autor und Committer `Bildungsbüro Lahn-Dill-Kreis <bildungsbuero-ldk@users.noreply.github.com>`, Push auf `main`.
  GitHub Pages veröffentlicht die Seite danach automatisch unter https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/
- Claude-Dashboard (https://claude.ai/artifact/2C9fY57hvinTL7t9ZVA316): `werkzeuge/ausgabe/werte.csv` und `bezug.csv` als
  Assets hochladen und in den Datensätzen `datasets/werte` und `datasets/bezug` jeweils `source.url`, `source.name` und
  `updated` ersetzen (vorher lesen, mit `if_version` schreiben). Die Beschreibung von `werte` nennt die Spanne der Ausgaben.
  Katalog und Regionen (`datasets/merkmale`, `datasets/regionen`) nur ersetzen, wenn sich `werkzeuge/katalog/` geändert hat.
  Kacheln und Bezugsgrößen sind im Dashboard fest in der Seite (`files/index.html`) hinterlegt; ändern sie sich, dort mit nachziehen.

## Aufbau
- `quellen/` – die Original-Excel-Dateien aller Ausgaben (unverändert).
- `werkzeuge/kreiszahlen.py` – Einlesen, Zuordnung zur Referenzausgabe Herbst 2026, Berichtszeiträume, Seitenbau.
- `werkzeuge/katalog/` – Merkmalskatalog und Regionen (Referenz Herbst 2026).
- `werkzeuge/kacheln.json` – Kennzahlen der Themen-Kacheln (Formeln über Merkmals-IDs).
- `werkzeuge/bezugsgroessen.json` – geprüfte Bezugsgrößen für den Prozent-Modus (Merkmal → Summe).
- `werkzeuge/vorlage.html`, `werkzeuge/d3.min.js` – Vorlage der Seite und Diagrammbibliothek (d3 v7.9.0, ISC-Lizenz).

## Sozialraum-Übersicht
Eigene Anleitung: `werkzeuge/sozialraum/ANLEITUNG.md`. Der Link von der Kreiszahlen-Seite steht in `werkzeuge/vorlage.html` (Kopfzeile).
