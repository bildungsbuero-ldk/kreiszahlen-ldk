# Sozialraum-Übersicht pflegen

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/sozialraum/ (passwortgeschützt)

## Datenschutz: was öffentlich ist und was nicht
Das Repository und GitHub Pages sind **öffentlich**. Deshalb gilt:
- Daten liegen im Klartext **nur lokal** im Ordner `intern/` (in `.gitignore` ausgeschlossen, wird nie hochgeladen).
- Im Repository liegen sie nur verschlüsselt: `werkzeuge/sozialraum/daten.verschluesselt` (AES-256-GCM, Schlüssel per
  PBKDF2-SHA256 aus dem Passwort). Auch die Seite `sozialraum/index.html` ist verschlüsselt und wird erst im Browser
  nach Eingabe des Passworts entschlüsselt. Die CSV-Downloads stecken in der verschlüsselten Seite.
- **Nicht öffentliche Daten** (z. B. Hilfen zur Erziehung nach Kommune, Zahlen der Jugendhilfeplanung oder des
  Gesundheitsamts) gehören ausschließlich nach `intern/`. Nie als CSV, Excel, Screenshot oder Zahl in Commit-Texten,
  Pull Requests, Issues, README oder `methodik.json` (die Methodik beschreibt nur, was gezählt wird, keine Werte).
- Das Werkzeug bricht ab, wenn `intern/` nicht ausgeschlossen ist oder Datendateien (CSV/Excel) unter `sozialraum/` oder
  `intern/` im Repository liegen.
- Hinweis: Ältere Commits (bis 06.11.2026) enthalten die damaligen CSV-Dateien im Klartext. Sie stammen alle aus
  öffentlichen Quellen.

## Passwort
- Das Passwort verwaltet das Bildungsbüro. Es steht nirgends im Repository.
- Zum Bauen: Umgebungsvariable `SOZIALRAUM_PASSWORT` setzen oder die Datei `werkzeuge/sozialraum/.passwort`
  (eine Zeile, in `.gitignore` ausgeschlossen) anlegen.
- **Passwort ändern**: `intern/` entpacken (mit altem Passwort), neues Passwort eintragen, Archiv
  `werkzeuge/sozialraum/daten.verschluesselt` löschen und `seite` ausführen. Dann werden Seite und Archiv mit dem neuen
  Passwort neu verschlüsselt. Wer „angemeldet bleiben“ gewählt hat, muss sich danach neu anmelden (das passiert auch
  nach jedem Neubau).
- Ohne Passwort lassen sich die Daten nicht wiederherstellen.

## Aufbau
- `sozialraum/index.html` – die Seite (gebaut und verschlüsselt, nicht von Hand bearbeiten)
- `intern/daten/` – lokal: alle Daten als CSV (UTF-8, Komma, Dezimalpunkt). `kennzahlen_kommunen.csv` erzeugt das Werkzeug.
- `intern/einreichungen/` – lokal: freigegebene Zusendungen im Original (`JJJJ-MM-TT_kurzbeschreibung.<endung>`)
- `intern/bericht.txt` – lokal: Prüfbericht mit Rangfolge des Bedarfsindex
- `werkzeuge/sozialraum/daten.verschluesselt` – verschlüsseltes Archiv von `intern/` (ohne Bericht)
- `quellen/sozialraum/` – Originaldateien amtlicher, öffentlicher Quellen (Gemeindestatistik, Statistischer Bericht K V 8)
- `werkzeuge/sozialraum.py` – Entpacken, Prüfen, Berechnen, Seitenbau, Verschlüsseln
- `werkzeuge/sozialraum/methodik.json` – Definitionen, Quellen und Lücken. Daraus entstehen die Info-Popups und die Seite „Methodik und Quellen“.
  Jede neue Kennzahl braucht hier einen Eintrag.
- `werkzeuge/sozialraum/vorlage.html` – Vorlage der Seite (nutzt `werkzeuge/d3.min.js`); `login.html` – Anmeldemaske

## Seite neu bauen
```
python3 werkzeuge/sozialraum.py entpacken  # Daten nach intern/ holen (passiert bei `seite` automatisch, wenn intern/ fehlt)
python3 werkzeuge/sozialraum.py pruefen    # nur prüfen
python3 werkzeuge/sozialraum.py seite      # prüfen, bauen, verschlüsseln, Archiv aktualisieren
```
Das Werkzeug prüft Spalten, Kommunennamen, Zahlenbereiche, Kategorien und doppelte Einträge. Bei Fehlern baut es
nichts. Danach Seite lokal öffnen (`python3 -m http.server` im Hauptordner, dann `/sozialraum/`), anmelden und alle
vier Reiter ansehen. Committet werden nur `sozialraum/index.html`, `werkzeuge/sozialraum/daten.verschluesselt` und
geänderte Werkzeug-/Vorlagendateien.

## Zugeschickte Daten verarbeiten
Ergänzungen schicken Nutzer:innen formlos an das Bildungsbüro (E-Mail, Excel, Liste). Das Bildungsbüro leitet sie zur
Einarbeitung weiter.

1. **Sichten**: Was wird geliefert, für welche Kommunen, Quelle, Stand, Veröffentlichung erlaubt? Personenbezogene
   Daten oder Werte unter 3 Fällen nicht übernehmen; im Zweifel nachfragen.
2. **Rücksprache**: Vorschlag für die Einarbeitung (welche Datei, welche Spalte, neue Kennzahl ja/nein, Änderung am
   Bedarfsindex ja/nein) dem Bildungsbüro vorlegen und freigeben lassen. Ohne Freigabe nichts einarbeiten.
3. **Einarbeiten** auf einem eigenen Branch:
   - Original nach `intern/einreichungen/JJJJ-MM-TT_kurzbeschreibung.<endung>` legen.
   - Angebote: Zeile(n) in `intern/daten/angebote_liste.csv`; erlaubte Kategorien und `reichweite` stehen in `werkzeuge/sozialraum.py`.
   - Neue Kennzahl: eigene CSV-Datei mit Spalte `gemeinde` (alle 23 Kommunen, leere Werte erlaubt), Prüfregel in `SCHEMA`
     in `werkzeuge/sozialraum.py`, Berechnung in `berechnen()`, Eintrag in `methodik.json` (Definition, Jahr, Quelle),
     Anzeige in `vorlage.html` (Auswahlliste `EINZEL`, Tabellen).
   - Bedarfsindex nur nach ausdrücklicher Entscheidung ändern (`INDEX` in `werkzeuge/sozialraum.py`) und den Abschnitt
     in der Methodik anpassen.
   - `stand` in `methodik.json` aktualisieren.
4. **Bauen und prüfen** (siehe oben), Stichproben gegen die Einreichung vergleichen.
5. **Pull Request** mit kurzer Beschreibung der Zusendung (ohne Namen, Kontaktdaten oder Werte). Nach dem Zusammenführen
   veröffentlicht GitHub Pages die Seite automatisch.

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
