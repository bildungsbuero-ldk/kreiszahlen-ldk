# Kreiszahlen Lahn-Dill-Kreis

Dashboard zu den Hessischen Kreiszahlen des Hessischen Statistischen Landesamts (Excel-Ausgaben ab Frühjahr 2023):
Bildungs- und Strukturdaten des Lahn-Dill-Kreises im Vergleich der hessischen Landkreise und kreisfreien Städte.

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/

- `index.html` – die Seite (eigenständig, lädt nichts von fremden Servern)
- `daten.json` – alle Werte, Regionen, Merkmale und Berichtszeiträume
- `quellen/` – die unveränderten Excel-Dateien des Statistischen Landesamts
- `werkzeuge/` – Aufbereitung und Seitenbau; Anleitung für neue Ausgaben in `werkzeuge/ANLEITUNG.md`

Quelle: Hessisches Statistisches Landesamt, Hessische Kreiszahlen (https://www.statistischebibliothek.de/mir/receive/HESerie_mods_00000870). Kennzahlen und Darstellung: eigene Berechnungen.

## Sozialräume: Bedarfe und Angebote

Übersicht der 23 Städte und Gemeinden im Lahn-Dill-Kreis als Grundlage für Sozialraumkonferenzen:
Bedarfsindex aus Sozialdaten, Angebote und Infrastruktur (Schwerpunkt psychische Gesundheit), Methodik und Quellen.

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/sozialraum/

- `sozialraum/index.html` – die Seite (gebaut mit `python3 werkzeuge/sozialraum.py seite`)
- `sozialraum/daten/` – alle Daten als CSV
- `sozialraum/vorlagen/` – Vorlagen für Einreichungen; eingereicht wird über das [GitHub-Formular](https://github.com/bildungsbuero-ldk/kreiszahlen-ldk/issues/new?template=datenlieferung.yml)
- `werkzeuge/sozialraum/` – Vorlage der Seite, Methodik-Katalog und Anleitung (`ANLEITUNG.md`)
- `quellen/sozialraum/` – unveränderte Originaldateien amtlicher Quellen

Quellen: Hessisches Statistisches Landesamt, Bertelsmann Stiftung (Wegweiser Kommune, CC0), Kassenärztliche Vereinigung Hessen,
Lahn-Dill-Kreis und eigene Recherche. Einzelnachweise auf der Seite unter „Methodik und Quellen“.
