# Kreiszahlen Lahn-Dill-Kreis

Dashboard zu den Hessischen Kreiszahlen des Hessischen Statistischen Landesamts (Excel-Ausgaben ab Frühjahr 2023):
Bildungs- und Strukturdaten des Lahn-Dill-Kreises im Vergleich der hessischen Landkreise und kreisfreien Städte.

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/

- `index.html` – die Seite (eigenständig, lädt nichts von fremden Servern)
- `daten.json` – alle Werte, Regionen, Merkmale und Berichtszeiträume
- `quellen/` – die unveränderten Excel-Dateien des Statistischen Landesamts
- `werkzeuge/` – Aufbereitung und Seitenbau; Anleitung für neue Ausgaben in `werkzeuge/ANLEITUNG.md`

Quelle: Hessisches Statistisches Landesamt, Hessische Kreiszahlen (https://www.statistischebibliothek.de/mir/receive/HESerie_mods_00000870). Kennzahlen und Darstellung: eigene Berechnungen.

## Gemeindezahlen

Die 23 Städte und Gemeinden des Lahn-Dill-Kreises im Vergleich, nach der Hessischen Gemeindestatistik (Excel-Ausgaben ab 2021),
mit Lahn-Dill-Kreis und Land Hessen als Bezugswerten. Gleicher Aufbau wie die Kreiszahlen; oben rechts schaltet man zwischen beiden Seiten um.

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/gemeinden.html

- `gemeinden.html`, `gemeinden.json` – Seite und Daten
- `quellen/HGSt_j<JJJJ>.xlsx` – die unveränderten Excel-Dateien
- `werkzeuge/gemeindezahlen.py` – Aufbereitung (Seitenbau gemeinsam mit den Kreiszahlen über `werkzeuge/kreiszahlen.py seite`)

Quelle: Hessisches Statistisches Landesamt, Hessische Gemeindestatistik (https://www.statistischebibliothek.de/mir/receive/HESerie_mods_00000449, aktuell auch https://statistik.hessen.de/publikationen/hessische-gemeindestatistik).

## Sozialräume: Bedarfe und Angebote

Übersicht der 23 Städte und Gemeinden im Lahn-Dill-Kreis als Grundlage für Sozialraumkonferenzen:
Bedarfsindex aus Sozialdaten, Angebote und Infrastruktur (Schwerpunkt psychische Gesundheit), Methodik und Quellen.

Seite: https://bildungsbuero-ldk.github.io/kreiszahlen-ldk/sozialraum/ (passwortgeschützt; Passwort beim Bildungsbüro)

- `sozialraum/index.html` – die Seite, verschlüsselt (gebaut mit `python3 werkzeuge/sozialraum.py seite`)
- `werkzeuge/sozialraum/daten.verschluesselt` – alle Daten, verschlüsselt. Im Klartext liegen sie nur lokal in `intern/` und nie im Repository
- Ergänzungen schicken Nutzer:innen an das Bildungsbüro; eingearbeitet wird nach Freigabe (siehe Anleitung)
- `werkzeuge/sozialraum/` – Vorlage der Seite, Methodik-Katalog und Anleitung (`ANLEITUNG.md`)
- `quellen/sozialraum/` – unveränderte Originaldateien amtlicher Quellen

Quellen: Hessisches Statistisches Landesamt, Bertelsmann Stiftung (Wegweiser Kommune, CC0), Kassenärztliche Vereinigung Hessen,
Lahn-Dill-Kreis und eigene Recherche. Einzelnachweise auf der Seite unter „Methodik und Quellen“.
