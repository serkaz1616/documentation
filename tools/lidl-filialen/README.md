# Lidl-Filialdaten von lidl.de

Holt Öffnungszeiten, Adresse und Ausstattung aller deutschen Lidl-Filialen direkt von
lidl.de und schreibt sie als zusätzliche Spalten in die Filial-Excel.
Die Daten kommen nur von Lidl, nicht von Google oder anderen Quellen, die Dritte ändern können.

## Woher die Daten kommen

- Liste aller Filialseiten: `https://www.lidl.de/s/de-DE/filialen/sitemap.xml`
- Je Filialseite liefert lidl.de eine Datendatei (`…/_payload.json`) mit:
  - `objectNumber`, z. B. `DE00668`, also Filialnummer 668. Darüber läuft die Zuordnung zur Excel.
  - Adresse, Bundesland und Koordinaten
  - Regelöffnungszeiten je Wochentag und Sonderöffnungszeiten (z. B. Feiertage)
  - Ausstattung (`infoIcons`: Paketstation, E-Ladesäule, Self-Checkout …)
  - Vertriebsgesellschaft (Regionalgesellschaft)

## Installation

```bash
pip install requests openpyxl
```

## Benutzung

```bash
# 1) Daten aller Filialen laden (~3.300 Seiten, ca. 10–15 Min.)
#    Kann man abbrechen und neu starten: bereits geladene Filialen werden übersprungen.
python lidl_filialen.py scrape

# 2) In die Excel schreiben (dauert Sekunden, ohne neuen Abruf)
python lidl_filialen.py merge --input Filialen_DE.xlsx --output Filialen_DE_lidl.xlsx

# oder beides in einem Lauf
python lidl_filialen.py all --input Filialen_DE.xlsx --output Filialen_DE_lidl.xlsx

# alles neu laden (z. B. monatlich)
python lidl_filialen.py scrape --refresh
```

Die Rohdaten liegen in `lidl_filialen_cache.json`. Weitere Felder für den Konfigurator
kommen in `lidl_columns()` dazu. Danach reicht ein neues `merge`, ohne erneuten Abruf.

## Ergebnis-Excel

Hauptblatt: die Originalspalten bleiben unverändert, danach kommen:

| Spalte | Inhalt |
|---|---|
| Lidl gefunden | ja/nein (nein = Filiale nicht auf lidl.de, z. B. geschlossen oder Nummer falsch) |
| Lidl Objektnummer, Filialname | z. B. `DE00668`, `Mannheim-Tattersall` |
| Lidl Straße / PLZ / Ort / Bundesland | Adresse laut Lidl |
| Breitengrad / Längengrad | Koordinaten |
| Adressabgleich | `OK` oder `Abweichung: PLZ, Straße, Ort` (rot markiert) |
| Öffnungszeiten (Lidl) | Zusammenfassung, z. B. `Mo-Sa 07:00-22:00; So geschlossen` |
| Öffnung Mo … So | je Tag einzeln |
| Sonderöffnungszeiten | z. B. `01.11.2026 geschlossen` |
| Ausstattung | z. B. `eCharger, package, shopNGo` |
| Vertriebsgesellschaft (Lidl), Lidl URL, Abgerufen am | |
| Möglicher Treffer | bei nicht gefundenen Filialen: Lidl-Filiale mit gleicher PLZ und ähnlicher Straße (vermutlich umnummeriert) |

Zusatzblätter:
- **Nur auf lidl.de**: Filialen, die Lidl listet, die aber nicht in der Excel stehen
- **Info Lidl-Abgleich**: Stand und Zahlen des Abgleichs
