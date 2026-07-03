# MM Risikoindikator Dashboard – HTML/XLSX/Kalman/MSCI-Optimierung

Statische Browser-Version für GitHub Pages. Die Website liest `Data.xlsx` direkt im Browser ein. Python oder Streamlit werden nicht benötigt.

## Dateien

```text
index.html      Oberfläche
app.js          Datenimport, Risikoindikator, Kalman-Filter und Optimierung
styles.css      Layout
Data.xlsx       Risikoindikatoren und Global Indices
```

## Excel-Struktur

Die Datei `Data.xlsx` muss im selben Ordner liegen wie `index.html`.

Erwartete Sheets:

```text
data (risk measures)
data (global indices)
```

- Spalte A: Datum
- Zeile 6: Anzeigename der Zeitreihe
- Ab Zeile 7: Werte

Neue Risikoindikatoren werden automatisch erkannt, sobald sie als zusätzliche Spalte im Sheet `data (risk measures)` ergänzt und in Zeile 6 benannt werden.

## Gewichtsoptimierung auf MSCI World

Der Bereich **Gewichte optimieren** maximiert die historische annualisierte Sharpe Ratio einer Long/Cash-Strategie auf Basis der Zeitreihe `MSCI WORLD` im Sheet `data (global indices)`.

- Optimiert werden nur die aktuell aktivierten Risk Measures.
- Die aktuelle Schwelle, die Vortags-Signal-Option und optional der Kalman-gefilterte Risk Indicator werden in der Optimierung berücksichtigt.
- Das maximale Gewicht je Indikator begrenzt Konzentrationsrisiken.
- Nach dem Klick werden die gefundenen Gewichte als Prozentwerte fest in die Gewichtsfelder geschrieben.

**Wichtig:** Das Ergebnis ist eine historische In-Sample-Optimierung bis zum ausgewählten Trainingsende. Es zeigt daher nicht automatisch eine künftig robuste Strategie. Für einen unverzerrten Test sollte das Trainingsende vor dem späteren Auswertungszeitraum liegen, beispielsweise Training bis 2022 und anschließende Prüfung ab 2023.

## GitHub Pages

1. Repository auf GitHub öffnen.
2. `Settings` → `Pages`.
3. `Deploy from a branch` auswählen.
4. Branch `main`, Ordner `/(root)` auswählen.
5. Speichern.

Das Repository muss für GitHub Pages im kostenlosen GitHub-Tarif öffentlich sein.
