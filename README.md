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

## Kalman-Filter und adaptive Strategie-Grenze

Die aktuelle Version verwendet den **Kalman-gefilterten Composite Z-Score durchgängig**:

- Risk Indicator: Normal-CDF des Kalman-gefilterten Composite Z-Scores.
- Strategie: Signale beruhen ausschließlich auf diesem Kalman Risk Indicator.
- Gewichtsoptimierung: maximiert die historische Sharpe Ratio auf Basis dieses Kalman-Signals für den MSCI World.

Die **adaptive Grenze** wird täglich kausal berechnet:

- Ausgangspunkt ist die einstellbare Basis-Grenze (Standard: 60).
- Ein Risikotag liegt vor, wenn der Kalman Risk Indicator über der Basis-Grenze liegt.
- Die Grenze betrachtet die Risikotage der vorherigen 252 Handelstage.
- Wenige Risikotage im Rückblick → Grenze wird bis maximal 15 Prozentpunkte gesenkt.
- Viele Risikotage im Rückblick → Grenze wird bis maximal 15 Prozentpunkte angehoben.
- Die Grenze zum Zeitpunkt *t* nutzt nur Informationen bis *t-1* und verursacht daher keinen Look-Ahead-Bias.


## Historische, dynamische Risiko-Grenze

Die Strategie verwendet keine feste Schwelle mehr. Der Kalman Risk Indicator wird jeden Tag relativ zu seiner eigenen vorherigen Verteilung eingeordnet:

- **Niedrig:** unter dem historischen 33%-Quantil
- **Mäßig:** zwischen 33%-Quantil und Median
- **Erhöht:** zwischen Median und High-Risk-Quantil
- **Hoch:** oberhalb des gewählten High-Risk-Quantils

Die High-Risk-Linie ist standardmäßig das **80%-Quantil** der vorherigen 252 Handelstage. Sie verändert sich daher automatisch mit dem langfristigen Risikoniveau. Die Strategie ist long bis der Kalman Risk Indicator diese dynamische High-Risk-Linie überschreitet. Alle Bänder verwenden nur frühere Beobachtungen und sind damit kausal.


## Dynamische Ausschlags-Strategie

Die Strategie verwendet keine feste Risiko-Grenze. Sie berechnet aus dem Kalman-gefilterten Risikoindikator einen kausalen **Ausschlags-Score**: `absolute Bewegung` relativ zu den bislang typischen Bewegungen und `relative Veränderung` gegenüber dem vorherigen Niveau werden kombiniert. Die tägliche Trigger-Linie ist ein Quantil der vorangegangenen Ausschlags-Scores. Überschreitet der Score diese Linie, wird das Signal auf Nicht investiert gesetzt.

Beispiele: Ein Sprung von 20 auf 40 % ist wegen der hohen relativen Veränderung auffällig. Ein Anstieg von 60 auf 80 % kann ebenfalls ein Signal auslösen, wenn die absolute Bewegung im Vergleich zur jüngeren Historie ungewöhnlich groß ist.
