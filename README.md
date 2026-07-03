# MM Risikoindikator Dashboard – HTML-Version mit `Data.xlsx`

Diese Version läuft statisch über **GitHub Pages**. Sie liest `Data.xlsx` im Browser ein; ein Python- oder Streamlit-Server ist nicht nötig.

## Dateien

```text
index.html    # Dashboard-Oberfläche
styles.css    # Layout und responsive Darstellung
app.js        # Excel-Import, Berechnungen und Diagramme
Data.xlsx     # Zeitreihen und Metadaten
```

## Wichtige Logik

- **Z-Score-Basis:** ausschließlich **Level**.
- **Score-Methode:** ausschließlich **Normal-CDF Wahrscheinlichkeit**.
- **Gewichte:** Eingabe als Prozentwerte. Bei allen fünf aktuellen Indikatoren sind anfangs jeweils 20 % gesetzt.
- **Neue Risikoindikatoren:** Eine zusätzliche Spalte in `data (risk measures)` wird automatisch erkannt, wenn
  - Zeile 6 einen Namen enthält und
  - ab Zeile 7 mindestens ein numerischer Wert vorhanden ist.
  Sie erscheint automatisch in der Risk-Measure-Liste und erhält beim ersten Laden ein gleichgewichtet vorbelegtes Prozentgewicht.

## Excel-Struktur

- Sheet `data (risk measures)`: Risikoindikatoren
- Sheet `data (global indices)`: Vergleichsindizes
- Spalte A: Datum
- Zeile 6: Dashboard-Name der Zeitreihe
- Ab Zeile 7: Zeitreihenwerte

## GitHub Pages

1. Dateien in das Repository hochladen.
2. Repository auf **Public** stellen, sofern GitHub Pages in deinem Tarif für private Repositories nicht verfügbar ist.
3. Unter **Settings → Pages** einstellen:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/ (root)`
4. Nach dem Deployment wird die Seite über die angezeigte GitHub-Pages-URL aufgerufen.

## Hinweis zu Daten

`Data.xlsx` wird für jeden Besucher der GitHub-Pages-Seite abrufbar. Keine Bloomberg- oder sonstigen nicht öffentlich weitergabefähigen Daten in ein öffentliches Repository hochladen.


## Kalman Filter

Der Tab **Kalman Filter** zeigt zwei Varianten des Risikoindikators:

- **Normal-CDF (ungefiltert):** Normal-CDF des täglichen Composite Z-Scores.
- **Normal-CDF (Kalman-gefiltert):** Zuerst wird der Composite Z-Score mit einem kausalen Local-Level-Kalman-Filter geglättet, danach wird die Normal-CDF berechnet.

Der Filter nutzt für den jeweiligen Handelstag nur historische Werte bis einschließlich dieses Tages. In der Sidebar kann die Strategie optional den gefilterten statt des ungefilterten Risk Indicators verwenden.
