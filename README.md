# MM Risikoindikator Dashboard – HTML-Version mit Data.xlsx

Diese Version läuft über GitHub Pages und liest `Data.xlsx` direkt im Browser ein.

## Dateien

```text
index.html
styles.css
app.js
Data.xlsx
README.md
.gitignore
```

## Wichtig

- `Data.xlsx` muss im gleichen Ordner liegen wie `index.html`.
- Neue Risikoindikatoren oder Global-Index-Spalten werden automatisch erkannt, wenn die Struktur gleich bleibt:
  - Sheet `data (risk measures)`
  - Sheet `data (global indices)`
  - Spalte A = Datum
  - Zeile 6 = Name der Zeitreihe
  - ab Zeile 7 = Werte
- GitHub Pages führt kein Python aus. Deshalb rechnet diese Version mit JavaScript im Browser.
- Die Excel-Datei ist bei GitHub Pages öffentlich abrufbar. Keine vertraulichen Daten öffentlich hochladen.

## GitHub Pages starten

1. Repository erstellen.
2. Diese Dateien hochladen.
3. In GitHub: `Settings` → `Pages`.
4. Source: `Deploy from a branch`.
5. Branch: `main`, Folder: `/root`.
6. Speichern.

Danach ist die Seite über den GitHub-Pages-Link erreichbar.
