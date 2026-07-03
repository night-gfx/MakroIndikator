# MM Risikoindikator Dashboard

Streamlit-Dashboard zur Berechnung eines Risikoindikators aus Rolling Z-Scores und zum Vergleich mit Global Indices.

## Dateien

```text
risikoindikator-dashboard/
├── app.py              # Streamlit-App
├── Data.xlsx           # Input-Daten
├── requirements.txt    # Python-Abhängigkeiten
├── .gitignore
└── README.md
```

## Lokal starten

### 1. Repository klonen oder Ordner öffnen

```bash
cd risikoindikator-dashboard
```

### 2. Virtuelle Umgebung erstellen

Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Mac/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Pakete installieren

```bash
pip install -r requirements.txt
```

### 4. Dashboard starten

```bash
python -m streamlit run app.py
```

Danach öffnet sich das Dashboard im Browser.

## Datenstruktur

Die Datei `Data.xlsx` muss im gleichen Ordner wie `app.py` liegen.

Erwartete Sheets:

- `data (risk measures)`
- `data (global indices)`

Erwartete Struktur:

- Spalte A: Datum
- Zeile 6: Name der Zeitreihe für das Dashboard
- Ab Zeile 7: Zeitreihenwerte

## GitHub hochladen

```bash
git init
git add .
git commit -m "Initial risk indicator dashboard"
git branch -M main
git remote add origin https://github.com/DEIN-USERNAME/risikoindikator-dashboard.git
git push -u origin main
```

`DEIN-USERNAME` durch deinen GitHub-Namen ersetzen.

## Streamlit Community Cloud Deployment

1. Repository auf GitHub hochladen.
2. Auf Streamlit Community Cloud einloggen.
3. Neues App-Deployment erstellen.
4. Repository auswählen.
5. Main file path setzen auf:

```text
app.py
```

6. Deploy starten.

## Hinweis

`Data.xlsx` wird aktuell mit ins Repository gelegt. Wenn die Datei vertrauliche Daten enthält, sollte sie nicht öffentlich auf GitHub hochgeladen werden.
