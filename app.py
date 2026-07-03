# ============================================================
# dashboard_combined.py
#
# Eine-Datei-Version des MM Risikoindikator Dashboards.
# Benötigt im gleichen Ordner: Data.xlsx
# Start: python -m streamlit run dashboard_combined.py
# ============================================================

# ============================================================
# BEGIN: data_loader.py
# ============================================================


from pathlib import Path
from typing import Dict, Tuple

import numpy as np
import pandas as pd


# ============================================================
# data_loader.py
#
# Lädt Data.xlsx aus dem gleichen Ordner wie dashboard.py.
#
# Erwartete Struktur:
# - Sheet "data (risk measures)" enthält Risikoindikatoren
# - Sheet "data (global indices)" enthält die zu timenden Indices
# - Spalte A: Datum
# - Zeile 6: Name der Zeitreihe für das Dashboard
# - Ab Zeile 7: Zeitreihenwerte
# ============================================================


RISK_SHEET = "data (risk measures)"
INDEX_SHEET = "data (global indices)"

FIELD_ROW = 0          # Excel-Zeile 1
TICKER_ROW = 1         # Excel-Zeile 2
REGION_ROW = 2         # Excel-Zeile 3
NAME_ROW = 3           # Excel-Zeile 4
TYPE_ROW = 4           # Excel-Zeile 5
SERIES_NAME_ROW = 5    # Excel-Zeile 6
DATA_START_ROW = 6     # Excel-Zeile 7


def find_data_file(base_dir: Path | None = None) -> Path:
    if base_dir is None:
        base_dir = Path(__file__).resolve().parent

    candidates = [
        "Data.xlsx",
        "Data.xlsm",
        "Data.xls",
        "data.xlsx",
        "data.xlsm",
        "data.xls",
    ]

    for filename in candidates:
        path = base_dir / filename
        if path.exists():
            return path

    raise FileNotFoundError(
        "Keine Excel-Datei gefunden. Bitte lege die Datei als 'Data.xlsx' "
        "in den gleichen Ordner wie dashboard.py."
    )


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def make_unique(names: list[str]) -> list[str]:
    counts = {}
    result = []

    for name in names:
        base = name if name else "Unnamed"
        counts[base] = counts.get(base, 0) + 1

        if counts[base] == 1:
            result.append(base)
        else:
            result.append(f"{base}_{counts[base]}")

    return result


def convert_to_number(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    s = series.astype(str).str.strip()

    s = s.replace(
        {
            "": np.nan,
            "nan": np.nan,
            "NaN": np.nan,
            "None": np.nan,
            "#N/A": np.nan,
            "#N/A Field Not Applicable": np.nan,
            "#VALUE!": np.nan,
        }
    )

    has_comma = s.str.contains(",", na=False)

    s.loc[has_comma] = (
        s.loc[has_comma]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(s, errors="coerce")


def load_timeseries_sheet(excel_path: Path, sheet_name: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )

    if raw.shape[0] <= DATA_START_ROW:
        raise ValueError(f"Sheet '{sheet_name}' hat zu wenige Zeilen.")

    dates = pd.to_datetime(
        raw.iloc[DATA_START_ROW:, 0],
        errors="coerce",
        dayfirst=True
    ).reset_index(drop=True)

    metadata_rows = []
    value_columns = []
    original_names = []

    for col_idx in range(1, raw.shape[1]):
        series_name = clean_text(raw.iat[SERIES_NAME_ROW, col_idx])

        if not series_name:
            continue

        values = convert_to_number(raw.iloc[DATA_START_ROW:, col_idx]).reset_index(drop=True)

        if values.dropna().empty:
            continue

        original_names.append(series_name)
        value_columns.append(values)

        metadata_rows.append(
            {
                "excel_column_number": col_idx + 1,
                "series_name": series_name,
                "field": clean_text(raw.iat[FIELD_ROW, col_idx]),
                "ticker": clean_text(raw.iat[TICKER_ROW, col_idx]),
                "region": clean_text(raw.iat[REGION_ROW, col_idx]),
                "name": clean_text(raw.iat[NAME_ROW, col_idx]),
                "type": clean_text(raw.iat[TYPE_ROW, col_idx]),
            }
        )

    unique_names = make_unique(original_names)

    data = pd.DataFrame({"Datum": dates})

    for unique_name, values in zip(unique_names, value_columns):
        data[unique_name] = values

    for row, unique_name in zip(metadata_rows, unique_names):
        row["dashboard_name"] = unique_name

    metadata = pd.DataFrame(metadata_rows)

    data = data.dropna(subset=["Datum"]).copy()
    data = data.sort_values("Datum")
    data = data.drop_duplicates(subset=["Datum"], keep="last")
    data = data.reset_index(drop=True)

    return data, metadata


def load_all_data(base_dir: Path | None = None) -> Dict[str, object]:
    excel_path = find_data_file(base_dir)

    risk_data, risk_metadata = load_timeseries_sheet(excel_path, RISK_SHEET)
    index_data, index_metadata = load_timeseries_sheet(excel_path, INDEX_SHEET)

    return {
        "excel_path": excel_path,
        "risk_measures": risk_data,
        "risk_metadata": risk_metadata,
        "global_indices": index_data,
        "index_metadata": index_metadata,
    }



# ============================================================
# END: data_loader.py
# ============================================================

# ============================================================
# BEGIN: risk_model.py
# ============================================================


from math import erf, sqrt

import numpy as np
import pandas as pd


# ============================================================
# risk_model.py
#
# Enthält nur die Kernlogik:
# - Rolling Z-Scores
# - gewichteter Durchschnitt der Z-Scores
# - Umrechnung in Risiko-Prozentwert
# ============================================================


def get_series_columns(df: pd.DataFrame) -> list[str]:
    return [col for col in df.columns if col != "Datum"]


def prepare_zscore_input(df: pd.DataFrame, cols: list[str], method: str) -> pd.DataFrame:
    out = df[["Datum"] + cols].copy()

    if method == "Tagesveränderung":
        out[cols] = out[cols].diff()
    elif method == "Prozentuale Veränderung":
        out[cols] = out[cols].pct_change()

    return out


def calculate_rolling_zscores(
    df: pd.DataFrame,
    cols: list[str],
    window: int,
    min_periods: int,
    method: str,
    clip_value: float,
    direction_map: dict[str, int],
) -> pd.DataFrame:
    x = prepare_zscore_input(df, cols, method)

    zscores = pd.DataFrame({"Datum": x["Datum"]})

    for col in cols:
        rolling_mean = x[col].rolling(window=window, min_periods=min_periods).mean()
        rolling_std = x[col].rolling(window=window, min_periods=min_periods).std(ddof=0)

        z = (x[col] - rolling_mean) / rolling_std
        z = z.replace([np.inf, -np.inf], np.nan)

        # Richtung:
        # +1 = höherer Wert bedeutet höheres Risiko
        # -1 = höherer Wert bedeutet niedrigeres Risiko
        z = z * direction_map.get(col, 1)

        if clip_value is not None and clip_value > 0:
            z = z.clip(-clip_value, clip_value)

        zscores[col] = z

    return zscores


def normalize_weights(weights: dict[str, float]) -> dict[str, float]:
    cleaned = {key: max(float(value), 0.0) for key, value in weights.items()}
    total = sum(cleaned.values())

    if total == 0:
        n = len(cleaned)
        return {key: 1 / n for key in cleaned} if n > 0 else {}

    return {key: value / total for key, value in cleaned.items()}


def calculate_weighted_average_zscore(
    zscores: pd.DataFrame,
    cols: list[str],
    weights: dict[str, float],
    output_name: str = "Composite Risk Z-Score"
) -> pd.DataFrame:
    weights_norm = normalize_weights(weights)

    out = zscores[["Datum"]].copy()

    weighted_sum = pd.Series(0.0, index=zscores.index)
    weight_sum = pd.Series(0.0, index=zscores.index)

    for col in cols:
        if col not in zscores.columns:
            continue

        w = weights_norm.get(col, 0.0)
        valid = zscores[col].notna()

        weighted_sum.loc[valid] += zscores.loc[valid, col] * w
        weight_sum.loc[valid] += w

    out[output_name] = weighted_sum / weight_sum.replace(0, np.nan)

    return out


def zscore_to_linear_percent(z: pd.Series, clip_value: float) -> pd.Series:
    """
    Lineare Umrechnung:
    -clip_value -> 0 %
     0          -> 50 %
    +clip_value -> 100 %
    """
    if clip_value is None or clip_value <= 0:
        clip_value = 3.0

    return ((z.clip(-clip_value, clip_value) + clip_value) / (2 * clip_value)) * 100


def zscore_to_normal_probability_percent(z: pd.Series) -> pd.Series:
    """
    Normal-CDF-Umrechnung:
    Gibt an, wie extrem der aktuelle Z-Score relativ zu einer Standardnormalverteilung ist.

    Beispiele:
    z = 0    -> ca. 50,0 %
    z = 1    -> ca. 84,1 %
    z = 2    -> ca. 97,7 %
    z = 3    -> ca. 99,9 %

    Wichtig:
    Das ist keine echte Crash-Wahrscheinlichkeit, sondern ein statistischer Stress-Perzentilwert.
    """
    def cdf(value):
        if pd.isna(value):
            return np.nan
        return 0.5 * (1.0 + erf(float(value) / sqrt(2.0))) * 100.0

    return z.map(cdf)


def add_risk_indicator_percent(
    composite: pd.DataFrame,
    score_method: str,
    clip_value: float,
    z_col: str = "Composite Risk Z-Score",
    output_col: str = "Risk Indicator %"
) -> pd.DataFrame:
    out = composite.copy()

    if score_method == "Normal-CDF Wahrscheinlichkeit":
        out[output_col] = zscore_to_normal_probability_percent(out[z_col])
    else:
        out[output_col] = zscore_to_linear_percent(out[z_col], clip_value)

    return out


def normalize_to_100(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df[["Datum"] + cols].copy()

    for col in cols:
        s = out[col].dropna()
        if s.empty:
            continue

        first_value = s.iloc[0]
        if pd.isna(first_value) or first_value == 0:
            continue

        out[col] = out[col] / first_value * 100

    return out


# ============================================================
# END: risk_model.py
# ============================================================

# ============================================================
# BEGIN: dashboard.py
# ============================================================


from pathlib import Path
from datetime import datetime

import altair as alt
import pandas as pd
import streamlit as st



# ============================================================
# dashboard.py
#
# Vereinfachtes Streamlit-Dashboard:
# - Risk Indicator Tab:
#   1) alle Risk-Measure-Z-Scores in einem Diagramm
#   2) gewichteter Durchschnitt + Risk Indicator % in einem Diagramm
# - Strategie:
#   Risk Indicator % vs. Global Indices mit Long-/Cash-Phasen
# - keine separaten Tabs für Risk Measures und Global Indices
#
# Start:
#   python -m streamlit run dashboard.py
# ============================================================


st.set_page_config(
    page_title="MM Risikoindikator Dashboard",
    page_icon="📊",
    layout="wide"
)


# ------------------------------------------------------------
# Daten laden mit Cache
# Cache wird automatisch erneuert, wenn sich Data.xlsx ändert.
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_data_cached(base_dir_str: str, file_mtime: float):
    return load_all_data(Path(base_dir_str))


def load_project_data():
    base_dir = Path(__file__).resolve().parent
    excel_path = find_data_file(base_dir)
    file_mtime = excel_path.stat().st_mtime
    return load_data_cached(str(base_dir), file_mtime)


# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------
def format_number(value, decimals: int = 2) -> str:
    if pd.isna(value):
        return "n/a"
    return f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def last_valid(series: pd.Series):
    s = series.dropna()
    if s.empty:
        return pd.NA
    return s.iloc[-1]


def filter_date_range(df: pd.DataFrame, start_date, end_date) -> pd.DataFrame:
    mask = (
        (df["Datum"].dt.date >= start_date)
        & (df["Datum"].dt.date <= end_date)
    )
    return df.loc[mask].copy()


def make_line_chart(df: pd.DataFrame, cols: list[str], y_title: str, height: int = 370):
    long_df = df[["Datum"] + cols].melt(
        id_vars="Datum",
        var_name="Zeitreihe",
        value_name="Wert"
    ).dropna()

    if long_df.empty:
        st.info("Keine Daten für das Diagramm vorhanden.")
        return

    chart = (
        alt.Chart(long_df)
        .mark_line()
        .encode(
            x=alt.X("Datum:T", title="Datum"),
            y=alt.Y("Wert:Q", title=y_title),
            color=alt.Color("Zeitreihe:N", title="Zeitreihe"),
            tooltip=[
                alt.Tooltip("Datum:T", title="Datum"),
                alt.Tooltip("Zeitreihe:N", title="Zeitreihe"),
                alt.Tooltip("Wert:Q", title="Wert", format=".2f"),
            ],
        )
        .properties(height=height)
        .interactive()
    )

    st.altair_chart(chart, use_container_width=True)


def make_composite_dual_axis_chart(df: pd.DataFrame):
    chart_df = df[["Datum", "Composite Risk Z-Score", "Risk Indicator %"]].dropna()

    if chart_df.empty:
        st.info("Keine Daten für das Composite-Diagramm vorhanden.")
        return

    z_chart = (
        alt.Chart(chart_df)
        .mark_line()
        .encode(
            x=alt.X("Datum:T", title="Datum"),
            y=alt.Y("Composite Risk Z-Score:Q", title="Gewichteter Durchschnitt der Z-Scores"),
            tooltip=[
                alt.Tooltip("Datum:T", title="Datum"),
                alt.Tooltip("Composite Risk Z-Score:Q", title="Composite Z", format=".2f"),
            ],
        )
    )

    percent_chart = (
        alt.Chart(chart_df)
        .mark_line(strokeDash=[6, 4])
        .encode(
            x=alt.X("Datum:T", title="Datum"),
            y=alt.Y("Risk Indicator %:Q", title="Risk Indicator %"),
            tooltip=[
                alt.Tooltip("Datum:T", title="Datum"),
                alt.Tooltip("Risk Indicator %:Q", title="Risk Indicator %", format=".1f"),
            ],
        )
    )

    chart = (
        alt.layer(z_chart, percent_chart)
        .resolve_scale(y="independent")
        .properties(height=370)
        .interactive()
    )

    st.altair_chart(chart, use_container_width=True)



def create_strategy_segments(df: pd.DataFrame, threshold: float) -> pd.DataFrame:
    """
    Erzeugt Zeitabschnitte für Hintergrundfärbung:
    Long, wenn Risk Indicator % <= threshold.
    Nicht investiert, wenn Risk Indicator % > threshold.
    """
    signal_df = df[["Datum", "Risk Indicator %"]].dropna().copy()

    if signal_df.empty:
        return pd.DataFrame(columns=["start", "end", "Status"])

    signal_df["Invested"] = signal_df["Risk Indicator %"] <= threshold
    signal_df["Group"] = (signal_df["Invested"] != signal_df["Invested"].shift()).cumsum()

    segments = (
        signal_df
        .groupby("Group", as_index=False)
        .agg(
            start=("Datum", "first"),
            end=("Datum", "last"),
            invested=("Invested", "first"),
        )
    )

    # Ende etwas nach rechts verlängern, damit der letzte Bereich sichtbar ist
    if len(signal_df) >= 2:
        last_step = signal_df["Datum"].diff().dropna().median()
        if pd.isna(last_step):
            last_step = pd.Timedelta(days=1)
    else:
        last_step = pd.Timedelta(days=1)

    segments["end"] = segments["end"] + last_step
    segments["Status"] = segments["invested"].map({
        True: "Long",
        False: "Nicht investiert"
    })

    return segments[["start", "end", "Status"]]





def make_strategy_chart(comparison: pd.DataFrame, selected_indices: list[str], threshold: float):
    """
    Strategie-Chart:
    - Linke Achse: Global Indices, eindeutig auf 100 indexiert
    - Rechte Achse: Risk Indicator %, unverändert 0 bis 100
    - Hintergrund: Long / Nicht investiert
    - Legende: nur die einzelnen Global-Index-Zeitreihen
    """
    if comparison.empty:
        st.info("Keine gemeinsamen Daten für den Strategie-Vergleich vorhanden.")
        return

    # --------------------------------------------------------
    # 1. Global Indices wirklich auf 100 indexieren
    #    Basis = erster verfügbarer Wert je Index im gewählten Zeitraum
    # --------------------------------------------------------
    index_frames = []
    base_rows = []

    for col in selected_indices:
        temp = comparison[["Datum", col]].dropna().copy()

        if temp.empty:
            continue

        base_value = temp[col].iloc[0]
        base_date = temp["Datum"].iloc[0]

        if pd.isna(base_value) or base_value == 0:
            continue

        temp["Indexiert auf 100"] = temp[col] / base_value * 100
        temp["Global Index"] = col

        index_frames.append(temp[["Datum", "Global Index", "Indexiert auf 100"]])

        base_rows.append({
            "Global Index": col,
            "Basisdatum": base_date,
            "Basiswert": base_value,
        })

    if not index_frames:
        st.info("Für die ausgewählten Global Indices gibt es im Zeitraum keine verwertbaren Daten.")
        return

    index_long = pd.concat(index_frames, ignore_index=True)
    base_table = pd.DataFrame(base_rows)

    risk_part = comparison[["Datum", "Risk Indicator %"]].dropna().copy()

    with st.expander("Indexierungsbasis anzeigen"):
        st.caption("Jeder Global Index wird mit seinem ersten verfügbaren Wert im gewählten Zeitraum auf 100 gesetzt.")
        st.dataframe(base_table, use_container_width=True)

    # --------------------------------------------------------
    # 2. Hintergrundphasen getrennt zeichnen, damit sie die
    #    Legende der Global-Index-Linien nicht überschreiben.
    # --------------------------------------------------------
    segments = create_strategy_segments(comparison, threshold)

    layers = []

    if not segments.empty:
        long_segments = segments[segments["Status"] == "Long"]
        cash_segments = segments[segments["Status"] == "Nicht investiert"]

        if not long_segments.empty:
            background_long = (
                alt.Chart(long_segments)
                .mark_rect(opacity=0.10, color="#2ca02c")
                .encode(
                    x=alt.X("start:T", title="Datum"),
                    x2="end:T",
                    tooltip=[
                        alt.Tooltip("start:T", title="Start"),
                        alt.Tooltip("end:T", title="Ende"),
                        alt.Tooltip("Status:N", title="Status"),
                    ],
                )
            )
            layers.append(background_long)

        if not cash_segments.empty:
            background_cash = (
                alt.Chart(cash_segments)
                .mark_rect(opacity=0.10, color="#d62728")
                .encode(
                    x=alt.X("start:T", title="Datum"),
                    x2="end:T",
                    tooltip=[
                        alt.Tooltip("start:T", title="Start"),
                        alt.Tooltip("end:T", title="Ende"),
                        alt.Tooltip("Status:N", title="Status"),
                    ],
                )
            )
            layers.append(background_cash)

    # --------------------------------------------------------
    # 3. Linke Achse: Global Indices indexiert auf 100
    # --------------------------------------------------------
    index_chart = (
        alt.Chart(index_long)
        .mark_line(strokeWidth=2)
        .encode(
            x=alt.X("Datum:T", title="Datum"),
            y=alt.Y(
                "Indexiert auf 100:Q",
                title="Global Indices, indexiert auf 100",
                axis=alt.Axis(orient="left")
            ),
            color=alt.Color(
                "Global Index:N",
                title="Global Indices",
                legend=alt.Legend(orient="bottom")
            ),
            tooltip=[
                alt.Tooltip("Datum:T", title="Datum"),
                alt.Tooltip("Global Index:N", title="Global Index"),
                alt.Tooltip("Indexiert auf 100:Q", title="Indexiert auf 100", format=".2f"),
            ],
        )
    )

    layers.append(index_chart)

    # --------------------------------------------------------
    # 4. Rechte Achse: Risk Indicator %, nicht indexiert
    # --------------------------------------------------------
    if not risk_part.empty:
        risk_chart = (
            alt.Chart(risk_part)
            .mark_line(strokeWidth=3, strokeDash=[6, 4], color="black")
            .encode(
                x=alt.X("Datum:T", title="Datum"),
                y=alt.Y(
                    "Risk Indicator %:Q",
                    title="Risk Indicator %",
                    scale=alt.Scale(domain=[0, 100]),
                    axis=alt.Axis(orient="right")
                ),
                tooltip=[
                    alt.Tooltip("Datum:T", title="Datum"),
                    alt.Tooltip("Risk Indicator %:Q", title="Risk Indicator %", format=".1f"),
                ],
            )
        )

        threshold_df = pd.DataFrame({
            "Datum": [index_long["Datum"].min(), index_long["Datum"].max()],
            "Grenze": [threshold, threshold]
        })

        threshold_line = (
            alt.Chart(threshold_df)
            .mark_line(strokeDash=[2, 2], color="gray")
            .encode(
                x=alt.X("Datum:T", title="Datum"),
                y=alt.Y(
                    "Grenze:Q",
                    title="Risk Indicator %",
                    scale=alt.Scale(domain=[0, 100]),
                    axis=alt.Axis(orient="right")
                ),
                tooltip=[alt.Tooltip("Grenze:Q", title="Risk-Indicator-Grenze", format=".1f")]
            )
        )

        layers.extend([risk_chart, threshold_line])
    else:
        st.warning(
            "Für den gewählten Zeitraum gibt es noch keinen Risk Indicator %. "
            "Die Global Indices werden trotzdem indexiert angezeigt. Prüfe ggf. Rolling Window / Min. Datenpunkte."
        )

    chart = (
        alt.layer(*layers)
        .resolve_scale(y="independent")
        .properties(height=480)
        .interactive()
    )

    st.altair_chart(chart, use_container_width=True)



def calculate_strategy_excess_return(
    comparison: pd.DataFrame,
    selected_indices: list[str],
    threshold: float,
    use_previous_signal: bool = True,
) -> pd.DataFrame:
    """
    Berechnet Strategie vs. Buy-and-Hold je Global Index.

    Wichtig:
    Die Berechnung startet erst ab dem ersten gültigen Risk Indicator.
    Vorher gibt es wegen Rolling Window / Min. Datenpunkte noch kein echtes Signal.
    Daher werden Strategie, Buy-and-Hold und Excess Return davor auf NA gesetzt.
    """
    if comparison.empty:
        return pd.DataFrame(columns=["Datum"])

    out = pd.DataFrame({"Datum": comparison["Datum"]})

    valid_risk = comparison["Risk Indicator %"].notna()

    if not valid_risk.any():
        out["Signal verwendet"] = pd.NA
        for col in selected_indices:
            out[f"{col} Buy-and-Hold"] = pd.NA
            out[f"{col} Strategie"] = pd.NA
            out[f"{col} Excess Return %"] = pd.NA
        return out

    first_valid_risk_idx = valid_risk[valid_risk].index[0]

    raw_signal = comparison["Risk Indicator %"] <= threshold
    raw_signal = raw_signal.where(valid_risk, pd.NA)

    if use_previous_signal:
        invested = raw_signal.shift(1)
    else:
        invested = raw_signal

    # Vor dem ersten verwendbaren Signal gibt es keine Strategie.
    invested.loc[invested.index < first_valid_risk_idx] = pd.NA

    out["Signal verwendet"] = invested.map({
        True: "Long",
        False: "Nicht investiert"
    })

    for col in selected_indices:
        if col not in comparison.columns:
            continue

        prices = comparison[col].astype(float)

        if prices.dropna().empty:
            continue

        # Ab dem ersten gültigen Risk Indicator neu starten.
        prices_valid = prices.copy()
        prices_valid.loc[prices_valid.index < first_valid_risk_idx] = pd.NA

        returns = prices_valid.pct_change(fill_method=None)

        benchmark_returns = returns.copy()
        strategy_returns = returns.where(invested == True, 0.0)

        benchmark_returns.loc[benchmark_returns.index < first_valid_risk_idx] = pd.NA
        strategy_returns.loc[strategy_returns.index < first_valid_risk_idx] = pd.NA

        benchmark_returns = benchmark_returns.fillna(0.0)
        strategy_returns = strategy_returns.fillna(0.0)

        benchmark_index = (1.0 + benchmark_returns).cumprod() * 100.0
        strategy_index = (1.0 + strategy_returns).cumprod() * 100.0

        # Vor dem ersten gültigen Risk Indicator keine Werte anzeigen.
        benchmark_index.loc[benchmark_index.index < first_valid_risk_idx] = pd.NA
        strategy_index.loc[strategy_index.index < first_valid_risk_idx] = pd.NA

        # Am Startdatum beide auf 100 setzen.
        benchmark_index.loc[first_valid_risk_idx] = 100.0
        strategy_index.loc[first_valid_risk_idx] = 100.0

        out[f"{col} Buy-and-Hold"] = benchmark_index
        out[f"{col} Strategie"] = strategy_index
        out[f"{col} Excess Return %"] = (strategy_index / benchmark_index - 1.0) * 100.0

    return out

def make_strategy_excess_return_chart(strategy_df: pd.DataFrame, selected_indices: list[str]):
    """
    Chart: Excess Return der Long/Cash-Strategie gegenüber Buy-and-Hold.
    Positive Werte = Strategie besser als Buy-and-Hold.
    Negative Werte = Strategie schlechter als Buy-and-Hold.
    """
    excess_cols = [
        f"{col} Excess Return %"
        for col in selected_indices
        if f"{col} Excess Return %" in strategy_df.columns
    ]

    if not excess_cols:
        st.info("Keine Excess-Return-Zeitreihen vorhanden.")
        return

    long_df = strategy_df[["Datum"] + excess_cols].melt(
        id_vars="Datum",
        var_name="Zeitreihe",
        value_name="Excess Return %"
    ).dropna()

    if long_df.empty:
        st.info("Keine Werte für das Excess-Return-Diagramm vorhanden.")
        return

    # Namen lesbarer machen
    long_df["Zeitreihe"] = long_df["Zeitreihe"].str.replace(" Excess Return %", "", regex=False)

    zero_line = pd.DataFrame({
        "Datum": [long_df["Datum"].min(), long_df["Datum"].max()],
        "Null": [0.0, 0.0]
    })

    excess_chart = (
        alt.Chart(long_df)
        .mark_line(strokeWidth=2.5)
        .encode(
            x=alt.X("Datum:T", title="Datum"),
            y=alt.Y("Excess Return %:Q", title="Excess Return vs. Buy-and-Hold (%)"),
            color=alt.Color(
                "Zeitreihe:N",
                title="Global Index",
                legend=alt.Legend(orient="bottom")
            ),
            tooltip=[
                alt.Tooltip("Datum:T", title="Datum"),
                alt.Tooltip("Zeitreihe:N", title="Global Index"),
                alt.Tooltip("Excess Return %:Q", title="Excess Return %", format=".2f"),
            ],
        )
    )

    zero = (
        alt.Chart(zero_line)
        .mark_line(strokeDash=[2, 2], color="gray")
        .encode(
            x=alt.X("Datum:T", title="Datum"),
            y=alt.Y("Null:Q", title="Excess Return vs. Buy-and-Hold (%)"),
        )
    )

    chart = (
        alt.layer(excess_chart, zero)
        .properties(height=380)
        .interactive()
    )

    st.altair_chart(chart, use_container_width=True)



def calculate_risk_indicator(
    risk_df: pd.DataFrame,
    selected_risk: list[str],
    z_method: str,
    rolling_window: int,
    min_periods: int,
    clip_value: float,
    score_method: str,
    direction_map: dict[str, int],
    weights_raw: dict[str, float],
):
    zscores = calculate_rolling_zscores(
        df=risk_df,
        cols=selected_risk,
        window=int(rolling_window),
        min_periods=int(min_periods),
        method=z_method,
        clip_value=float(clip_value),
        direction_map=direction_map,
    )

    composite = calculate_weighted_average_zscore(
        zscores=zscores,
        cols=selected_risk,
        weights=weights_raw,
        output_name="Composite Risk Z-Score",
    )

    composite = add_risk_indicator_percent(
        composite=composite,
        score_method=score_method,
        clip_value=float(clip_value),
        z_col="Composite Risk Z-Score",
        output_col="Risk Indicator %",
    )

    return zscores, composite


# ------------------------------------------------------------
# Header
# ------------------------------------------------------------
st.title("MM Risikoindikator Dashboard")
st.caption("Vereinfachte Version: Rolling Z-Scores, gewichteter Risk Indicator und Vergleich mit Global Indices")


# ------------------------------------------------------------
# Daten laden
# ------------------------------------------------------------
try:
    loaded = load_project_data()
except Exception as error:
    st.error("Daten konnten nicht geladen werden.")
    st.write(error)
    st.info(
        "Bitte prüfe, ob Data.xlsx im gleichen Ordner liegt wie dashboard.py "
        "und ob die Sheets 'data (risk measures)' und 'data (global indices)' existieren."
    )
    st.stop()

risk_df = loaded["risk_measures"]
risk_meta = loaded["risk_metadata"]
index_df = loaded["global_indices"]
index_meta = loaded["index_metadata"]

risk_cols = get_series_columns(risk_df)
index_cols = get_series_columns(index_df)


# ------------------------------------------------------------
# Sidebar: globale Einstellungen
# ------------------------------------------------------------
st.sidebar.title("Einstellungen")

auto_refresh = st.sidebar.toggle("Auto-Refresh aktivieren", value=False)

refresh_seconds = st.sidebar.number_input(
    "Refresh-Intervall in Sekunden",
    min_value=5,
    max_value=600,
    value=30,
    step=5
)

if auto_refresh:
    st.sidebar.info("Auto-Refresh ist aktiv. Bitte Browser-Ansicht geöffnet lassen.")
    st_autorefresh_placeholder = st.empty()
    # Einfacher Refresh über meta refresh.
    st.markdown(
        f"<meta http-equiv='refresh' content='{int(refresh_seconds)}'>",
        unsafe_allow_html=True
    )

st.sidebar.markdown("### Risk-Indicator-Setup")

selected_risk = st.sidebar.multiselect(
    "Risk Measures auswählen",
    options=risk_cols,
    default=risk_cols[:min(4, len(risk_cols))],
)

z_method = st.sidebar.selectbox(
    "Z-Score-Basis",
    options=["Level", "Tagesveränderung", "Prozentuale Veränderung"],
    index=0,
)

rolling_window = st.sidebar.number_input(
    "Rolling Window",
    min_value=20,
    max_value=3000,
    value=756,
    step=21,
)

min_periods = st.sidebar.number_input(
    "Min. Datenpunkte",
    min_value=20,
    max_value=int(rolling_window),
    value=min(252, int(rolling_window)),
    step=21,
)

clip_value = st.sidebar.number_input(
    "Z-Score begrenzen auf ±",
    min_value=0.5,
    max_value=10.0,
    value=3.0,
    step=0.5,
)

score_method = st.sidebar.selectbox(
    "Umrechnung in Prozentwert",
    options=["Linear 0-100", "Normal-CDF Wahrscheinlichkeit"],
    index=0,
    help=(
        "Linear 0-100: -clip=0%, 0=50%, +clip=100%. "
        "Normal-CDF: statistischer Perzentilwert einer Standardnormalverteilung, keine echte Crash-Wahrscheinlichkeit."
    ),
)

with st.sidebar.expander("Richtung invertieren"):
    direction_map = {}
    for col in selected_risk:
        invert = st.checkbox(
            col,
            value=False,
            key=f"invert_{col}",
            help="Aktivieren, falls ein höherer Wert dieser Zeitreihe niedrigeres Risiko bedeutet.",
        )
        direction_map[col] = -1 if invert else 1

with st.sidebar.expander("Gewichte"):
    weights_raw = {}
    for col in selected_risk:
        weights_raw[col] = st.number_input(
            col,
            min_value=0.0,
            max_value=100.0,
            value=1.0,
            step=0.25,
            key=f"weight_{col}",
        )

if not selected_risk:
    st.warning("Bitte mindestens einen Risikoindikator in der Sidebar auswählen.")
    st.stop()

weights_norm = normalize_weights(weights_raw)

# Gemeinsamer Datumsbereich
min_date = min(risk_df["Datum"].min(), index_df["Datum"].min()).date()
max_date = max(risk_df["Datum"].max(), index_df["Datum"].max()).date()

date_range = st.sidebar.date_input(
    "Datumsbereich",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date


# ------------------------------------------------------------
# Risk Indicator berechnen
# ------------------------------------------------------------
zscores, composite = calculate_risk_indicator(
    risk_df=risk_df,
    selected_risk=selected_risk,
    z_method=z_method,
    rolling_window=int(rolling_window),
    min_periods=int(min_periods),
    clip_value=float(clip_value),
    score_method=score_method,
    direction_map=direction_map,
    weights_raw=weights_raw,
)

zscores_view = filter_date_range(zscores, start_date, end_date)
composite_view = filter_date_range(composite, start_date, end_date)


# ------------------------------------------------------------
# Tabs
# ------------------------------------------------------------
tab_risk_indicator, tab_strategy, tab_check = st.tabs(
    ["Risk Indicator", "Strategie", "Datencheck"]
)


with tab_risk_indicator:
    st.write(f"Geladene Datei: `{loaded['excel_path'].name}` · Letztes Laden: {datetime.now().strftime('%H:%M:%S')}")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric(
            "Composite Risk Z-Score",
            format_number(last_valid(composite_view["Composite Risk Z-Score"])),
        )

    with c2:
        st.metric(
            "Risk Indicator %",
            format_number(last_valid(composite_view["Risk Indicator %"]), decimals=1),
        )

    with c3:
        st.metric(
            "Anzahl Risk Measures",
            len(selected_risk),
        )

    st.markdown("### 1. Rolling Z-Scores der Risk Measures")
    make_line_chart(
        df=zscores_view,
        cols=selected_risk,
        y_title="Rolling Z-Score",
        height=380,
    )

    st.markdown("### 2. Gewichteter Durchschnitt und Risk Indicator %")
    make_composite_dual_axis_chart(composite_view)

    with st.expander("Gewichte und Daten anzeigen"):
        weight_table = pd.DataFrame({
            "Risk Measure": list(weights_norm.keys()),
            "Gewicht": [weights_norm[col] for col in weights_norm],
        })
        weight_table["Gewicht"] = weight_table["Gewicht"].map(lambda x: f"{x:.1%}")

        st.markdown("#### Normalisierte Gewichte")
        st.dataframe(weight_table, use_container_width=True)

        st.markdown("#### Z-Score-Daten")
        st.dataframe(zscores_view[["Datum"] + selected_risk], use_container_width=True)

        st.markdown("#### Composite-Daten")
        st.dataframe(composite_view, use_container_width=True)



with tab_strategy:
    st.subheader("Strategie: Risk Indicator als Timing-Signal")

    st.caption(
        "Regel: Long, wenn der Risk Indicator % kleiner oder gleich deiner Grenze ist. "
        "Nicht investiert, wenn der Risk Indicator % über der Grenze liegt."
    )

    selected_indices = st.multiselect(
        "Global Indices auswählen",
        options=index_cols,
        default=index_cols[:min(3, len(index_cols))],
    )

    risk_threshold = st.number_input(
        "Risk-Indicator-Grenze für Long-Signal",
        min_value=0.0,
        max_value=100.0,
        value=60.0,
        step=1.0,
        help="Long, wenn Risk Indicator % <= Grenze. Nicht investiert, wenn Risk Indicator % > Grenze.",
    )

    if not selected_indices:
        st.info("Bitte mindestens einen Global Index auswählen.")
    else:
        comparison = pd.merge(
            composite[["Datum", "Composite Risk Z-Score", "Risk Indicator %"]],
            index_df[["Datum"] + selected_indices],
            on="Datum",
            how="inner",
        ).sort_values("Datum")

        comparison_view = filter_date_range(comparison, start_date, end_date)

        st.write(
            f"Gemeinsame Datenpunkte im Zeitraum: {len(comparison_view)} · "
            f"Index-Zeitreihen: {len(selected_indices)}"
        )

        signal = comparison_view["Risk Indicator %"] <= risk_threshold
        invested_share = signal.mean() if len(signal.dropna()) > 0 else pd.NA
        latest_risk = last_valid(comparison_view["Risk Indicator %"])
        latest_signal = "Long" if pd.notna(latest_risk) and latest_risk <= risk_threshold else "Nicht investiert"

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("Aktuelles Signal", latest_signal)

        with c2:
            st.metric("Aktueller Risk Indicator %", format_number(latest_risk, decimals=1))

        with c3:
            if pd.isna(invested_share):
                st.metric("Long-Anteil im Zeitraum", "n/a")
            else:
                st.metric("Long-Anteil im Zeitraum", f"{invested_share:.1%}")

        st.markdown("### Risk Indicator vs. Global Indices mit Long-/Cash-Phasen")

        make_strategy_chart(
            comparison=comparison_view,
            selected_indices=selected_indices,
            threshold=float(risk_threshold),
        )

        st.markdown("### Excess Return der Strategie gegenüber Buy-and-Hold")
        st.caption(
            "Positive Werte bedeuten: Die Long/Cash-Strategie liegt über der normalen Buy-and-Hold-Zeitreihe. "
            "Negative Werte bedeuten: Die Strategie liegt darunter. "
            "In Nicht-investiert-Phasen wird keine Indexrendite mitgenommen. Die Berechnung startet erst ab dem ersten gültigen Risk Indicator nach Min. Datenpunkten."
        )

        use_previous_signal = st.toggle(
            "Signal vom Vortag verwenden",
            value=True,
            help=(
                "Aktiviert = realistischer ohne Look-Ahead: Das Signal von gestern steuert die Rendite von gestern auf heute. "
                "Deaktiviert = Signal desselben Tages."
            ),
        )

        strategy_df = calculate_strategy_excess_return(
            comparison=comparison_view,
            selected_indices=selected_indices,
            threshold=float(risk_threshold),
            use_previous_signal=use_previous_signal,
        )

        make_strategy_excess_return_chart(
            strategy_df=strategy_df,
            selected_indices=selected_indices,
        )

        with st.expander("Strategiedaten anzeigen"):
            export_df = comparison_view.copy()
            export_df["Signal"] = export_df["Risk Indicator %"].apply(
                lambda x: "Long" if pd.notna(x) and x <= risk_threshold else "Nicht investiert"
            )

            indexed_indices = normalize_to_100(comparison_view, selected_indices)
            indexed_indices = indexed_indices.rename(
                columns={col: f"{col} indexiert" for col in selected_indices}
            )

            export_df = pd.merge(
                export_df[["Datum", "Composite Risk Z-Score", "Risk Indicator %", "Signal"]],
                indexed_indices,
                on="Datum",
                how="left",
            )

            strategy_cols = [
                col for col in strategy_df.columns
                if col != "Datum"
            ]

            if strategy_cols:
                export_df = pd.merge(
                    export_df,
                    strategy_df[["Datum"] + strategy_cols],
                    on="Datum",
                    how="left"
                )

            st.dataframe(export_df, use_container_width=True)


with tab_check:
    st.subheader("Datencheck")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### Risk Measures")
        st.write(f"Zeilen: {len(risk_df)}")
        st.write(f"Zeitreihen: {len(risk_cols)}")
        st.dataframe(risk_df.isna().sum().to_frame("Fehlende Werte"), use_container_width=True)

    with c2:
        st.markdown("#### Global Indices")
        st.write(f"Zeilen: {len(index_df)}")
        st.write(f"Zeitreihen: {len(index_cols)}")
        st.dataframe(index_df.isna().sum().to_frame("Fehlende Werte"), use_container_width=True)

    with st.expander("Risk-Metadaten"):
        st.dataframe(risk_meta, use_container_width=True)

    with st.expander("Index-Metadaten"):
        st.dataframe(index_meta, use_container_width=True)


# ============================================================
# END: dashboard.py
# ============================================================
