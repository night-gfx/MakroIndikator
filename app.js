let D = null;
const RISK_SHEET = 'data (risk measures)';
const INDEX_SHEET = 'data (global indices)';
const SERIES_NAME_ROW = 5;
const DATA_START_ROW = 6;
const $ = id => document.getElementById(id);

function fmt(x, d = 2) {
  return x == null || Number.isNaN(x)
    ? 'n/a'
    : x.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
}
function clean(v) { return v == null ? '' : String(v).trim(); }
function excelDateToISO(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (m) {
    const y = m[3].length === 2 ? Number(`20${m[3]}`) : Number(m[3]);
    return `${y}-${String(Number(m[2])).padStart(2, '0')}-${String(Number(m[1])).padStart(2, '0')}`;
  }
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}
function toNumber(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  let s = String(v).trim();
  if (['nan', 'NaN', 'None', '#N/A', '#N/A Field Not Applicable', '#VALUE!'].includes(s)) return null;
  if (s.includes(',')) s = s.replaceAll('.', '').replace(',', '.');
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}
function makeUnique(names) {
  const counts = {}, out = [];
  for (const name of names) {
    const base = name || 'Unnamed';
    counts[base] = (counts[base] || 0) + 1;
    out.push(counts[base] === 1 ? base : `${base}_${counts[base]}`);
  }
  return out;
}
function parseSheet(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet fehlt: ${sheetName}`);
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const originalNames = [], valueCols = [], metaRows = [];
  const maxCols = Math.max(...raw.map(r => r.length));

  for (let c = 1; c < maxCols; c++) {
    const seriesName = clean(raw[SERIES_NAME_ROW]?.[c]);
    if (!seriesName) continue;
    const values = raw.slice(DATA_START_ROW).map(r => toNumber(r?.[c]));
    if (values.every(v => v == null)) continue;
    originalNames.push(seriesName);
    valueCols.push(values);
    metaRows.push({
      excel_column_number: c + 1,
      series_name: seriesName,
      field: clean(raw[0]?.[c]),
      ticker: clean(raw[1]?.[c]),
      region: clean(raw[2]?.[c]),
      name: clean(raw[3]?.[c]),
      type: clean(raw[4]?.[c])
    });
  }

  const unique = makeUnique(originalNames);
  const rows = [];
  for (let r = DATA_START_ROW; r < raw.length; r++) {
    const iso = excelDateToISO(raw[r]?.[0]);
    if (!iso) continue;
    const row = { Datum: iso };
    unique.forEach((name, i) => row[name] = valueCols[i][r - DATA_START_ROW]);
    rows.push(row);
  }
  rows.sort((a, b) => a.Datum.localeCompare(b.Datum));
  const dedup = Array.from(new Map(rows.map(r => [r.Datum, r])).values());
  metaRows.forEach((m, i) => m.dashboard_name = unique[i]);
  return { records: dedup, columns: unique, metadata: metaRows };
}
async function loadData() {
  const res = await fetch('Data.xlsx', { cache: 'no-store' });
  if (!res.ok) throw new Error('Data.xlsx konnte nicht geladen werden. Liegt sie im gleichen Ordner wie index.html?');
  const ab = await res.arrayBuffer();
  const workbook = XLSX.read(ab, { type: 'array', cellDates: false });
  D = {
    risk: parseSheet(workbook, RISK_SHEET),
    indices: parseSheet(workbook, INDEX_SHEET)
  };
}

function erf(x) {
  const s = x >= 0 ? 1 : -1;
  x = Math.abs(x);
  const a1 = .254829592, a2 = -.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429, p = .3275911;
  const t = 1 / (1 + p * x);
  const y = 1 - (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t) * Math.exp(-x * x);
  return s * y;
}
function cdf(z) { return .5 * (1 + erf(z / Math.sqrt(2))) * 100; }
function safeId(prefix, c) { return prefix + btoa(unescape(encodeURIComponent(c))).replaceAll('=', ''); }
function getSelected(prefix, cols) { return cols.filter(c => $(safeId(prefix, c))?.checked); }
function val(id, fallback = 0) {
  const el = $(id);
  if (!el) return fallback;
  const number = Number(el.value);
  return Number.isFinite(number) ? number : fallback;
}
function series(records, col) { return records.map(r => r[col]); }
function rollZ(arr, window, minp, clip, dir) {
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    const vals = [];
    for (let j = Math.max(0, i - window + 1); j <= i; j++) {
      if (arr[j] != null && !Number.isNaN(arr[j])) vals.push(arr[j]);
    }
    if (vals.length < minp || arr[i] == null) { out.push(null); continue; }
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    const sd = Math.sqrt(vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length);
    let z = sd ? ((arr[i] - mean) / sd) * dir : null;
    if (z != null) z = Math.max(-clip, Math.min(clip, z));
    out.push(z);
  }
  return out;
}
function kalmanFilter(values) {
  // Kausales Local-Level-Kalman-Modell: nur Werte bis einschließlich heute.
  const valid = values.filter(v => v != null && Number.isFinite(v));
  if (!valid.length) return values.map(() => null);

  const mean = valid.reduce((a, b) => a + b, 0) / valid.length;
  const variance = valid.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(valid.length - 1, 1);
  const measurementNoise = Math.max(variance, 1e-6);
  // Prozessrauschen relativ zum Messrauschen: reaktionsfähiger als ein langer SMA,
  // aber weiterhin deutlich stabiler als der ungefilterte tägliche Wert.
  const processNoise = Math.max(measurementNoise * 0.03, 1e-7);

  let state = null;
  let covariance = measurementNoise;
  return values.map(observation => {
    if (observation == null || !Number.isFinite(observation)) return null;
    if (state == null) {
      state = observation;
      covariance = measurementNoise;
      return state;
    }
    const predictedState = state;
    const predictedCovariance = covariance + processNoise;
    const gain = predictedCovariance / (predictedCovariance + measurementNoise);
    state = predictedState + gain * (observation - predictedState);
    covariance = (1 - gain) * predictedCovariance;
    return state;
  });
}
function selectedRisk() { return getSelected('risk_', D.risk.columns); }
function selectedIndices() { return getSelected('idx_', D.indices.columns); }
function updateWeightTotal() {
  const selected = selectedRisk();
  const total = selected.reduce((sum, c) => {
    const input = $(safeId('w_', c));
    return sum + Math.max(0, Number(input?.value) || 0);
  }, 0);
  const label = $('weightTotal');
  label.textContent = `${fmt(total, 1)} %`;
  label.style.color = Math.abs(total - 100) < 0.01 ? '#166534' : '#92400e';
  label.style.background = Math.abs(total - 100) < 0.01 ? '#dcfce7' : '#fef3c7';
}

function normalizedWeightsFromInputs(cols) {
  const raw = {};
  const total = cols.reduce((sum, c) => {
    const input = $(safeId('w_', c));
    raw[c] = Math.max(0, Number(input?.value) || 0);
    return sum + raw[c];
  }, 0);
  const normalized = {};
  if (!cols.length) return normalized;
  if (total === 0) cols.forEach(c => normalized[c] = 1 / cols.length);
  else cols.forEach(c => normalized[c] = raw[c] / total);
  return normalized;
}
function compositeFromWeights(zscores, cols, weights) {
  const n = zscores.Datum.length;
  const comp = [], risk = [];
  for (let i = 0; i < n; i++) {
    let sum = 0, validWeight = 0;
    cols.forEach(c => {
      const z = zscores[c][i];
      if (z != null) { sum += z * weights[c]; validWeight += weights[c]; }
    });
    const composite = validWeight ? sum / validWeight : null;
    comp.push(composite);
    risk.push(composite == null ? null : cdf(composite));
  }
  const kalmanComp = kalmanFilter(comp);
  const kalmanRisk = kalmanComp.map(z => z == null ? null : cdf(z));
  return { comp, risk, kalmanComp, kalmanRisk };
}
function empiricalQuantile(values, q) {
  const valid = values.filter(v => v != null && Number.isFinite(v)).sort((a, b) => a - b);
  if (!valid.length) return null;
  const pos = (valid.length - 1) * Math.max(0, Math.min(1, q));
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  if (lo === hi) return valid[lo];
  return valid[lo] + (valid[hi] - valid[lo]) * (pos - lo);
}
function historicalRiskBands(kalmanRisk) {
  // Kausal: Alle Bänder am Tag t basieren ausschließlich auf Werten bis t-1.
  // Dadurch gibt es keine feste Risiko-Grenze und keinen Look-Ahead-Bias.
  const lookback = Math.max(20, Math.floor(val('adaptiveLookback')));
  const highQ = Math.max(0.55, Math.min(0.99, val('highRiskQuantile') / 100));
  const low = [], middle = [], high = [];
  const minHistory = Math.min(20, lookback);

  for (let i = 0; i < kalmanRisk.length; i++) {
    const history = [];
    for (let j = Math.max(0, i - lookback); j < i; j++) {
      const risk = kalmanRisk[j];
      if (risk != null && Number.isFinite(risk)) history.push(risk);
    }
    if (history.length < minHistory) {
      low.push(null); middle.push(null); high.push(null);
      continue;
    }
    // Niedrig = unteres Drittel, Mittel = Median, High = frei wählbares oberes Quantil.
    low.push(empiricalQuantile(history, 1 / 3));
    middle.push(empiricalQuantile(history, 0.50));
    high.push(empiricalQuantile(history, highQ));
  }
  return { low, middle, high, highQ };
}
function classifyRiskRegime(risk, bands) {
  if (risk == null || bands.low == null || bands.middle == null || bands.high == null) return 'n/a';
  if (risk <= bands.low) return 'Niedrig';
  if (risk <= bands.middle) return 'Mäßig';
  if (risk <= bands.high) return 'Erhöht';
  return 'Hoch';
}
function findMsciWorldColumn() {
  return D.indices.columns.find(c => clean(c).toUpperCase() === 'MSCI WORLD')
    || D.indices.columns.find(c => clean(c).toUpperCase().includes('MSCI WORLD'))
    || null;
}
function mean(values) { return values.reduce((a, b) => a + b, 0) / values.length; }
function sampleStd(values) {
  if (values.length < 2) return null;
  const m = mean(values);
  return Math.sqrt(values.reduce((a, b) => a + (b - m) ** 2, 0) / (values.length - 1));
}
function annualizedSharpe(strategyReturns) {
  const usable = strategyReturns.filter(x => x != null && Number.isFinite(x));
  if (usable.length < 60) return -Infinity;
  const sd = sampleStd(usable);
  if (!sd || !Number.isFinite(sd)) return -Infinity;
  return mean(usable) / sd * Math.sqrt(252);
}
function strategyReturnsForMsci(dates, riskSeries, trainStart, trainEnd) {
  const msci = findMsciWorldColumn();
  if (!msci) return null;
  const priceByDate = new Map(D.indices.records.map(r => [r.Datum, r[msci]]));
  const usePrevious = $('prevSignal').checked;
  const prices = dates.map(d => priceByDate.get(d) ?? null);
  const thresholds = historicalRiskBands(riskSeries).high;
  const signals = riskSeries.map((r, i) => r == null ? null : r <= thresholds[i]);
  const returns = [];

  for (let i = 1; i < dates.length; i++) {
    if (dates[i] < trainStart || dates[i] > trainEnd) continue;
    if (prices[i] == null || prices[i - 1] == null || prices[i - 1] === 0) continue;
    const decision = usePrevious ? signals[i - 1] : signals[i];
    if (decision == null) continue;
    const dailyReturn = prices[i] / prices[i - 1] - 1;
    returns.push(decision ? dailyReturn : 0);
  }
  return returns;
}
function seededRandom(seed) {
  let t = seed >>> 0;
  return () => {
    t += 0x6D2B79F5;
    let r = Math.imul(t ^ (t >>> 15), 1 | t);
    r ^= r + Math.imul(r ^ (r >>> 7), 61 | r);
    return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
  };
}
function randomWeightsWithCap(n, maxWeight, random) {
  if (n === 0) return [];
  const cap = Math.max(1 / n, Math.min(1, maxWeight));
  for (let attempt = 0; attempt < 250; attempt++) {
    const v = Array.from({ length: n }, () => -Math.log(Math.max(random(), 1e-12)));
    const total = v.reduce((a, b) => a + b, 0);
    const weights = v.map(x => x / total);
    if (Math.max(...weights) <= cap + 1e-12) return weights;
  }
  return Array(n).fill(1 / n);
}
function scoreCandidateWeights(zscores, cols, candidate, dates, trainStart, trainEnd, useKalman) {
  const weights = {};
  cols.forEach((c, i) => weights[c] = candidate[i]);
  const result = compositeFromWeights(zscores, cols, weights);
  const riskSeries = useKalman ? result.kalmanRisk : result.risk;
  const returns = strategyReturnsForMsci(dates, riskSeries, trainStart, trainEnd);
  if (!returns) return { score: -Infinity, result: null };
  return { score: annualizedSharpe(returns), result };
}
function optimizeWeightsForMsci() {
  const cols = selectedRisk();
  const status = $('optimizationStatus');
  const btn = $('optimizeBtn');
  const msci = findMsciWorldColumn();
  const trainStart = $('startDate')?.value || '';
  const trainEnd = $('optEndDate')?.value || $('endDate')?.value || '';
  const maxWeightPct = val('maxWeight');
  const maxWeight = maxWeightPct / 100;

  if (!msci) {
    status.textContent = 'MSCI WORLD wurde im Sheet „data (global indices)“ nicht gefunden.';
    status.className = 'optimizerStatus warning';
    return;
  }
  if (!cols.length) {
    status.textContent = 'Bitte mindestens einen Risk Measure auswählen.';
    status.className = 'optimizerStatus warning';
    return;
  }
  if (maxWeight + 1e-12 < 1 / cols.length) {
    status.textContent = `Max. Gewicht muss mindestens ${fmt(100 / cols.length, 1)} % betragen, damit 100 % verteilt werden können.`;
    status.className = 'optimizerStatus warning';
    return;
  }
  if (!trainStart || !trainEnd || trainEnd <= trainStart) {
    status.textContent = 'Bitte einen gültigen Trainingszeitraum wählen.';
    status.className = 'optimizerStatus warning';
    return;
  }

  btn.disabled = true;
  status.textContent = 'Optimiere Gewichte auf MSCI WORLD …';
  status.className = 'optimizerStatus';

  window.setTimeout(() => {
    const dates = D.risk.records.map(r => r.Datum);
    const windowLength = val('rollingWindow');
    const minp = val('minPeriods');
    const clip = val('clipValue');
    const zscores = { Datum: dates };
    cols.forEach(c => {
      const levels = series(D.risk.records, c);
      zscores[c] = rollZ(levels, windowLength, minp, clip, $(safeId('inv_', c)).checked ? -1 : 1);
    });

    const random = seededRandom(20260703);
    const useKalman = true;
    const existing = normalizedWeightsFromInputs(cols);
    const candidates = [
      cols.map(c => existing[c]),
      Array(cols.length).fill(1 / cols.length)
    ];
    const randomDraws = cols.length <= 6 ? 1800 : 1000;
    for (let i = 0; i < randomDraws; i++) candidates.push(randomWeightsWithCap(cols.length, maxWeight, random));

    let best = { score: -Infinity, weights: null };
    for (const candidate of candidates) {
      const test = scoreCandidateWeights(zscores, cols, candidate, dates, trainStart, trainEnd, useKalman);
      if (test.score > best.score) best = { score: test.score, weights: candidate.slice() };
    }

    // Lokale Feinabstimmung um die beste zufällige Lösung.
    for (let i = 0; i < 500 && best.weights; i++) {
      const proposal = best.weights.slice();
      const from = Math.floor(random() * cols.length);
      let to = Math.floor(random() * cols.length);
      if (to === from) to = (to + 1) % cols.length;
      const amount = Math.min(proposal[from], 0.005 + random() * 0.04);
      proposal[from] -= amount;
      proposal[to] += amount;
      if (Math.max(...proposal) > maxWeight + 1e-12 || Math.min(...proposal) < -1e-12) continue;
      const test = scoreCandidateWeights(zscores, cols, proposal, dates, trainStart, trainEnd, useKalman);
      if (test.score > best.score) best = { score: test.score, weights: proposal.slice() };
    }

    if (!best.weights || !Number.isFinite(best.score)) {
      status.textContent = 'Keine Optimierung möglich: Im Trainingszeitraum gibt es zu wenige gültige Beobachtungen nach Rolling Window / Signallogik.';
      status.className = 'optimizerStatus warning';
      btn.disabled = false;
      return;
    }

    cols.forEach((c, i) => {
      const input = $(safeId('w_', c));
      if (input) input.value = (best.weights[i] * 100).toFixed(1);
    });
    updateWeightTotal();
    update();
    status.textContent = `Optimiert auf ${msci}: historische annualisierte Sharpe Ratio ${fmt(best.score, 2)} | Training: ${trainStart} bis ${trainEnd}. Die Gewichte wurden fest übernommen.`;
    status.className = 'optimizerStatus success';
    btn.disabled = false;
  }, 20);
}
function setup() {
  const dates = D.risk.records.map(r => r.Datum).concat(D.indices.records.map(r => r.Datum)).sort();
  if ($('startDate')) $('startDate').value = dates[0] || '';
  if ($('endDate')) $('endDate').value = dates[dates.length - 1] || '';
  if ($('optEndDate')) $('optEndDate').value = dates[dates.length - 1] || '';

  const equalWeight = D.risk.columns.length ? 100 / D.risk.columns.length : 0;
  $('riskList').innerHTML = D.risk.columns.map(c => `
    <div class="checkitem">
      <input id="${safeId('risk_', c)}" type="checkbox" checked>
      <span title="${c}">${c}</span>
      <div class="weightInputWrap">
        <input id="${safeId('w_', c)}" type="number" min="0" max="100" step="0.5" value="${equalWeight.toFixed(1)}" aria-label="Gewicht in Prozent für ${c}">
        <span class="weightSign">%</span>
      </div>
      <label class="row" title="Invertieren: höherer Wert bedeutet niedrigeres Risiko"><input id="${safeId('inv_', c)}" type="checkbox">inv</label>
    </div>`).join('');

  $('indexList').innerHTML = D.indices.columns.map((c, i) => `
    <label class="row"><input id="${safeId('idx_', c)}" type="checkbox" ${i < 3 ? 'checked' : ''}> ${c}</label>`).join('');

  renderMeta();
  document.querySelectorAll('input').forEach(e => e.addEventListener('change', () => { updateWeightTotal(); update(); }));
  $('updateBtn').addEventListener('click', update);
  $('optimizeBtn').addEventListener('click', optimizeWeightsForMsci);
  document.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => {
    document.querySelectorAll('.tab,.tabPage').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    $(b.dataset.tab).classList.add('active');
    setTimeout(update, 50);
  }));
  updateWeightTotal();
  update();
}
function calc() {
  const cols = selectedRisk();
  const dates = D.risk.records.map(r => r.Datum);
  const window = val('rollingWindow');
  const minp = val('minPeriods');
  const clip = val('clipValue');
  const zscores = { Datum: dates };

  // Fixed: only level Z-scores; no daily changes / percentage changes.
  cols.forEach(c => {
    const levels = series(D.risk.records, c);
    zscores[c] = rollZ(levels, window, minp, clip, $(safeId('inv_', c)).checked ? -1 : 1);
  });

  const normalized = normalizedWeightsFromInputs(cols);
  const result = compositeFromWeights(zscores, cols, normalized);
  const riskBands = historicalRiskBands(result.kalmanRisk);
  return { dates, zscores, ...result, adaptiveThreshold: riskBands.high, riskBands, cols };
}
function dateMask(d) {
  const start = $('startDate')?.value || '';
  const end = $('endDate')?.value || '9999-12-31';
  return d >= start && d <= end;
}
function commonLayout(height) {
  return {
    height,
    margin: { t: 18, r: 56, b: 50, l: 62 },
    hovermode: 'x unified',
    legend: { orientation: 'h', x: 0, xanchor: 'left', y: 1.11, yanchor: 'bottom', font: { size: 11 } },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    xaxis: { gridcolor: '#edf0f4', zeroline: false },
    yaxis: { gridcolor: '#edf0f4', zeroline: false }
  };
}
function update() {
  const C = calc();
  const mask = C.dates.map(dateMask);
  const ds = C.dates.filter((_, i) => mask[i]);
  const zTr = C.cols.map(c => ({ x: ds, y: C.zscores[c].filter((_, i) => mask[i]), mode: 'lines', name: c, line: { width: 1.8 } }));
  Plotly.react('zChart', zTr, { ...commonLayout(410), yaxis: { title: 'Rolling Z-Score', gridcolor: '#edf0f4', zeroline: true, zerolinecolor: '#cbd5e1' } }, { responsive: true });

  const compositeLayout = commonLayout(410);
  compositeLayout.yaxis = { title: 'Kalman Composite Z-Score', gridcolor: '#edf0f4', zeroline: true, zerolinecolor: '#cbd5e1' };
  compositeLayout.yaxis2 = { title: 'Kalman Risk Indicator (%)', overlaying: 'y', side: 'right', range: [0, 100], gridcolor: '#edf0f4' };
  Plotly.react('compositeChart', [
    { x: ds, y: C.kalmanComp.filter((_, i) => mask[i]), name: 'Kalman Composite Z-Score', yaxis: 'y', mode: 'lines', line: { width: 2.2, color: '#2563eb' } },
    { x: ds, y: C.kalmanRisk.filter((_, i) => mask[i]), name: 'Kalman Risk Indicator (%)', yaxis: 'y2', mode: 'lines', line: { dash: 'dash', width: 2.2, color: '#111827' } }
  ], compositeLayout, { responsive: true });

  const lastRisk = [...C.kalmanRisk].reverse().find(x => x != null);
  const lastZ = [...C.kalmanComp].reverse().find(x => x != null);
  $('kpiZ').textContent = fmt(lastZ);
  $('kpiRisk').textContent = fmt(lastRisk, 1);
  $('kpiN').textContent = C.cols.length;
  updateKalmanCharts(C, mask, ds);
  updateStrategy(C);
}
function updateKalmanCharts(C, mask, ds) {
  const filteredRisk = C.kalmanRisk.filter((_, i) => mask[i]);
  const filteredZ = C.kalmanComp.filter((_, i) => mask[i]);
  const lowBand = C.riskBands.low.filter((_, i) => mask[i]);
  const midBand = C.riskBands.middle.filter((_, i) => mask[i]);
  const highBand = C.riskBands.high.filter((_, i) => mask[i]);

  const probLayout = commonLayout(420);
  probLayout.yaxis = { title: 'Kalman Normal-CDF Wahrscheinlichkeit (%)', range: [0, 100], gridcolor: '#edf0f4', zeroline: false };
  Plotly.react('kalmanProbabilityChart', [
    { x: ds, y: filteredRisk, name: 'Kalman Risk Indicator (%)', mode: 'lines', line: { width: 2.6, color: '#2563eb' } },
    { x: ds, y: lowBand, name: 'Niedrig (historisches 33%-Quantil)', mode: 'lines', line: { width: 1.4, dash: 'dot', color: '#16a34a' } },
    { x: ds, y: midBand, name: 'Mäßig (historischer Median)', mode: 'lines', line: { width: 1.5, dash: 'dot', color: '#f59e0b' } },
    { x: ds, y: highBand, name: `High Risk (historisches ${Math.round(C.riskBands.highQ * 100)}%-Quantil)`, mode: 'lines', line: { width: 2, dash: 'dot', color: '#dc2626' } }
  ], probLayout, { responsive: true });

  const zLayout = commonLayout(420);
  zLayout.yaxis = { title: 'Kalman Composite Z-Score', gridcolor: '#edf0f4', zeroline: true, zerolinecolor: '#cbd5e1' };
  Plotly.react('kalmanZChart', [
    { x: ds, y: filteredZ, name: 'Kalman Composite Z-Score', mode: 'lines', line: { width: 2.6, color: '#2563eb' } }
  ], zLayout, { responsive: true });

  const lastFilteredZ = [...C.kalmanComp].reverse().find(x => x != null);
  const lastFilteredRisk = [...C.kalmanRisk].reverse().find(x => x != null);
  const lastThreshold = [...C.adaptiveThreshold].reverse().find(x => x != null);
  $('kpiKalmanZ').textContent = fmt(lastFilteredZ);
  $('kpiKalmanRisk').textContent = fmt(lastFilteredRisk, 1);
  $('kpiKalmanThreshold').textContent = fmt(lastThreshold, 1);
}
function updateStrategy(C) {
  const idx = selectedIndices();
  const riskSeries = C.kalmanRisk;
  const thresholdSeries = C.riskBands.high;
  const compMap = new Map(C.dates.map((d, i) => [d, { risk: riskSeries[i], threshold: thresholdSeries[i], low: C.riskBands.low[i], middle: C.riskBands.middle[i] }]));
  const rows = D.indices.records
    .map(r => Object.assign({ risk: compMap.get(r.Datum)?.risk ?? null, threshold: compMap.get(r.Datum)?.threshold ?? null, low: compMap.get(r.Datum)?.low ?? null, middle: compMap.get(r.Datum)?.middle ?? null }, r))
    .filter(r => dateMask(r.Datum));
  const dates = rows.map(r => r.Datum);

  const strategyTraces = [];
  idx.forEach(c => {
    const values = rows.map(r => r[c]);
    const base = values.find(v => v != null && v !== 0);
    if (base != null) strategyTraces.push({ x: dates, y: values.map(v => v != null ? v / base * 100 : null), name: c, mode: 'lines', line: { width: 1.8 } });
  });
  strategyTraces.push({ x: dates, y: rows.map(r => r.risk), name: 'Kalman Risk Indicator (%)', mode: 'lines', yaxis: 'y2', line: { color: '#111827', width: 2.5 } });
  strategyTraces.push({ x: dates, y: rows.map(r => r.low), name: 'Niedrig (historisch)', mode: 'lines', yaxis: 'y2', line: { dash: 'dot', color: '#16a34a', width: 1.2 } });
  strategyTraces.push({ x: dates, y: rows.map(r => r.middle), name: 'Mäßig (historischer Median)', mode: 'lines', yaxis: 'y2', line: { dash: 'dot', color: '#f59e0b', width: 1.3 } });
  strategyTraces.push({ x: dates, y: rows.map(r => r.threshold), name: `High Risk (historisches ${Math.round(C.riskBands.highQ * 100)}%-Quantil)`, mode: 'lines', yaxis: 'y2', line: { dash: 'dot', color: '#dc2626', width: 2 } });

  const topLayout = commonLayout(500);
  topLayout.yaxis = { title: 'Global Indices (Basis = 100)', gridcolor: '#edf0f4', zeroline: false };
  topLayout.yaxis2 = { title: 'Kalman Risk Indicator / historische Bänder (%)', overlaying: 'y', side: 'right', range: [0, 100], showgrid: false };
  Plotly.react('strategyChart', strategyTraces, topLayout, { responsive: true });

  const signals = rows.map(r => r.risk == null || r.threshold == null ? null : r.risk <= r.threshold);
  const validSignals = signals.filter(x => x != null);
  const lastRow = [...rows].reverse().find(r => r.risk != null && r.threshold != null);
  $('kpiSignal').textContent = !lastRow ? 'n/a' : (lastRow.risk <= lastRow.threshold ? 'Long' : 'Nicht investiert');
  $('kpiRisk2').textContent = fmt(lastRow?.risk, 1);
  $('kpiAdaptiveThreshold').textContent = fmt(lastRow?.threshold, 1);
  $('kpiRiskRegime').textContent = !lastRow ? 'n/a' : classifyRiskRegime(lastRow.risk, lastRow);

  const performance = [], excess = [];
  idx.forEach(c => {
    const values = rows.map(r => r[c]);
    const first = rows.findIndex(r => r.risk != null && r.threshold != null && r[c] != null);
    if (first < 0) return;
    let benchmark = 100, strategy = 100;
    const benchmarkSeries = Array(rows.length).fill(null);
    const strategySeries = Array(rows.length).fill(null);
    const excessSeries = Array(rows.length).fill(null);
    benchmarkSeries[first] = 100;
    strategySeries[first] = 100;
    excessSeries[first] = 0;

    for (let i = first + 1; i < rows.length; i++) {
      const ret = (values[i] != null && values[i - 1] != null && values[i - 1] !== 0) ? values[i] / values[i - 1] - 1 : 0;
      benchmark *= 1 + ret;
      const signal = $('prevSignal').checked ? signals[i - 1] : signals[i];
      if (signal === true) strategy *= 1 + ret;
      benchmarkSeries[i] = benchmark;
      strategySeries[i] = strategy;
      excessSeries[i] = benchmark ? (strategy / benchmark - 1) * 100 : null;
    }

    performance.push({ x: dates, y: strategySeries, name: `Strategie – ${c}`, mode: 'lines', line: { width: 2.3 } });
    performance.push({ x: dates, y: benchmarkSeries, name: `Buy-and-Hold – ${c}`, mode: 'lines', line: { dash: 'dot', width: 1.8 } });
    excess.push({ x: dates, y: excessSeries, name: c, mode: 'lines', line: { width: 2 } });
  });

  const perfLayout = commonLayout(385);
  perfLayout.yaxis = { title: 'Indexiert (Basis = 100)', gridcolor: '#edf0f4', zeroline: false };
  Plotly.react('performanceChart', performance, perfLayout, { responsive: true });

  excess.push({ x: dates, y: dates.map(() => 0), name: '0 %', mode: 'lines', line: { dash: 'dot', color: '#64748b', width: 1.3 } });
  const excessLayout = commonLayout(385);
  excessLayout.yaxis = { title: 'Excess Return (%)', gridcolor: '#edf0f4', zeroline: true, zerolinecolor: '#94a3b8' };
  Plotly.react('excessChart', excess, excessLayout, { responsive: true });
}
function renderMeta() {
  function table(meta) {
    return `<table><thead><tr><th>Dashboard Name</th><th>Field</th><th>Ticker</th><th>Region</th><th>Type</th></tr></thead><tbody>${meta.map(m => `<tr><td>${m.dashboard_name}</td><td>${m.field}</td><td>${m.ticker}</td><td>${m.region}</td><td>${m.type}</td></tr>`).join('')}</tbody></table>`;
  }
  $('riskMeta').innerHTML = table(D.risk.metadata);
  $('indexMeta').innerHTML = table(D.indices.metadata);
}

loadData().then(setup).catch(err => {
  document.body.innerHTML = `<main><section class="panel"><h1>Fehler beim Laden</h1><p>${err.message}</p><p>Prüfe, ob <code>Data.xlsx</code> im gleichen Ordner wie <code>index.html</code> liegt.</p></section></main>`;
});
