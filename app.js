let D = null;
const RISK_SHEET = 'data (risk measures)';
const INDEX_SHEET = 'data (global indices)';
const SERIES_NAME_ROW = 5;
const DATA_START_ROW = 6;
const $ = id => document.getElementById(id);

function fmt(x,d=2){return x==null||Number.isNaN(x)?'n/a':x.toLocaleString('de-DE',{minimumFractionDigits:d,maximumFractionDigits:d});}
function clean(v){return v==null?'':String(v).trim();}
function excelDateToISO(v){
  if(v==null || v==='') return null;
  if(typeof v==='number'){
    const d = XLSX.SSF.parse_date_code(v);
    if(!d) return null;
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if(m){
    const y = m[3].length===2 ? Number('20'+m[3]) : Number(m[3]);
    return `${y}-${String(Number(m[2])).padStart(2,'0')}-${String(Number(m[1])).padStart(2,'0')}`;
  }
  const d = new Date(s);
  if(!Number.isNaN(d.getTime())) return d.toISOString().slice(0,10);
  return null;
}
function toNumber(v){
  if(v==null || v==='') return null;
  if(typeof v==='number') return Number.isFinite(v)?v:null;
  let s = String(v).trim();
  if(['nan','NaN','None','#N/A','#N/A Field Not Applicable','#VALUE!'].includes(s)) return null;
  if(s.includes(',')) s = s.replaceAll('.','').replace(',','.');
  const n = Number(s);
  return Number.isFinite(n)?n:null;
}
function makeUnique(names){
  const counts={}, out=[];
  for(const name of names){
    const base = name || 'Unnamed';
    counts[base] = (counts[base]||0)+1;
    out.push(counts[base]===1 ? base : `${base}_${counts[base]}`);
  }
  return out;
}
function parseSheet(workbook, sheetName){
  const ws = workbook.Sheets[sheetName];
  if(!ws) throw new Error(`Sheet fehlt: ${sheetName}`);
  const raw = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
  const originalNames=[], valueCols=[], metaRows=[];
  const maxCols = Math.max(...raw.map(r=>r.length));
  for(let c=1;c<maxCols;c++){
    const seriesName = clean(raw[SERIES_NAME_ROW]?.[c]);
    if(!seriesName) continue;
    const values = raw.slice(DATA_START_ROW).map(r => toNumber(r?.[c]));
    if(values.every(v=>v==null)) continue;
    originalNames.push(seriesName);
    valueCols.push(values);
    metaRows.push({
      excel_column_number: c+1,
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
  for(let r=DATA_START_ROW; r<raw.length; r++){
    const iso = excelDateToISO(raw[r]?.[0]);
    if(!iso) continue;
    const row = {Datum: iso};
    unique.forEach((name,i)=> row[name] = valueCols[i][r-DATA_START_ROW]);
    rows.push(row);
  }
  rows.sort((a,b)=>a.Datum.localeCompare(b.Datum));
  const dedup = Array.from(new Map(rows.map(r=>[r.Datum,r])).values());
  metaRows.forEach((m,i)=>m.dashboard_name=unique[i]);
  return {records: dedup, columns: unique, metadata: metaRows};
}
async function loadData(){
  const res = await fetch('Data.xlsx', {cache:'no-store'});
  if(!res.ok) throw new Error('Data.xlsx konnte nicht geladen werden. Liegt sie im gleichen Ordner wie index.html?');
  const ab = await res.arrayBuffer();
  const workbook = XLSX.read(ab, {type:'array', cellDates:false});
  D = {risk: parseSheet(workbook, RISK_SHEET), indices: parseSheet(workbook, INDEX_SHEET)};
}
function erf(x){const s=x>=0?1:-1; x=Math.abs(x); const a1=.254829592,a2=-.284496736,a3=1.421413741,a4=-1.453152027,a5=1.061405429,p=.3275911; const t=1/(1+p*x); const y=1-(((((a5*t+a4)*t)+a3)*t+a2)*t+a1)*t*Math.exp(-x*x); return s*y;}
function cdf(z){return .5*(1+erf(z/Math.sqrt(2)))*100;}
function safeId(prefix,c){return prefix + btoa(unescape(encodeURIComponent(c))).replaceAll('=','');}
function getSelected(prefix, cols){return cols.filter(c=>$(safeId(prefix,c))?.checked);}
function val(id){return Number($(id).value);}
function series(records,col){return records.map(r=>r[col]);}
function transform(arr,method){let out=arr.slice(); if(method==='Tagesveränderung') out=arr.map((v,i)=>i&&v!=null&&arr[i-1]!=null?v-arr[i-1]:null); if(method==='Prozentuale Veränderung') out=arr.map((v,i)=>i&&v!=null&&arr[i-1]!=null&&arr[i-1]!==0?v/arr[i-1]-1:null); return out;}
function rollZ(arr,window,minp,clip,dir){let out=[]; for(let i=0;i<arr.length;i++){let vals=[]; for(let j=Math.max(0,i-window+1);j<=i;j++) if(arr[j]!=null&&!Number.isNaN(arr[j])) vals.push(arr[j]); if(vals.length<minp){out.push(null);continue} let mean=vals.reduce((a,b)=>a+b,0)/vals.length; let sd=Math.sqrt(vals.reduce((a,b)=>a+(b-mean)**2,0)/vals.length); let z=sd?((arr[i]-mean)/sd)*dir:null; if(z!=null) z=Math.max(-clip,Math.min(clip,z)); out.push(z)} return out;}
function selectedRisk(){return getSelected('risk_',D.risk.columns);}
function selectedIndices(){return getSelected('idx_',D.indices.columns);}
function setup(){let dates=D.risk.records.map(r=>r.Datum).concat(D.indices.records.map(r=>r.Datum)).sort(); $('startDate').value=dates[0]; $('endDate').value=dates[dates.length-1]; $('riskList').innerHTML=D.risk.columns.map((c,i)=>`<div class="checkitem"><input id="${safeId('risk_',c)}" type="checkbox" ${i<4?'checked':''}><span title="${c}">${c}</span><input id="${safeId('w_',c)}" type="number" min="0" step="0.25" value="1"><label class="row" title="invertieren"><input id="${safeId('inv_',c)}" type="checkbox">inv</label></div>`).join(''); $('indexList').innerHTML=D.indices.columns.map((c,i)=>`<label class="row"><input id="${safeId('idx_',c)}" type="checkbox" ${i<3?'checked':''}> ${c}</label>`).join(''); renderMeta(); document.querySelectorAll('input,select').forEach(e=>e.addEventListener('change',update)); $('updateBtn').addEventListener('click',update); document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>{document.querySelectorAll('.tab,.tabPage').forEach(x=>x.classList.remove('active')); b.classList.add('active'); $(b.dataset.tab).classList.add('active'); setTimeout(update,50)})); update();}
function calc(){const cols=selectedRisk(); const dates=D.risk.records.map(r=>r.Datum); const method=$('zMethod').value, window=val('rollingWindow'), minp=val('minPeriods'), clip=val('clipValue'); let zscores={Datum:dates}; cols.forEach(c=>{const arr=transform(series(D.risk.records,c),method); zscores[c]=rollZ(arr,window,minp,clip,$(safeId('inv_',c)).checked?-1:1)}); let weights={}, total=0; cols.forEach(c=>{weights[c]=Math.max(0,Number($(safeId('w_',c)).value)||0); total+=weights[c]}); if(total===0){cols.forEach(c=>weights[c]=1/cols.length)}else{cols.forEach(c=>weights[c]/=total)} let comp=[], risk=[]; for(let i=0;i<dates.length;i++){let s=0,w=0; cols.forEach(c=>{let z=zscores[c][i]; if(z!=null){s+=z*weights[c]; w+=weights[c]}}); let cz=w?s/w:null; comp.push(cz); risk.push(cz==null?null:($('scoreMethod').value==='Normal-CDF Wahrscheinlichkeit'?cdf(cz):((Math.max(-clip,Math.min(clip,cz))+clip)/(2*clip))*100));} return {dates,zscores,comp,risk,cols};}
function dateMask(d){return d>=$('startDate').value&&d<=$('endDate').value;}
function update(){let C=calc(); let mask=C.dates.map(dateMask); let ds=C.dates.filter((_,i)=>mask[i]); let zTr=C.cols.map(c=>({x:ds,y:C.zscores[c].filter((_,i)=>mask[i]),mode:'lines',name:c})); Plotly.react('zChart',zTr,{margin:{t:20},hovermode:'x unified',legend:{orientation:'h'}}); Plotly.react('compositeChart',[{x:ds,y:C.comp.filter((_,i)=>mask[i]),name:'Composite Risk Z-Score',yaxis:'y',mode:'lines'},{x:ds,y:C.risk.filter((_,i)=>mask[i]),name:'Risk Indicator %',yaxis:'y2',mode:'lines',line:{dash:'dash'}}],{margin:{t:20},hovermode:'x unified',legend:{orientation:'h'},yaxis:{title:'Composite Z'},yaxis2:{title:'Risk Indicator %',overlaying:'y',side:'right',range:[0,100]}}); let lastRisk=[...C.risk].reverse().find(x=>x!=null), lastZ=[...C.comp].reverse().find(x=>x!=null); $('kpiZ').textContent=fmt(lastZ); $('kpiRisk').textContent=fmt(lastRisk,1); $('kpiN').textContent=C.cols.length; updateStrategy(C);}
function updateStrategy(C){const idx=selectedIndices(), th=val('threshold'); let compMap=new Map(C.dates.map((d,i)=>[d,{risk:C.risk[i],comp:C.comp[i]}])); let rows=D.indices.records.map(r=>Object.assign({risk:compMap.get(r.Datum)?.risk??null},r)).filter(r=>dateMask(r.Datum)); let dates=rows.map(r=>r.Datum); let traces=[]; idx.forEach(c=>{let vals=rows.map(r=>r[c]); let base=vals.find(v=>v!=null&&v!==0); traces.push({x:dates,y:vals.map(v=>v!=null&&base?v/base*100:null),name:c,mode:'lines'});}); traces.push({x:dates,y:rows.map(r=>r.risk),name:'Risk Indicator %',mode:'lines',yaxis:'y2',line:{dash:'dash',color:'black'}}); traces.push({x:dates,y:dates.map(_=>th),name:'Grenze',mode:'lines',yaxis:'y2',line:{dash:'dot',color:'gray'}}); Plotly.react('strategyChart',traces,{margin:{t:20},hovermode:'x unified',legend:{orientation:'h'},yaxis:{title:'Global Indices, indexiert auf 100'},yaxis2:{title:'Risk Indicator %',overlaying:'y',side:'right',range:[0,100]}}); let signals=rows.map(r=>r.risk==null?null:r.risk<=th); let valid=signals.filter(x=>x!=null); let lastRisk=[...rows].reverse().find(r=>r.risk!=null)?.risk; $('kpiSignal').textContent=lastRisk==null?'n/a':(lastRisk<=th?'Long':'Nicht investiert'); $('kpiRisk2').textContent=fmt(lastRisk,1); $('kpiLongShare').textContent=valid.length?(valid.filter(Boolean).length/valid.length).toLocaleString('de-DE',{style:'percent',maximumFractionDigits:1}):'n/a'; let excess=[]; idx.forEach(c=>{let vals=rows.map(r=>r[c]); let first=rows.findIndex(r=>r.risk!=null&&r[c]!=null); if(first<0)return; let b=100,s=100,ys=Array(rows.length).fill(null); ys[first]=0; for(let i=first+1;i<rows.length;i++){let ret=(vals[i]!=null&&vals[i-1]!=null&&vals[i-1]!==0)?vals[i]/vals[i-1]-1:0; b*=1+ret; let sig=$('prevSignal').checked?signals[i-1]:signals[i]; if(sig===true) s*=1+ret; ys[i]=b?((s/b)-1)*100:null;} excess.push({x:dates,y:ys,name:c,mode:'lines'});}); excess.push({x:dates,y:dates.map(_=>0),name:'0%',mode:'lines',line:{dash:'dot',color:'gray'}}); Plotly.react('excessChart',excess,{margin:{t:20},hovermode:'x unified',legend:{orientation:'h'},yaxis:{title:'Excess Return vs. Buy-and-Hold (%)'}});}
function renderMeta(){function table(meta){return `<table><thead><tr><th>Dashboard Name</th><th>Field</th><th>Ticker</th><th>Region</th><th>Type</th></tr></thead><tbody>${meta.map(m=>`<tr><td>${m.dashboard_name}</td><td>${m.field}</td><td>${m.ticker}</td><td>${m.region}</td><td>${m.type}</td></tr>`).join('')}</tbody></table>`} $('riskMeta').innerHTML=table(D.risk.metadata); $('indexMeta').innerHTML=table(D.indices.metadata);}
loadData().then(setup).catch(err=>{document.body.innerHTML=`<main><section class="panel"><h1>Fehler beim Laden</h1><p>${err.message}</p><p>Prüfe, ob <code>Data.xlsx</code> im gleichen Ordner wie <code>index.html</code> liegt.</p></section></main>`;});
