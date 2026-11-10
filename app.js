/* Excel Practice Lab — all frontend, no server.
   Data: data/formulas.json (42 formulas) + data/datasets.json (4 datasets).
   The formula checker below is a small Excel-compatible evaluator written for this page. */
'use strict';

const PLAYLIST = 'PLbvk4Aqr_n70';
const GITHUB_REPO = 'https://github.com/dralihassankhan1990-spec/excel-for-IT-project-management';
const YT_CHANNEL = 'https://www.youtube.com/@SochAndSystems';

let FORMULAS = [];
let DATASETS = [];          // {id,title,desc,headers,rows,custom?}
let activeDs = null;        // current dataset object
let activeFormula = null;   // current formula object
let activeCat = 'All';
let searchQ = '';

/* ---------------- exercises: one per formula ----------------
   Model formulas use {Header} tokens (replaced by the real column letter
   for the active dataset) and {N} = last data row number.
   {ID} = first column of the dataset. */
const EXERCISES = {
 1:{q:'Who is the Owner of the first item in the list?', model:'=XLOOKUP({ID1},{WorkID}2:{WorkID}{N},{Owner}2:{Owner}{N},"Not Found")', need:['WorkID','Owner']},
 2:{q:'Which Project does the first item belong to?', model:'=VLOOKUP({ID1},{WorkID}2:{Project}{N},2,FALSE)', need:['WorkID','Project']},
 3:{q:'What is the DueDate of the first item?', model:'=INDEX({DueDate}2:{DueDate}{N},MATCH({ID1},{WorkID}2:{WorkID}{N},0))', need:['WorkID','DueDate']},
 4:{q:'In which position is the first item in the list? (It is in row 2, so the answer is 1.)', model:'=MATCH({ID1},{WorkID}2:{WorkID}{N},0)', need:['WorkID']},
 5:{q:'Total ActualHrs for the project "Inventory API"?', model:'=SUMIF({Project}2:{Project}{N},"Inventory API",{ActualHrs}2:{ActualHrs}{N})', need:['Project','ActualHrs']},
 6:{q:'Total ActualHrs for "HRMS Upgrade" in Sprint 2?', model:'=SUMIFS({ActualHrs}2:{ActualHrs}{N},{Project}2:{Project}{N},"HRMS Upgrade",{Sprint}2:{Sprint}{N},2)', need:['Project','Sprint','ActualHrs']},
 7:{q:'How many tasks are "Delayed"?', model:'=COUNTIF({Status}2:{Status}{N},"Delayed")', need:['Status']},
 8:{q:'How many tasks are Delayed AND Critical?', model:'=COUNTIFS({Status}2:{Status}{N},"Delayed",{Priority}2:{Priority}{N},"Critical")', need:['Status','Priority']},
 9:{q:'Average ActualHrs of tasks with the same Status as the first task (row 2)?', model:'=AVERAGEIF({Status}2:{Status}{N},{Status}2,{ActualHrs}2:{ActualHrs}{N})', need:['Status','ActualHrs']},
 10:{q:'Average VarianceHrs of tasks with the same Status and Project as the first task?', model:'=AVERAGEIFS({VarianceHrs}2:{VarianceHrs}{N},{Status}2:{Status}{N},{Status}2,{Project}2:{Project}{N},{Project}2)', need:['VarianceHrs','Status','Project']},
 11:{q:'Row 2: show "Over" if ActualHrs is more than PlannedHrs, else "OK".', model:'=IF({ActualHrs}2>{PlannedHrs}2,"Over","OK")', need:['ActualHrs','PlannedHrs']},
 12:{q:'Row 2: if %Complete is 1 show "Done", if 0.5 or more show "Half", otherwise "Start".', model:'=IFS({%Complete}2=1,"Done",{%Complete}2>=0.5,"Half",TRUE,"Start")', need:['%Complete']},
 13:{q:'Row 3: show "ESCALATE" if the task is Delayed AND Critical, else "OK".', model:'=IF(AND({Status}3="Delayed",{Priority}3="Critical"),"ESCALATE","OK")', row:3, need:['Status','Priority']},
 14:{q:'Owner of "IT-999" (does not exist) — show "Not Found" instead of an error.', model:'=IFERROR(XLOOKUP("IT-999",{WorkID}2:{WorkID}{N},{Owner}2:{Owner}{N}),"Not Found")', need:['WorkID','Owner']},
 15:{q:'Show the DueDate of row 2 as month-year, like "May-2025".', model:'=TEXT({DueDate}2,"mmm-yyyy")', need:['DueDate']},
 16:{q:'Join the WorkID and Project of row 2 with " - " between them.', model:'=CONCAT({WorkID}2," - ",{Project}2)', need:['WorkID','Project']},
 17:{q:'Take the first 2 letters of the WorkID in row 2.', model:'=LEFT({WorkID}2,2)', need:['WorkID']},
 18:{q:'How many characters are in the Owner name in row 2?', model:'=LEN({Owner}2)', need:['Owner']},
 19:{q:'Which month number (1–12) is the DueDate of row 2?', model:'=MONTH({DueDate}2)', need:['DueDate']},
 20:{q:'Last day of the month of the DueDate in row 2?', model:'=EOMONTH({DueDate}2,0)', need:['DueDate']},
 21:{q:'Working days from StartDate to DueDate in row 2?', model:'=NETWORKDAYS({StartDate}2,{DueDate}2)', need:['StartDate','DueDate']},
 22:{q:'List the unique Projects (the first one is enough to check).', model:'=UNIQUE({Project}2:{Project}{N})', need:['Project']},
 23:{q:'Largest VarianceHrs value (sort VarianceHrs largest first, take the first).', model:'=SORT({VarianceHrs}2:{VarianceHrs}{N},1,-1)', need:['VarianceHrs']},
 24:{q:'How many rows have the same Status AND Priority as the first task (row 2)?', model:'=COUNTA(FILTER({WorkID}2:{WorkID}{N},({Status}2:{Status}{N}={Status}2)*({Priority}2:{Priority}{N}={Priority}2)))', need:['WorkID','Status','Priority']},
 25:{q:'What is the 3rd largest VarianceHrs?', model:'=LARGE({VarianceHrs}2:{VarianceHrs}{N},3)', need:['VarianceHrs']},
 26:{q:'Rank of the StoryPoints in row 2 (largest = 1)?', model:'=RANK({StoryPoints}2,{StoryPoints}2:{StoryPoints}{N},0)', need:['StoryPoints']},
 27:{q:'95th percentile of ActualHrs?', model:'=PERCENTILE({ActualHrs}2:{ActualHrs}{N},0.95)', need:['ActualHrs']},
 28:{q:'Planned cost of row 2 at 1,000 per hour?', model:'={PlannedHrs}2*1000', need:['PlannedHrs']},
 29:{q:'Total PlannedHrs for "Service Desk Migration" in Sprint 2?', model:'=SUMIFS({PlannedHrs}2:{PlannedHrs}{N},{Project}2:{Project}{N},"Service Desk Migration",{Sprint}2:{Sprint}{N},2)', need:['PlannedHrs','Project','Sprint']},
 30:{q:'How many rows have Status typed as "done" (lowercase — validation blocks this)?', model:'=COUNTIF({Status}2:{Status}{N},"done")', need:['Status']},
 31:{q:'How many rows would be highlighted red (Delayed)?', model:'=COUNTIF({Status}2:{Status}{N},"Delayed")', need:['Status']},
 32:{q:'After refresh, how many data rows are there?', model:'=COUNTA({ID}2:{ID}{N})', need:[]},
 33:{q:'Row 3: is it overdue today? Show "Overdue" if DueDate is before today, else "On track".', model:'=IF({DueDate}3<TODAY(),"Overdue","On track")', row:3, need:['DueDate']},
 34:{q:'Next working day after the DueDate in row 2?', model:'=WORKDAY({DueDate}2,1)', need:['DueDate']},
 35:{q:'One month after the DueDate in row 2?', model:'=EDATE({DueDate}2,1)', need:['DueDate']},
 36:{q:'Days between StartDate and DueDate in row 2?', model:'=DATEDIF({StartDate}2,{DueDate}2,"d")', need:['StartDate','DueDate']},
 37:{q:'Row 2: Sprint 1 → "First", Sprint 2 → "Second", anything else → "Later".', model:'=SWITCH({Sprint}2,1,"First",2,"Second","Later")', need:['Sprint']},
 38:{q:'Join all unique Modules with " | " between them.', model:'=TEXTJOIN(" | ",TRUE,UNIQUE({Module}2:{Module}{N}))', need:['Module']},
 39:{q:'Story-point weighted % complete for the whole project?', model:'=SUMPRODUCT({StoryPoints}2:{StoryPoints}{N},{%Complete}2:{%Complete}{N})/SUM({StoryPoints}2:{StoryPoints}{N})', need:['StoryPoints','%Complete']},
 40:{q:'How many tasks have no FinishDate yet?', model:'=COUNTBLANK({FinishDate}2:{FinishDate}{N})', need:['FinishDate']},
 41:{q:'Owner of "IT-999" — show "Not Found" if missing.', model:'=IFNA(XLOOKUP("IT-999",{WorkID}2:{WorkID}{N},{Owner}2:{Owner}{N}),"Not Found")', need:['WorkID','Owner']},
 42:{q:'Remove the dash from the WorkID in row 2.', model:'=SUBSTITUTE({WorkID}2,"-","")', need:['WorkID']},
};

/* ---------------- helpers ---------------- */
const $ = s => document.querySelector(s);
const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const colLetter = i => { let s=''; i++; while(i>0){ const m=(i-1)%26; s=String.fromCharCode(65+m)+s; i=Math.floor((i-1)/26);} return s; };
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const EPOCH = Date.UTC(1899,11,30) / 86400000; // Excel serial 0 = 1899-12-30
function toSerial(v){
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && DATE_RE.test(v.trim())){
    const [y,m,d] = v.trim().split('-').map(Number);
    return Math.round(Date.UTC(y,m-1,d)/86400000 - EPOCH);
  }
  return null;
}
function fromSerial(n){
  const ms = (n + EPOCH) * 86400000;
  const d = new Date(ms);
  return d.toISOString().slice(0,10);
}
function isDateCol(h){ return /date/i.test(h || ''); }
function fmtVal(v, dateHint){
  if (v && typeof v === 'object' && v.err) return v.err;
  if (Array.isArray(v)){
    const flat = v.flat(Infinity).filter(x => x !== '' && x !== null && x !== undefined);
    if (!flat.length) return '(empty)';
    const first = flat.slice(0,5).map(x => fmtVal(x, dateHint)).join(', ');
    return flat.length > 5 ? first + ' … (' + flat.length + ' values)' : first;
  }
  if (typeof v === 'number'){
    if (dateHint && Number.isInteger(v) && v >= 43831 && v <= 47481) return fromSerial(v);
    return (Math.round(v*100)/100).toString();
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v === null || v === undefined || v === '') return '(blank)';
  return String(v);
}
function valuesEqual(a, b){
  const fa = Array.isArray(a) ? a.flat(Infinity).filter(x=>x!==''&&x!=null) : [a];
  const fb = Array.isArray(b) ? b.flat(Infinity).filter(x=>x!==''&&x!=null) : [b];
  if (!fa.length || !fb.length) return fa.length === fb.length;
  const x = fa[0], y = fb[0];
  if (typeof x === 'number' && typeof y === 'number') return Math.abs(x-y) < 0.011;
  return String(x).trim().toLowerCase() === String(y).trim().toLowerCase();
}

/* ---------------- data loading ---------------- */
async function loadData(){
  const [f, d] = await Promise.all([
    fetch('data/formulas.json').then(r => r.json()),
    fetch('data/datasets.json').then(r => r.json()),
  ]);
  FORMULAS = f; DATASETS = d;
  activeDs = DATASETS[0];
}
function headerMap(ds){
  const m = {};
  ds.headers.forEach((h,i) => { m[h] = colLetter(i); });
  m['ID'] = colLetter(0);
  // fallbacks so one exercise works on every dataset: the ID column is the
  // first column, AssignedTo is the Owner, Severity is the Priority, and a
  // bug's ReportedDate plays the StartDate role.
  if (!m['WorkID']) m['WorkID'] = colLetter(0);
  if (!m['Owner'] && m['AssignedTo']) m['Owner'] = m['AssignedTo'];
  if (!m['Priority'] && m['Severity']) m['Priority'] = m['Severity'];
  if (!m['Module'] && m['Category']) m['Module'] = m['Category'];
  if (!m['StartDate'] && m['ReportedDate']) m['StartDate'] = m['ReportedDate'];
  return m;
}
function resolveModel(model, ds){
  const map = headerMap(ds);
  const lastRow = ds.rows.length + 1; // header is row 1
  // {N} first: the header pattern would also match a bare "N"
  // {ID1} = the ID value of the first data row, quoted (works on every dataset)
  const firstId = ds.rows.length ? JSON.stringify(String(ds.rows[0][0])) : '""';
  return model.replace(/\{N\}/g, String(lastRow))
    .replace(/\{ID1\}/g, firstId)
    .replace(/\{([A-Za-z%]+)\}/g, (_, h) => map[h] || h);
}
function missingHeaders(need, ds){
  const map = headerMap(ds);
  return (need||[]).filter(h => !map[h]);
}

/* ---------------- task row + paging + download (v2) ---------------- */
const PAGE_SIZE = 50;
let dsPage = 0;
function realHeader(h, ds){
  if (ds.headers.includes(h)) return h;
  if (h === 'WorkID' || h === 'ID') return ds.headers[0];
  const fb = {Owner:'AssignedTo', Priority:'Severity', Module:'Category', StartDate:'ReportedDate'};
  return fb[h] || null;
}
function isRowTask(model){
  return /\{ID1\}/.test(model) || /\{[A-Za-z%]+\}[23](?![0-9:])/.test(model);
}
function taskRowHTML(ex, ds){
  const n = ex.row || 2;
  const row = ds.rows[n - 2];
  if (!row) return '';
  const cols = ['ID', ...(ex.need || [])]
    .map(h => realHeader(h, ds))
    .filter((h, i, a) => h && a.indexOf(h) === i);
  const cells = cols.map(h => {
    const i = ds.headers.indexOf(h);
    return {letter: colLetter(i), header: h, value: row[i]};
  });
  const caption = isRowTask(ex.model)
    ? 'This is the row your task is about.'
    : 'First row of the dataset — this task works across all rows.';
  return `<div class="task-row"><b>Task row — row ${n} of the ${esc(ds.title)} dataset</b>
    <span class="muted small"> ${esc(caption)}</span>
    <table class="mini"><thead><tr>${cells.map(c => `<th><span class="col-letter">Column ${c.letter}</span>${esc(c.header)}</th>`).join('')}</tr></thead>
    <tbody><tr>${cells.map(c => `<td>${c.value === null || c.value === undefined || c.value === '' ? '(blank)' : esc(c.value)}</td>`).join('')}</tr></tbody></table></div>`;
}
function toCSV(ds){
  const q = v => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [ds.headers.map(q).join(',')];
  for (const r of ds.rows) lines.push(r.map(q).join(','));
  return lines.join('\n');
}
function downloadDataset(){
  const ds = activeDs;
  const blob = new Blob([toCSV(ds)], {type: 'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (ds.id && ds.id !== 'custom'
    ? ds.id
    : ds.title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')) + '.csv';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

/* ---------------- formula browser ---------------- */
function renderCats(){
  const cats = ['All', ...new Set(FORMULAS.map(f => f.cat))];
  $('#catChips').innerHTML = cats.map(c =>
    `<button class="chip${c===activeCat?' active':''}" data-cat="${esc(c)}" type="button">${esc(c)}</button>`).join('');
  document.querySelectorAll('#catChips .chip').forEach(ch =>
    ch.addEventListener('click', () => { activeCat = ch.dataset.cat; renderCats(); renderGrid(); }));
}
function filteredFormulas(){
  return FORMULAS.filter(f =>
    (activeCat === 'All' || f.cat === activeCat) &&
    (!searchQ || (f.name + ' ' + f.what + ' ' + f.syntax).toLowerCase().includes(searchQ)));
}
function renderGrid(){
  const list = filteredFormulas();
  $('#formulaGrid').innerHTML = list.length ? list.map(f => `
    <article class="fcard glass${activeFormula && activeFormula.n===f.n ? ' selected':''}" data-n="${f.n}" tabindex="0">
      <span class="num">FORMULA ${String(f.n).padStart(2,'0')} · DEMO ${String(f.demo).padStart(2,'0')}</span>
      <h3>${esc(f.name)}</h3>
      <p>${esc(f.what)}</p>
      <span class="cat">${esc(f.cat)}</span>
    </article>`).join('')
    : '<p class="muted">No formula matches that search.</p>';
  document.querySelectorAll('.fcard').forEach(card => {
    const pick = () => selectFormula(Number(card.dataset.n), true);
    card.addEventListener('click', pick);
    card.addEventListener('keydown', e => { if (e.key === 'Enter') pick(); });
  });
}

/* ---------------- formula detail + practice ---------------- */
function selectFormula(n, scroll){
  activeFormula = FORMULAS.find(f => f.n === n);
  renderGrid();
  renderDetail();
  if (scroll) document.getElementById('practice').scrollIntoView({behavior:'smooth'});
}
function renderDetail(){
  const f = activeFormula;
  if (!f) return;
  const ex = EXERCISES[f.n];
  const missing = missingHeaders(ex.need, activeDs);
  const videoUrl = `https://www.youtube-nocookie.com/embed?listType=playlist&list=${PLAYLIST}&index=${f.demo}`;
  $('#detail').innerHTML = `
    <span class="cat">${esc(f.cat)}</span>
    <h2>${String(f.n).padStart(2,'0')} · ${esc(f.name)}</h2>
    <p>${esc(f.what)}</p>
    <div class="video-wrap">
      <iframe src="${videoUrl}" title="Demo video: ${esc(f.name)}" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen loading="lazy"></iframe>
    </div>
    <p class="muted">▶ Demo ${String(f.demo).padStart(2,'0')} of 42 — watch it, then try it yourself below.
    <a href="https://www.youtube.com/playlist?list=${PLAYLIST}" target="_blank" rel="noopener">Open the full playlist</a></p>
    <div class="syntax-box"><b style="color:var(--gold-soft)">Syntax</b><br><code class="formula">${esc(f.syntax)}</code></div>
    <div class="kv">
      <div class="kv-box"><b>How to read it</b>${esc(f.read)}</div>
      <div class="kv-box"><b>Example (from the course)</b><code>${esc(f.example)}</code><br><span class="muted">${esc(f.data)}</span><br>Result: <b>${esc(f.result)}</b></div>
      <div class="kv-box"><b>If it goes wrong</b>${esc(f.error)}<br><b>Fix:</b> ${esc(f.fix)}</div>
      <div class="kv-box"><b>Why it matters</b>${esc(f.benefit)}</div>
    </div>
    <div class="practice">
      <h3>✍️ Your turn — practise on the ${esc(activeDs.title)} dataset</h3>
      <div class="task"><b>Task:</b> ${esc(ex.q)}</div>
      ${taskRowHTML(ex, activeDs)}
      ${missing.length
        ? `<div class="feedback neutral show">This exercise needs the column${missing.length>1?'s':''} <b>${missing.map(esc).join(', ')}</b>, which the current dataset does not have. Switch to the <b>IT Project Data</b> dataset above to practise it — the demo and explanation still work.</div>`
        : ''}
      <div class="formula-row">
        <input id="formulaInput" type="text" placeholder='Type your formula, e.g. =XLOOKUP("IT-005", A2:A121, F2:F121, "Not Found")' autocomplete="off" spellcheck="false">
        <button id="checkBtn" class="btn btn-gold" type="button" ${missing.length?'disabled':''}>Check answer</button>
        <button id="solBtn" class="btn btn-outline" type="button">Show solution</button>
      </div>
      <div id="feedback" class="feedback"></div>
      <div id="solution" class="solution"><b>Solution (for the ${esc(activeDs.title)} dataset):</b><code id="solCode"></code>
        <p class="muted small">Column letters change with the dataset — this solution is built for the one you have open now.</p></div>
      <p class="muted small">Tip: the table in section 2 shows the column letters. Data starts at row 2; the last row is ${activeDs.rows.length + 1}. Dates work as real dates — e.g. MONTH of a DueDate cell.</p>
    </div>`;
  $('#checkBtn').addEventListener('click', checkAnswer);
  $('#solBtn').addEventListener('click', () => {
    const model = resolveModel(ex.model, activeDs);
    $('#solCode').textContent = model;
    $('#solution').classList.add('show');
  });
  $('#formulaInput').addEventListener('keydown', e => { if (e.key === 'Enter') checkAnswer(); });
}
function checkAnswer(){
  const f = activeFormula, ex = EXERCISES[f.n];
  const fb = $('#feedback');
  const input = $('#formulaInput').value.trim();
  if (!input){ fb.className = 'feedback neutral show'; fb.innerHTML = 'Type a formula first — start with <code>=</code>.'; return; }
  const model = resolveModel(ex.model, activeDs);
  const dateHint = /Date|EOMONTH|EDATE|WORKDAY|DATEDIF|NETWORKDAYS/.test(model);
  let expected, userVal, userErr = null;
  try { expected = evaluateFormula(model, activeDs); }
  catch(e){ fb.className = 'feedback neutral show'; fb.innerHTML = 'This exercise cannot be checked on the current dataset. Switch datasets and try again.'; return; }
  try { userVal = evaluateFormula(input, activeDs); }
  catch(e){ userErr = e.message || String(e); }
  if (userErr){
    fb.className = 'feedback bad show';
    fb.innerHTML = `<b>Excel would say: ${esc(userErr)}</b><br>Your formula did not run. Check brackets, commas and quotes — or press <b>Show solution</b> to compare.`;
    return;
  }
  if (userVal && userVal.err){
    if (expected && expected.err && userVal.err === expected.err){
      fb.className = 'feedback good show';
      fb.innerHTML = `✅ <b>Correct!</b> Your formula returns <span class="val">${esc(userVal.err)}</span> — exactly what the model formula returns on this dataset. (On the IT Project Data dataset the same formula finds a real match — try it there.)`;
      return;
    }
    fb.className = 'feedback bad show';
    fb.innerHTML = `<b>Your formula returns ${esc(userVal.err)}</b> — the same error Excel would show.<br>Expected: <span class="val">${esc(fmtVal(expected, dateHint))}</span>. Check the hint in “If it goes wrong” above, or press <b>Show solution</b>.`;
    return;
  }
  if (valuesEqual(userVal, expected)){
    fb.className = 'feedback good show';
    fb.innerHTML = `✅ <b>Correct!</b> Your answer: <span class="val">${esc(fmtVal(userVal, dateHint))}</span><br>Expected: <span class="val">${esc(fmtVal(expected, dateHint))}</span>. Try the same formula on another dataset — the answer will change, the method stays the same.`;
  } else {
    fb.className = 'feedback bad show';
    fb.innerHTML = `Not quite. Your answer: <span class="val">${esc(fmtVal(userVal, dateHint))}</span><br>Expected: <span class="val">${esc(fmtVal(expected, dateHint))}</span><br>Check the range and the criteria, then try again — or press <b>Show solution</b>.`;
  }
}

/* ---------------- datasets ---------------- */
function renderDsTabs(){
  $('#dsTabs').innerHTML = DATASETS.map(ds => `
    <button class="ds-tab${ds.id===activeDs.id?' active':''}" data-id="${ds.id}" type="button">
      ${esc(ds.title)}<small>${ds.rows.length} rows · ${ds.headers.length} columns</small>
    </button>`).join('');
  document.querySelectorAll('.ds-tab').forEach(t =>
    t.addEventListener('click', () => {
      activeDs = DATASETS.find(d => d.id === t.dataset.id);
      dsPage = 0;
      renderDsTabs(); renderTable();
      if (activeFormula) renderDetail();
    }));
}
function renderTable(){
  const ds = activeDs;
  $('#dsTitle').textContent = ds.title;
  const total = ds.rows.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (dsPage >= pages) dsPage = pages - 1;
  if (dsPage < 0) dsPage = 0;
  const start = dsPage * PAGE_SIZE;
  const end = Math.min(total, start + PAGE_SIZE);
  const shown = ds.rows.slice(start, end);
  $('#dsMeta').textContent = ds.desc + ` — ${total} rows · Showing rows ${total ? start + 1 : 0}–${end} of ${total}. Formulas and practice always use all ${total} rows.`;
  let html = '<thead><tr>' + ds.headers.map((h,i) =>
    `<th><span class="col-letter">Column ${colLetter(i)}</span>${esc(h)}</th>`).join('') + '</tr></thead><tbody>';
  for (const row of shown){
    html += '<tr>' + row.map((v,i) => {
      const num = typeof v === 'number';
      return `<td class="${num?'num':''}">${v===null||v===undefined||v===''?'':esc(v)}</td>`;
    }).join('') + '</tr>';
  }
  html += '</tbody>';
  $('#dataTable').innerHTML = html;
  $('#pager').innerHTML = pages > 1
    ? `<button class="btn btn-outline" id="pgPrev" type="button" ${dsPage === 0 ? 'disabled' : ''}>← Previous</button>
       <span class="muted">Page ${dsPage + 1} of ${pages}</span>
       <button class="btn btn-outline" id="pgNext" type="button" ${dsPage === pages - 1 ? 'disabled' : ''}>Next →</button>`
    : `<span class="muted">All ${total} rows on one page</span>`;
  const prev = $('#pgPrev'), next = $('#pgNext');
  if (prev) prev.addEventListener('click', () => { dsPage--; renderTable(); });
  if (next) next.addEventListener('click', () => { dsPage++; renderTable(); });
}

/* ---------------- upload own dataset ---------------- */
function parseCSV(text){
  const rows = []; let row = [], cur = '', inQ = false;
  const s = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  for (let i=0;i<s.length;i++){
    const c = s[i];
    if (inQ){
      if (c === '"'){ if (s[i+1] === '"'){ cur += '"'; i++; } else inQ = false; }
      else cur += c;
    } else if (c === '"') inQ = true;
    else if (c === ','){ row.push(cur); cur=''; }
    else if (c === '\n'){ row.push(cur); rows.push(row); row=[]; cur=''; }
    else cur += c;
  }
  if (cur !== '' || row.length){ row.push(cur); rows.push(row); }
  return rows.filter(r => r.some(x => String(x).trim() !== ''));
}
function loadCustomDataset(rows, name){
  if (rows.length < 2){ alert('Need at least a header row and one data row.'); return; }
  const headers = rows[0].map(h => String(h).trim() || 'Column');
  const data = rows.slice(1).map(r => headers.map((_,i) => {
    const v = (r[i] ?? '').toString().trim();
    if (v === '') return '';
    const num = Number(v);
    return (!isNaN(num) && /^[-+]?\d+(\.\d+)?$/.test(v)) ? num : v;
  }));
  DATASETS = DATASETS.filter(d => !d.custom);
  const ds = { id:'custom', title: name || 'My dataset', desc:'Uploaded by you — read in your browser only.', headers, rows: data, custom:true };
  DATASETS.push(ds);
  activeDs = ds;
  dsPage = 0;
  renderDsTabs(); renderTable();
  if (activeFormula) renderDetail();
  document.getElementById('datasets').scrollIntoView({behavior:'smooth'});
}
function initUpload(){
  $('#csvFile').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const rd = new FileReader();
    rd.onload = () => loadCustomDataset(parseCSV(String(rd.result)), file.name.replace(/\.csv$/i,''));
    rd.readAsText(file);
    e.target.value = '';
  });
  $('#pasteBtn').addEventListener('click', () => { $('#pasteBox').hidden = false; $('#pasteText').focus(); });
  $('#pasteCancel').addEventListener('click', () => { $('#pasteBox').hidden = true; });
  $('#pasteLoad').addEventListener('click', () => {
    loadCustomDataset(parseCSV($('#pasteText').value), 'Pasted data');
    $('#pasteBox').hidden = true;
  });
}

/* ---------------- init ---------------- */
document.addEventListener('DOMContentLoaded', async () => {
  try {
    await loadData();
  } catch(e){
    document.body.innerHTML = '<main style="max-width:640px;margin:80px auto;color:#fff;font-family:sans-serif">'
      + '<h1>Excel Practice Lab</h1><p>This page needs to be served over http (not opened as a file) so it can load its data files. '
      + 'On Vercel it works as-is. Locally, run <code>python3 -m http.server</code> in this folder and open http://localhost:8000.</p></main>';
    return;
  }
  renderCats(); renderGrid(); renderDsTabs(); renderTable(); initUpload();
  $('#dlBtn').addEventListener('click', downloadDataset);
  $('#search').addEventListener('input', e => { searchQ = e.target.value.trim().toLowerCase(); renderGrid(); });
  selectFormula(1, false); // start with XLOOKUP
});

/* ================= Excel formula evaluator =================
   Small recursive-descent parser + evaluator. Supports cell refs (A1, $A$1,
   A:A, 1:1), ranges, operators, and the functions used by the 42 exercises.
   Dates in the data are ISO strings (YYYY-MM-DD) and behave as Excel dates. */

function evaluateFormula(text, ds){
  const grid = { headers: ds.headers, rows: ds.rows };
  const src = String(text).trim();
  if (!src.startsWith('=')) throw new Error('A formula must start with =');
  const tokens = tokenize(src.slice(1));
  const parser = { toks: tokens, i: 0, grid };
  const ast = parseExpr(parser);
  if (parser.i < tokens.length) throw new Error('Unexpected "' + tokens[parser.i].v + '" — check brackets and commas.');
  return evalNode(ast, grid);
}

/* ---------- tokenizer ---------- */
function tokenize(s){
  const toks = []; let i = 0;
  const isLetter = c => /[A-Za-z_]/.test(c);
  const isDigit = c => /[0-9]/.test(c);
  while (i < s.length){
    const c = s[i];
    if (c === ' ' || c === '\t' || c === '\n'){ i++; continue; }
    if (c === '"'){
      let j = i + 1, out = '';
      while (j < s.length){
        if (s[j] === '"'){ if (s[j+1] === '"'){ out += '"'; j += 2; } else break; }
        else { out += s[j]; j++; }
      }
      if (j >= s.length) throw new Error('Unclosed " in the formula.');
      toks.push({t:'str', v:out}); i = j + 1; continue;
    }
    if (isDigit(c) || (c === '.' && isDigit(s[i+1] || ''))){
      let j = i;
      while (j < s.length && /[0-9A-Za-z_.]/.test(s[j])) j++;
      const run = s.slice(i, j);
      // a digit-start run with letters (e.g. a sheet name like 1_IT_PM_DATA) is a name, not a number
      if (/^[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?$/.test(run)){
        toks.push({t:'num', v:Number(run)}); i = j; continue;
      }
      if (s[j] === '!'){ toks.push({t:'sheet', v:run}); i = j + 1; continue; }
      toks.push({t:'name', v:run}); i = j; continue;
    }
    if (isLetter(c)){
      let j = i;
      while (j < s.length && /[A-Za-z0-9_.]/.test(s[j])) j++;
      let word = s.slice(i, j);
      // sheet prefix: Name! or 'Name'!
      if (s[j] === '!'){ toks.push({t:'sheet', v:word}); i = j + 1; continue; }
      toks.push({t:'name', v:word}); i = j; continue;
    }
    if (c === "'"){
      let j = s.indexOf("'", i + 1);
      if (j === -1) throw new Error('Unclosed sheet name.');
      if (s[j+1] === '!'){ toks.push({t:'sheet', v:s.slice(i+1, j)}); i = j + 2; continue; }
      throw new Error("Unexpected ' — check the formula.");
    }
    if (c === '$'){ i++; continue; } // $ in references is ignored
    if (c === '<' && s[i+1] === '>'){ toks.push({t:'op', v:'<>'}); i += 2; continue; }
    if (c === '<' && s[i+1] === '='){ toks.push({t:'op', v:'<='}); i += 2; continue; }
    if (c === '>' && s[i+1] === '='){ toks.push({t:'op', v:'>='}); i += 2; continue; }
    if ('+-*/%^&=<>(),:'.includes(c)){ toks.push({t:'op', v:c}); i++; continue; }
    throw new Error('Unexpected character "' + c + '".');
  }
  return toks;
}

/* ---------- parser (produces an AST) ---------- */
function peek(p){ return p.toks[p.i]; }
function next(p){ return p.toks[p.i++]; }
function expectOp(p, v){
  const t = next(p);
  if (!t || t.t !== 'op' || t.v !== v) throw new Error('Expected "' + v + '".');
}
function parseExpr(p){ return parseCompare(p); }
function parseCompare(p){
  let left = parseConcat(p);
  while (peek(p) && peek(p).t === 'op' && ['=','<>','<','>','<=','>='].includes(peek(p).v)){
    const op = next(p).v;
    const right = parseConcat(p);
    left = {type:'bin', op, left, right};
  }
  return left;
}
function parseConcat(p){
  let left = parseAdd(p);
  while (peek(p) && peek(p).t === 'op' && peek(p).v === '&'){
    next(p);
    const right = parseAdd(p);
    left = {type:'bin', op:'&', left, right};
  }
  return left;
}
function parseAdd(p){
  let left = parseMul(p);
  while (peek(p) && peek(p).t === 'op' && (peek(p).v === '+' || peek(p).v === '-')){
    const op = next(p).v;
    const right = parseMul(p);
    left = {type:'bin', op, left, right};
  }
  return left;
}
function parseMul(p){
  let left = parsePow(p);
  while (peek(p) && peek(p).t === 'op' && (peek(p).v === '*' || peek(p).v === '/')){
    const op = next(p).v;
    const right = parsePow(p);
    left = {type:'bin', op, left, right};
  }
  return left;
}
function parsePow(p){
  const base = parseUnary(p);
  if (peek(p) && peek(p).t === 'op' && peek(p).v === '^'){
    next(p);
    const exp = parseUnary(p);
    return {type:'bin', op:'^', left:base, right:exp};
  }
  return base;
}
function parseUnary(p){
  if (peek(p) && peek(p).t === 'op' && peek(p).v === '-'){ next(p); return {type:'neg', v:parseUnary(p)}; }
  if (peek(p) && peek(p).t === 'op' && peek(p).v === '+'){ next(p); return parseUnary(p); }
  return parsePostfix(p);
}
function parsePostfix(p){
  let node = parsePrimary(p);
  while (peek(p) && peek(p).t === 'op' && peek(p).v === '%'){
    next(p);
    node = {type:'bin', op:'/', left:node, right:{type:'num', v:100}};
  }
  return node;
}
const CELL_RE = /^([A-Za-z]{1,3})([0-9]+)$/;
const COL_RE = /^([A-Za-z]{1,3})$/;
function colToIndex(letters){
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function parsePrimary(p){
  const t = next(p);
  if (!t) throw new Error('Formula ends too soon — a value or bracket is missing.');
  if (t.t === 'num') return {type:'num', v:t.v};
  if (t.t === 'str') return {type:'str', v:t.v};
  if (t.t === 'op' && t.v === '('){
    const node = parseExpr(p);
    expectOp(p, ')');
    return node;
  }
  if (t.t === 'sheet'){
    // sheet name already skipped; the next token must be a ref
    return parseRefAfterSheet(p);
  }
  if (t.t === 'name'){
    const up = t.v.toUpperCase();
    if (up === 'TRUE') return {type:'bool', v:true};
    if (up === 'FALSE') return {type:'bool', v:false};
    // function call?
    if (peek(p) && peek(p).t === 'op' && peek(p).v === '('){
      next(p);
      const args = [];
      if (!(peek(p) && peek(p).t === 'op' && peek(p).v === ')')){
        while (true){
          args.push(parseExpr(p));
          if (peek(p) && peek(p).t === 'op' && peek(p).v === ','){ next(p); continue; }
          break;
        }
      }
      expectOp(p, ')');
      return {type:'func', name:up, args};
    }
    // cell reference A1 ?
    const m = CELL_RE.exec(t.v);
    if (m) return maybeRange(p, {type:'cell', col:colToIndex(m[1]), row:Number(m[2])});
    // $F$2 arrives as name "F" + number 2 (the $ signs are skipped by the tokenizer)
    if (COL_RE.test(t.v) && peek(p) && peek(p).t === 'num'){
      const rowTok = next(p);
      return maybeRange(p, {type:'cell', col:colToIndex(t.v), row:Math.floor(rowTok.v)});
    }
    // whole-column A or whole-row 1 handled in maybeRange via following ':'
    if (COL_RE.test(t.v) && peek(p) && peek(p).t === 'op' && peek(p).v === ':'){
      next(p);
      const t2 = next(p);
      if (!t2 || t2.t !== 'name') throw new Error('Expected a column after ":".');
      const m2 = COL_RE.exec(t2.v);
      if (!m2) throw new Error('Expected a column after ":".');
      return {type:'colrange', c1:colToIndex(t.v), c2:colToIndex(m2[1])};
    }
    if (/^[0-9]+$/.test(t.v)) return {type:'num', v:Number(t.v)};
    throw new Error('#NAME? — "' + t.v + '" is not a function or a cell reference.');
  }
  throw new Error('Unexpected "' + (t.v ?? '') + '" — check the formula.');
}
function parseRefAfterSheet(p){
  const t = next(p);
  if (!t || t.t !== 'name') throw new Error('Expected a cell reference after the sheet name.');
  const m = CELL_RE.exec(t.v);
  if (m) return maybeRange(p, {type:'cell', col:colToIndex(m[1]), row:Number(m[2])});
  throw new Error('Expected a cell reference after the sheet name.');
}
function maybeRange(p, first){
  if (peek(p) && peek(p).t === 'op' && peek(p).v === ':'){
    next(p);
    const t2 = next(p);
    if (!t2) throw new Error('Expected a cell reference after ":".');
    if (t2.t === 'name'){
      const m2 = CELL_RE.exec(t2.v);
      if (m2) return {type:'range', c1:first.col, r1:first.row, c2:colToIndex(m2[1]), r2:Number(m2[2])};
      const c2 = COL_RE.exec(t2.v);
      if (c2) return {type:'colrange', c1:first.col, c2:colToIndex(c2[1])};
    }
    if (t2.t === 'num') return {type:'rowrange', r1:first.row, r2:Math.floor(t2.v)};
    throw new Error('Expected a cell reference after ":".');
  }
  return first;
}

/* ---------- evaluation ---------- */
const isErr = v => v && typeof v === 'object' && 'err' in v;
const ERR = code => ({err: code});
function cellValue(grid, col, row){
  // row 1 = header, rows 2..N+1 = data
  if (row === 1) return grid.headers[col] ?? '';
  const r = grid.rows[row - 2];
  if (!r) return '';
  return r[col] ?? '';
}
function rangeValue(grid, c1, r1, c2, r2){
  const out = [];
  for (let r = Math.min(r1,r2); r <= Math.max(r1,r2); r++){
    const row = [];
    for (let c = Math.min(c1,c2); c <= Math.max(c1,c2); c++) row.push(cellValue(grid, c, r));
    out.push(row);
  }
  return out;
}
function flat(v){
  if (Array.isArray(v)) return v.flat(Infinity);
  return [v];
}
function toNum(v){
  if (isErr(v)) return v;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'string'){
    const s = v.trim();
    const d = toSerial(s);
    if (d !== null) return d;
    const n = Number(s);
    if (!isNaN(n) && s !== '') return n;
    return ERR('#VALUE!');
  }
  return ERR('#VALUE!');
}
function toStr(v){
  if (isErr(v)) return v;
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (Array.isArray(v)) return toStr(flat(v)[0]);
  return String(v);
}
function applyBin(op, a, b){
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (Array.isArray(a) || Array.isArray(b)){
    const A = Array.isArray(a) ? a : [[a]];
    const B = Array.isArray(b) ? b : [[b]];
    const rows = Math.max(A.length, B.length);
    const out = [];
    for (let i = 0; i < rows; i++){
      const ra = A[Math.min(i, A.length-1)], rb = B[Math.min(i, B.length-1)];
      const cols = Math.max(ra.length, rb.length);
      const row = [];
      for (let j = 0; j < cols; j++)
        row.push(applyBin(op, ra[Math.min(j, ra.length-1)], rb[Math.min(j, rb.length-1)]));
      out.push(row);
    }
    return out.length === 1 && out[0].length === 1 ? out[0][0] : out;
  }
  if (op === '&') return toStr(a) + toStr(b);
  if (op === '=') return looseEq(a, b);
  if (op === '<>') return !looseEq(a, b);
  if (op === '<' || op === '>' || op === '<=' || op === '>='){
    const na = toNum(a), nb = toNum(b);
    if (isErr(na)) return na; if (isErr(nb)) return nb;
    if (typeof a === 'string' && typeof b === 'string' && toSerial(a) === null && toSerial(b) === null){
      const x = a.toLowerCase(), y = b.toLowerCase();
      return op === '<' ? x < y : op === '>' ? x > y : op === '<=' ? x <= y : x >= y;
    }
    return op === '<' ? na < nb : op === '>' ? na > nb : op === '<=' ? na <= nb : na >= nb;
  }
  const na = toNum(a), nb = toNum(b);
  if (isErr(na)) return na; if (isErr(nb)) return nb;
  if (op === '+') return na + nb;
  if (op === '-') return na - nb;
  if (op === '*') return na * nb;
  if (op === '/') return nb === 0 ? ERR('#DIV/0!') : na / nb;
  if (op === '^') return Math.pow(na, nb);
  return ERR('#VALUE!');
}
function looseEq(a, b){
  if (a === '' || a === null || a === undefined) return b === '' || b === null || b === undefined;
  if (b === '' || b === null || b === undefined) return false;
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  const sa = String(a), sb = String(b);
  const na = Number(sa), nb = Number(sb);
  if (sa.trim() !== '' && sb.trim() !== '' && !isNaN(na) && !isNaN(nb)) return na === nb;
  const da = toSerial(sa), db = toSerial(sb);
  if (da !== null && db !== null) return da === db;
  return sa.toLowerCase() === sb.toLowerCase();
}
function evalNode(node, grid){
  switch (node.type){
    case 'num': return node.v;
    case 'str': return node.v;
    case 'bool': return node.v;
    case 'neg': { const v = evalNode(node.v, grid); if (isErr(v)) return v; const n = toNum(v); return isErr(n) ? n : -n; }
    case 'bin': {
      const a = evalNode(node.left, grid);
      if (isErr(a)) return a;
      const b = evalNode(node.right, grid);
      if (isErr(b)) return b;
      return applyBin(node.op, a, b);
    }
    case 'cell': return cellValue(grid, node.col, node.row);
    case 'range': return rangeValue(grid, node.c1, node.r1, node.c2, node.r2);
    case 'colrange': {
      const lastRow = grid.rows.length + 1;
      return rangeValue(grid, node.c1, 2, node.c2, lastRow);
    }
    case 'rowrange': {
      const out = [];
      for (let c = 0; c < grid.headers.length; c++) out.push(cellValue(grid, c, node.r1));
      return [out];
    }
    case 'func': return evalFunc(node, grid);
    default: return ERR('#VALUE!');
  }
}
const LAZY = new Set(['IF','IFS','IFERROR','IFNA','SWITCH']);
function evalFunc(node, grid){
  const name = node.name;
  if (LAZY.has(name)) return evalLazy(name, node.args, grid);
  const args = node.args.map(a => evalNode(a, grid));
  for (const a of args) if (isErr(a)) return a;
  const fn = FUNCS[name];
  if (!fn) return ERR('#NAME?');
  try { return fn(args, grid); }
  catch(e){ return ERR('#VALUE!'); }
}
function evalLazy(name, argNodes, grid){
  const ev = n => evalNode(n, grid);
  try {
    if (name === 'IF'){
      const c = ev(argNodes[0]); if (isErr(c)) return c;
      return truthy(c) ? ev(argNodes[1]) : (argNodes[2] ? ev(argNodes[2]) : false);
    }
    if (name === 'IFS'){
      for (let i = 0; i + 1 < argNodes.length; i += 2){
        const c = ev(argNodes[i]); if (isErr(c)) return c;
        if (truthy(c)) return ev(argNodes[i+1]);
      }
      return ERR('#N/A');
    }
    if (name === 'IFERROR'){
      const v = ev(argNodes[0]);
      return isErr(v) ? ev(argNodes[1]) : v;
    }
    if (name === 'IFNA'){
      const v = ev(argNodes[0]);
      return (isErr(v) && v.err === '#N/A') ? ev(argNodes[1]) : v;
    }
    if (name === 'SWITCH'){
      const v = ev(argNodes[0]); if (isErr(v)) return v;
      for (let i = 1; i + 1 < argNodes.length; i += 2){
        const c = ev(argNodes[i]); if (isErr(c)) return c;
        if (looseEq(v, c)) return ev(argNodes[i+1]);
      }
      const hasDefault = argNodes.length % 2 === 0;
      return hasDefault ? ev(argNodes[argNodes.length-1]) : ERR('#N/A');
    }
  } catch(e){ return ERR('#VALUE!'); }
  return ERR('#VALUE!');
}
function truthy(v){
  if (isErr(v)) return false;
  if (Array.isArray(v)){
    const f = flat(v);
    return f.length ? truthy(f[0]) : false;
  }
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') return v.trim().toLowerCase() === 'true' ? true : v.trim() !== '';
  return !!v;
}

/* ---------- function library ---------- */
function numsOnly(vals){
  const out = [];
  for (const v of flat(vals)){
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'string'){
      const d = toSerial(v.trim());
      if (d !== null) out.push(d);
      else { const n = Number(v.trim()); if (v.trim() !== '' && !isNaN(n)) out.push(n); }
    } else if (typeof v === 'boolean') out.push(v ? 1 : 0);
  }
  return out;
}
function wildcardMatch(text, pattern){
  const re = '^' + String(pattern).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
  return new RegExp(re, 'i').test(String(text));
}
function criteriaMatch(cell, crit){
  if (typeof crit === 'string'){
    const s = crit.trim();
    const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(s);
    if (m){
      const op = m[1], rhs = m[2];
      const cn = Number(rhs), vn = toNum(cell);
      if (!isNaN(cn) && rhs.trim() !== '' && !isErr(vn) && typeof cell !== 'string' || (typeof cell === 'number')){
        if (isErr(vn)) return false;
        return op === '<' ? vn < cn : op === '>' ? vn > cn : op === '<=' ? vn <= cn : op === '>=' ? vn >= cn : op === '<>' ? vn !== cn : vn === cn;
      }
      const cmp = String(cell).toLowerCase().localeCompare(rhs.toLowerCase());
      if (op === '=') return String(cell).toLowerCase() === rhs.toLowerCase();
      if (op === '<>') return String(cell).toLowerCase() !== rhs.toLowerCase();
      if (op === '<') return cmp < 0; if (op === '>') return cmp > 0;
      if (op === '<=') return cmp <= 0; if (op === '>=') return cmp >= 0;
      return false;
    }
    if (s.includes('*') || s.includes('?')) return wildcardMatch(cell ?? '', s);
    return looseEq(cell, s);
  }
  return looseEq(cell, crit);
}
function ifsPairs(args, start){
  // returns list of [rangeValues(flat), criteria]
  const pairs = [];
  for (let i = start; i + 1 < args.length; i += 2) pairs.push([flat(args[i]), args[i+1]]);
  return pairs;
}
function ifsMatch(pairs, idx){
  return pairs.every(([vals, crit]) => criteriaMatch(vals[idx] ?? '', crit));
}
function fmtDateSerial(n, fmt){
  const d = new Date((n + EPOCH) * 86400000);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const M = d.getUTCMonth(), Y = d.getUTCFullYear(), D = d.getUTCDate();
  const pad = x => String(x).padStart(2,'0');
  switch ((fmt||'').toLowerCase()){
    case 'mmm-yyyy': return months[M] + '-' + Y;
    case 'mmm-yy': return months[M] + '-' + String(Y).slice(2);
    case 'dd-mmm-yyyy': return pad(D) + '-' + months[M] + '-' + Y;
    case 'dd/mm/yyyy': return pad(D) + '/' + pad(M+1) + '/' + Y;
    case 'mm/dd/yyyy': return pad(M+1) + '/' + pad(D) + '/' + Y;
    case 'yyyy': return String(Y);
    case 'mmm': return months[M];
    case 'dd': return pad(D);
    default: return fromSerial(n);
  }
}
function dateParts(serial){
  const d = new Date((serial + EPOCH) * 86400000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), dow: d.getUTCDay() };
}
function serialFromParts(y, m, d){
  return Math.round(Date.UTC(y, m - 1, d) / 86400000 - EPOCH);
}

const FUNCS = {
  SUM: a => numsOnly(a).reduce((s, x) => s + x, 0),
  AVERAGE: a => { const n = numsOnly(a); return n.length ? n.reduce((s,x)=>s+x,0)/n.length : ERR('#DIV/0!'); },
  COUNT: a => numsOnly(a).length,
  COUNTA: a => flat(a).filter(v => v !== '' && v !== null && v !== undefined).length,
  COUNTBLANK: a => flat(a).filter(v => v === '' || v === null || v === undefined).length,
  MAX: a => { const n = numsOnly(a); return n.length ? Math.max(...n) : 0; },
  MIN: a => { const n = numsOnly(a); return n.length ? Math.min(...n) : 0; },
  LARGE: a => { const n = numsOnly(a[0]).sort((x,y)=>y-x); const k = Math.floor(toNum(a[1])); return (k>=1 && k<=n.length) ? n[k-1] : ERR('#NUM!'); },
  SMALL: a => { const n = numsOnly(a[0]).sort((x,y)=>x-y); const k = Math.floor(toNum(a[1])); return (k>=1 && k<=n.length) ? n[k-1] : ERR('#NUM!'); },
  RANK: a => {
    const v = toNum(a[0]); const list = numsOnly(a[1]).sort((x,y)=>y-x);
    const order = a.length > 2 ? toNum(a[2]) : 0;
    const sorted = order === 1 ? [...list].sort((x,y)=>x-y) : list;
    const idx = sorted.indexOf(v);
    return idx === -1 ? ERR('#N/A') : idx + 1;
  },
  PERCENTILE: a => {
    const n = numsOnly(a[0]).sort((x,y)=>x-y);
    const k = toNum(a[1]);
    if (!n.length || k < 0 || k > 1) return ERR('#NUM!');
    const pos = k * (n.length - 1), lo = Math.floor(pos), hi = Math.ceil(pos);
    return n[lo] + (n[hi] - n[lo]) * (pos - lo);
  },
  'PERCENTILE.INC': a => FUNCS.PERCENTILE(a),
  SUMPRODUCT: a => {
    const lists = a.map(x => flat(x).map(v => { const n = toNum(v); return isErr(n) ? 0 : n; }));
    const len = Math.min(...lists.map(l => l.length));
    let s = 0;
    for (let i = 0; i < len; i++) s += lists.reduce((p, l) => p * l[i], 1);
    return s;
  },
  ABS: a => { const n = toNum(a[0]); return isErr(n) ? n : Math.abs(n); },
  ROUND: a => { const n = toNum(a[0]); const d = a.length > 1 ? Math.floor(toNum(a[1])) : 0; return isErr(n) ? n : Number(n.toFixed(d)); },
  INT: a => { const n = toNum(a[0]); return isErr(n) ? n : Math.floor(n); },
  MOD: a => { const x = toNum(a[0]), y = toNum(a[1]); return (isErr(x)||isErr(y)) ? ERR('#VALUE!') : (y===0 ? ERR('#DIV/0!') : x - y*Math.floor(x/y)); },
  AND: a => { for (const v of flat(a)) if (!truthy(v)) return false; return true; },
  OR: a => { for (const v of flat(a)) if (truthy(v)) return true; return false; },
  NOT: a => !truthy(a[0]),
  TRUE: () => true,
  FALSE: () => false,
  SUMIF: a => {
    const range = flat(a[0]), sumRange = a.length > 2 ? flat(a[2]) : range;
    let s = 0;
    range.forEach((cell, i) => { if (criteriaMatch(cell, a[1])) { const n = toNum(sumRange[i] ?? ''); if (!isErr(n)) s += n; } });
    return s;
  },
  COUNTIF: a => flat(a[0]).filter(cell => criteriaMatch(cell, a[1])).length,
  AVERAGEIF: a => {
    const range = flat(a[0]), avgRange = a.length > 2 ? flat(a[2]) : range;
    const vals = [];
    range.forEach((cell, i) => { if (criteriaMatch(cell, a[1])) { const n = toNum(avgRange[i] ?? ''); if (!isErr(n)) vals.push(n); } });
    return vals.length ? vals.reduce((s,x)=>s+x,0)/vals.length : ERR('#DIV/0!');
  },
  SUMIFS: a => {
    const sumRange = flat(a[0]); const pairs = ifsPairs(a, 1);
    let s = 0;
    sumRange.forEach((_, i) => { if (ifsMatch(pairs, i)) { const n = toNum(sumRange[i]); if (!isErr(n)) s += n; } });
    return s;
  },
  COUNTIFS: a => {
    const pairs = ifsPairs(a, 0);
    const len = pairs.length ? pairs[0][0].length : 0;
    let c = 0;
    for (let i = 0; i < len; i++) if (ifsMatch(pairs, i)) c++;
    return c;
  },
  AVERAGEIFS: a => {
    const avgRange = flat(a[0]); const pairs = ifsPairs(a, 1);
    const vals = [];
    avgRange.forEach((_, i) => { if (ifsMatch(pairs, i)) { const n = toNum(avgRange[i]); if (!isErr(n)) vals.push(n); } });
    return vals.length ? vals.reduce((s,x)=>s+x,0)/vals.length : ERR('#DIV/0!');
  },
  XLOOKUP: a => {
    const lookup = a[0];
    const larr = flat(a[1]), rarr = flat(a[2]);
    const notFound = a.length > 3 ? a[3] : ERR('#N/A');
    const searchMode = a.length > 5 ? toNum(a[5]) : 1;
    const idxs = searchMode === -1 ? [...larr.keys()].reverse() : [...larr.keys()];
    for (const i of idxs) if (looseEq(larr[i], lookup)) return rarr[i] ?? '';
    return notFound;
  },
  VLOOKUP: a => {
    const lookup = a[0], table = a[1], col = Math.floor(toNum(a[2]));
    const exact = a.length > 3 ? (toNum(a[3]) === 0 || String(a[3]).toUpperCase() === 'FALSE') : true;
    if (!Array.isArray(table) || !table.length) return ERR('#N/A');
    if (col < 1 || col > table[0].length) return ERR('#REF!');
    if (exact){
      for (const row of table) if (looseEq(row[0], lookup)) return row[col-1] ?? '';
      return ERR('#N/A');
    }
    let best = null;
    for (const row of table){
      const c = toNum(row[0]), l = toNum(lookup);
      if (!isErr(c) && !isErr(l) && c <= l) best = row;
      else break;
    }
    return best ? (best[col-1] ?? '') : ERR('#N/A');
  },
  INDEX: a => {
    const range = a[0];
    if (!Array.isArray(range)) return range;
    const r = a.length > 1 ? Math.floor(toNum(a[1])) : 1;
    const c = a.length > 2 ? Math.floor(toNum(a[2])) : 1;
    if (r === 0 && c === 0) return range;
    if (r === 0) return range.map(row => [row[c-1] ?? '']);
    if (c === 0) return [range[r-1] ?? []];
    const row = range[r-1];
    return row ? (row[c-1] ?? '') : ERR('#REF!');
  },
  MATCH: a => {
    const lookup = a[0], arr = flat(a[1]);
    const type = a.length > 2 ? Math.floor(toNum(a[2])) : 1;
    if (type === 0){
      for (let i = 0; i < arr.length; i++) if (looseEq(arr[i], lookup)) return i + 1;
      return ERR('#N/A');
    }
    if (type === 1){
      let best = -1;
      for (let i = 0; i < arr.length; i++){
        const c = toNum(arr[i]), l = toNum(lookup);
        if (!isErr(c) && !isErr(l) && c <= l) best = i; else if (!isErr(c) && !isErr(l)) break;
      }
      return best === -1 ? ERR('#N/A') : best + 1;
    }
    let best = -1;
    for (let i = 0; i < arr.length; i++){
      const c = toNum(arr[i]), l = toNum(lookup);
      if (!isErr(c) && !isErr(l) && c >= l) best = i; else if (!isErr(c) && !isErr(l)) break;
    }
    return best === -1 ? ERR('#N/A') : best + 1;
  },
  TEXT: a => {
    const v = a[0], fmt = String(a[1] ?? '');
    if (typeof v === 'number' && /[dmy]/i.test(fmt) && !/%/.test(fmt)) return fmtDateSerial(Math.floor(v), fmt);
    if (typeof v === 'string'){ const s = toSerial(v); if (s !== null && /[dmy]/i.test(fmt)) return fmtDateSerial(s, fmt); }
    const n = toNum(v);
    if (!isErr(n)){
      if (/%/.test(fmt)) return (n * 100).toFixed((fmt.match(/0\.0+/) || [''])[0].length - 1 || 0) + '%';
      const dec = (fmt.match(/0\.(0+)/) || [])[1];
      if (dec) return n.toFixed(dec.length);
      if (/#,##0/.test(fmt)) return Math.round(n).toLocaleString('en-US');
      if (fmt === '0') return String(Math.round(n));
    }
    return toStr(v);
  },
  CONCAT: a => flat(a).map(v => toStr(v)).join(''),
  CONCATENATE: a => flat(a).map(v => toStr(v)).join(''),
  TEXTJOIN: a => {
    const delim = toStr(a[0]), ignore = truthy(a[1]);
    const parts = [];
    for (let i = 2; i < a.length; i++)
      for (const v of flat(a[i])){ const s = toStr(v); if (!ignore || s !== '') parts.push(s); }
    return parts.join(delim);
  },
  LEFT: a => toStr(a[0]).slice(0, a.length > 1 ? Math.floor(toNum(a[1])) : 1),
  RIGHT: a => { const s = toStr(a[0]); const n = a.length > 1 ? Math.floor(toNum(a[1])) : 1; return s.slice(-n); },
  MID: a => { const s = toStr(a[0]); const st = Math.floor(toNum(a[1])) - 1, n = Math.floor(toNum(a[2])); return s.slice(st, st + n); },
  LEN: a => toStr(a[0]).length,
  TRIM: a => toStr(a[0]).replace(/\s+/g, ' ').trim(),
  UPPER: a => toStr(a[0]).toUpperCase(),
  LOWER: a => toStr(a[0]).toLowerCase(),
  SUBSTITUTE: a => {
    const s = toStr(a[0]), old = toStr(a[1]), nw = toStr(a[2]);
    if (a.length > 3){
      const n = Math.floor(toNum(a[3])); let count = 0;
      return s.split(old).map((part, i, arr) => i < arr.length - 1 ? part + (++count === n ? nw : old) : part).join('');
    }
    return s.split(old).join(nw);
  },
  CLEAN: a => toStr(a[0]).replace(/[\x00-\x1F\x7F]/g, ''),
  DATE: a => serialFromParts(Math.floor(toNum(a[0])), Math.floor(toNum(a[1])), Math.floor(toNum(a[2]))),
  DAY: a => { const s = toNum(a[0]); return isErr(s) ? s : dateParts(Math.floor(s)).d; },
  MONTH: a => { const s = toNum(a[0]); return isErr(s) ? s : dateParts(Math.floor(s)).m; },
  YEAR: a => { const s = toNum(a[0]); return isErr(s) ? s : dateParts(Math.floor(s)).y; },
  EOMONTH: a => {
    const s = Math.floor(toNum(a[0])), off = Math.floor(toNum(a[1]));
    if (isErr(s)) return s;
    const p = dateParts(s);
    let y = p.y, m = p.m + off;
    while (m > 12){ m -= 12; y++; } while (m < 1){ m += 12; y--; }
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return serialFromParts(y, m, last);
  },
  EDATE: a => {
    const s = Math.floor(toNum(a[0])), off = Math.floor(toNum(a[1]));
    if (isErr(s)) return s;
    const p = dateParts(s);
    let y = p.y, m = p.m + off;
    while (m > 12){ m -= 12; y++; } while (m < 1){ m += 12; y--; }
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return serialFromParts(y, m, Math.min(p.d, last));
  },
  DATEDIF: a => {
    const s1 = Math.floor(toNum(a[0])), s2 = Math.floor(toNum(a[1]));
    const unit = String(a[2]).toLowerCase();
    if (isErr(s1) || isErr(s2)) return ERR('#VALUE!');
    const p1 = dateParts(s1), p2 = dateParts(s2);
    if (unit === 'd') return s2 - s1;
    if (unit === 'm') return (p2.y - p1.y) * 12 + (p2.m - p1.m) - (p2.d < p1.d ? 1 : 0);
    if (unit === 'y') return p2.y - p1.y - (p2.m < p1.m || (p2.m === p1.m && p2.d < p1.d) ? 1 : 0);
    return ERR('#VALUE!');
  },
  NETWORKDAYS: a => {
    const s1 = Math.floor(toNum(a[0])), s2 = Math.floor(toNum(a[1]));
    if (isErr(s1) || isErr(s2)) return ERR('#VALUE!');
    let c = 0;
    for (let d = Math.min(s1,s2); d <= Math.max(s1,s2); d++){
      const dow = dateParts(d).dow;
      if (dow !== 0 && dow !== 6) c++;
    }
    return s2 >= s1 ? c : -c;
  },
  WORKDAY: a => {
    let d = Math.floor(toNum(a[0])); const n = Math.floor(toNum(a[1]));
    if (isErr(d)) return d;
    const step = n >= 0 ? 1 : -1;
    let left = Math.abs(n);
    while (left > 0){ d += step; const dow = dateParts(d).dow; if (dow !== 0 && dow !== 6) left--; }
    return d;
  },
  TODAY: () => { const t = new Date(); return Math.round(Date.UTC(t.getFullYear(), t.getMonth(), t.getDate())/86400000 - EPOCH); },
  NOW: () => Date.now()/86400000 - EPOCH,
  UNIQUE: a => {
    const seen = new Set(), out = [];
    for (const v of flat(a[0])){ const k = String(v).toLowerCase(); if (!seen.has(k)){ seen.add(k); out.push([v]); } }
    return out;
  },
  SORT: a => {
    const rows = Array.isArray(a[0]) ? a[0].map(r => [...r]) : [[a[0]]];
    const idx = a.length > 1 ? Math.floor(toNum(a[1])) - 1 : 0;
    const order = a.length > 2 ? toNum(a[2]) : 1;
    rows.sort((x, y) => {
      const vx = x[idx] ?? '', vy = y[idx] ?? '';
      const nx = toNum(vx), ny = toNum(vy);
      const cmp = (!isErr(nx) && !isErr(ny) && typeof vx !== 'string' && typeof vy !== 'string') || (typeof vx === 'number' && typeof vy === 'number')
        ? nx - ny : String(vx).toLowerCase().localeCompare(String(vy).toLowerCase());
      return cmp * (order >= 0 ? 1 : -1);
    });
    return rows;
  },
  SORTBY: a => {
    const rows = Array.isArray(a[0]) ? a[0].map(r => [...r]) : [[a[0]]];
    const keys = [];
    for (let i = 1; i + 1 < a.length; i += 2) keys.push({ vals: flat(a[i]), order: toNum(a[i+1]) || 1 });
    const idx = rows.map((_, i) => i);
    idx.sort((p, q) => {
      for (const k of keys){
        const vx = k.vals[p] ?? '', vy = k.vals[q] ?? '';
        const nx = toNum(vx), ny = toNum(vy);
        const cmp = (typeof vx === 'number' && typeof vy === 'number') ? nx - ny
          : String(vx).toLowerCase().localeCompare(String(vy).toLowerCase());
        if (cmp !== 0) return cmp * (k.order >= 0 ? 1 : -1);
      }
      return 0;
    });
    return idx.map(i => rows[i]);
  },
  FILTER: a => {
    const rows = Array.isArray(a[0]) ? a[0] : [[a[0]]];
    const inc = flat(a[1]);
    const out = rows.filter((_, i) => truthy(inc[i]));
    if (!out.length) return a.length > 2 ? a[2] : ERR('#CALC!');
    return out;
  },
};
