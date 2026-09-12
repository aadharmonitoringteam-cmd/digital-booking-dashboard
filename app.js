let allRows=[], filteredRows=[], page=1, charts={};
const $=id=>document.getElementById(id);
const countCols=[], amountCols=[];
let headers=[];

function num(v){const n=Number(v);return Number.isFinite(n)?n:0}
function fmtNum(n){return new Intl.NumberFormat("en-IN",{maximumFractionDigits:0}).format(n)}
function fmtMoney(n){return "₹"+new Intl.NumberFormat("en-IN",{maximumFractionDigits:2}).format(n)}
function cleanName(v){return v==null?"":String(v).trim()}

function classifyColumns(){
  countCols.length=0; amountCols.length=0;
  headers.forEach((h,i)=>{if(String(h||"").endsWith("(Cnt)"))countCols.push(i);if(String(h||"").endsWith("(Amt)"))amountCols.push(i)});
}
function rowCount(r){return countCols.reduce((s,i)=>s+num(r[i]),0)}
function rowAmount(r){return amountCols.reduce((s,i)=>s+num(r[i]),0)}

async function readWorkbook(arrayBuffer, sourceName){
  const wb=XLSX.read(arrayBuffer,{type:"array"});
  const ws=wb.Sheets["Daily Booking Report"] || wb.Sheets[wb.SheetNames[0]];
  if(!ws) throw new Error("Daily Booking Report sheet not found.");
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:0});
  if(!raw.length) throw new Error("The sheet is empty.");
  headers=raw[0];
  classifyColumns();
  allRows=raw.slice(1).filter(r=>r.length && r[0]!==null && r[0]!==undefined && r[0]!=="");
  $("sourceStatus").textContent=`Source: ${sourceName} • ${allRows.length.toLocaleString("en-IN")} office rows`;
  buildDivisionOptions();
  applyFilters();
}
function buildDivisionOptions(){
  const divs=[...new Set(allRows.map(r=>cleanName(r[headers.indexOf("Division")])).filter(Boolean))].sort();
  $("divisionFilter").innerHTML='<option value="">All Divisions</option>'+divs.map(d=>`<option>${esc(d)}</option>`).join("");
}
function applyFilters(){
  const divIdx=headers.indexOf("Division"), officeIdx=headers.indexOf("Office Name"), idIdx=headers.indexOf("Office ID");
  const div=$("divisionFilter").value, q=$("officeFilter").value.toLowerCase().trim();
  filteredRows=allRows.filter(r=>(!div||cleanName(r[divIdx])===div)&&(!q||cleanName(r[officeIdx]).toLowerCase().includes(q)||cleanName(r[idIdx]).toLowerCase().includes(q)));
  page=1; render();
}
function render(){
  const divIdx=headers.indexOf("Division"), officeIdx=headers.indexOf("Office Name");
  const totalCnt=filteredRows.reduce((s,r)=>s+rowCount(r),0), totalAmt=filteredRows.reduce((s,r)=>s+rowAmount(r),0);
  $("totalCnt").textContent=fmtNum(totalCnt); $("totalAmt").textContent=fmtMoney(totalAmt);
  $("officeCnt").textContent=fmtNum(filteredRows.length);
  $("divisionCnt").textContent=fmtNum(new Set(filteredRows.map(r=>cleanName(r[divIdx])).filter(Boolean)).size);
  renderCharts(divIdx,officeIdx); renderTable(divIdx,officeIdx);
}
function aggregateBy(keyIdx, rows=filteredRows){
  const m=new Map();
  rows.forEach(r=>{const k=cleanName(r[keyIdx])||"Unknown";const x=m.get(k)||{cnt:0,amt:0};x.cnt+=rowCount(r);x.amt+=rowAmount(r);m.set(k,x)});
  return [...m.entries()].sort((a,b)=>b[1].cnt-a[1].cnt);
}
function destroy(name){if(charts[name]){charts[name].destroy();delete charts[name]}}
function makeChart(name, id, type, labels, data, label){
  destroy(name); charts[name]=new Chart($(id),{type,data:{labels,datasets:[{label,data,borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:type==="doughnut"?{}:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:"#edf0f3"}}}}});
}
function renderCharts(divIdx,officeIdx){
  const div=aggregateBy(divIdx), top=aggregateBy(officeIdx).slice(0,10);
  makeChart("div","divisionChart","bar",div.map(x=>x[0]),div.map(x=>x[1].cnt),"Transactions");
  makeChart("amt","amountChart","bar",div.map(x=>x[0]),div.map(x=>x[1].amt),"Amount");
  // Aggregate all count columns by payment/category, using their header labels.
  const mode= new Map();
  countCols.forEach(i=>{const label=String(headers[i]).replace(/\s*\(Cnt\)$/,""); mode.set(label,filteredRows.reduce((s,r)=>s+num(r[i]),0))});
  const modeSorted=[...mode.entries()].filter(x=>x[1]>0).sort((a,b)=>b[1]-a[1]).slice(0,12);
  makeChart("mode","modeChart","doughnut",modeSorted.map(x=>x[0]),modeSorted.map(x=>x[1]),"Count");
  makeChart("office","officeChart","bar",top.map(x=>x[0]),top.map(x=>x[1].cnt),"Transactions");
}
function renderTable(divIdx,officeIdx){
  const rowsPer=Number($("rowsFilter").value), pages=Math.max(1,Math.ceil(filteredRows.length/rowsPer)); if(page>pages)page=pages;
  const start=(page-1)*rowsPer, slice=filteredRows.slice().sort((a,b)=>rowCount(b)-rowCount(a)).slice(start,start+rowsPer);
  $("tableBody").innerHTML=slice.length?slice.map((r,i)=>`<tr><td>${start+i+1}</td><td>${esc(cleanName(r[officeIdx])||"—")}</td><td>${esc(cleanName(r[divIdx])||"—")}</td><td>${fmtNum(rowCount(r))}</td><td>${fmtMoney(rowAmount(r))}</td></tr>`).join(""):`<tr><td colspan="5" class="empty">No matching offices</td></tr>`;
  $("tableMeta").textContent=`${filteredRows.length.toLocaleString("en-IN")} matching offices`;
  $("pageInfo").textContent=`Page ${page} of ${pages}`;$("prevBtn").disabled=page<=1;$("nextBtn").disabled=page>=pages;
}
function esc(v){return String(v).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function downloadCSV(){
  const divIdx=headers.indexOf("Division"), officeIdx=headers.indexOf("Office Name"), rows=filteredRows.slice().sort((a,b)=>rowCount(b)-rowCount(a));
  const out=[["Office Name","Division","Transactions","Amount"],...rows.map(r=>[cleanName(r[officeIdx]),cleanName(r[divIdx]),rowCount(r),rowAmount(r)])];
  const blob=new Blob([out.map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(",")).join("\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="booking-dashboard.csv";a.click();URL.revokeObjectURL(a.href);
}
async function loadDefault(){
  try{const res=await fetch("data/Digital%20Txn%20Status.xlsx?"+Date.now());if(!res.ok)throw Error("Excel file not found");await readWorkbook(await res.arrayBuffer(),"GitHub / data folder")}
  catch(e){$("sourceStatus").textContent="Could not load default Excel. Use Upload Daily Excel.";console.error(e)}
}
$("fileInput").addEventListener("change",async e=>{const f=e.target.files[0];if(!f)return;try{await readWorkbook(await f.arrayBuffer(),f.name)}catch(err){alert(err.message)}});
$("resetBtn").onclick=loadDefault;
$("divisionFilter").onchange=applyFilters;$("officeFilter").oninput=applyFilters;$("rowsFilter").onchange=()=>{page=1;render()};
$("clearBtn").onclick=()=>{$("divisionFilter").value="";$("officeFilter").value="";applyFilters()};
$("prevBtn").onclick=()=>{if(page>1){page--;renderTable(headers.indexOf("Division"),headers.indexOf("Office Name"))}};
$("nextBtn").onclick=()=>{const n=Math.max(1,Math.ceil(filteredRows.length/Number($("rowsFilter").value)));if(page<n){page++;renderTable(headers.indexOf("Division"),headers.indexOf("Office Name"))}};
$("downloadBtn").onclick=downloadCSV;
loadDefault();
