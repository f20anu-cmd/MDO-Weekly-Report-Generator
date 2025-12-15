/* Complete app.js (FULL FILE)
   - Loads dropdowns from data/master_data.xlsx
   - Buttons always wired even if excel load fails
   - No cursor jump while typing (no rerender on each keystroke)
   - Full sections + photos + next week plan (PLAN ONLY) + activity plan + special achievement

   🔧 SURGICAL FIX APPLIED:
   - Next Week Plan: Actual column removed completely (state, render, add-row template, payload remains plan-only)
*/

const $ = (id) => document.getElementById(id);

const CFG = {
  excelPath: "data/master_data.xlsx",
  maxNpiRows: 9,
  maxOtherRows: 10,
  maxNextWeekRows: 10,
  maxActPhotos: 16,
  maxSpPhotos: 4,
  months: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  weeks: ["1","2","3","4","5"]
};

const Master = {
  regionToTerritories: new Map(),
  npiProducts: [],
  otherProducts: [],
  allProducts: [],
  activityTypes: [],
  // product -> { realised, category }
  productMeta: new Map(),
  // product -> { realised, incentiveRate }
  npiMeta: new Map()
};

const State = {
  npiRows: [],
  otherRows: [],
  activityRows: [],
  photoRows: [],
  nextWeekRows: [],      // 🔧 Next week rows: PLAN ONLY
  actPlanRows: [],
  spDesc: "",
  spPhotoRows: []
};

function rs(n){
  const v = Math.round(Number(n || 0));
  return v.toLocaleString("en-IN");
}
function toNum(v){
  if (v === null || v === undefined) return 0;
  const x = Number(String(v).replace(/,/g,"").trim());
  return Number.isFinite(x) ? x : 0;
}

function setOptions(selectEl, values, placeholder){
  selectEl.innerHTML = "";
  const ph = document.createElement("option");
  ph.value = "";
  ph.textContent = placeholder || "Select";
  selectEl.appendChild(ph);
  for(const v of values){
    const op = document.createElement("option");
    op.value = v;
    op.textContent = v;
    selectEl.appendChild(op);
  }
}

function setStatus(kind, text){
  const el = $("masterStatus");
  if(!el) return;
  el.classList.remove("pill--ok","pill--warn","pill--loading");
  if(kind === "ok") el.classList.add("pill--ok");
  else if(kind === "warn") el.classList.add("pill--warn");
  else el.classList.add("pill--loading");
  el.textContent = text;
}

function showFatal(msg){
  const el = $("fatal");
  if(!el) return;
  el.classList.remove("hidden");
  el.innerHTML = msg;
}

function clearFatal(){
  const el = $("fatal");
  if(!el) return;
  el.classList.add("hidden");
  el.textContent = "";
}

/* ---------- Excel load ---------- */
function safeSheet(wb, ...names){
  for(const n of names){
    if(wb.Sheets[n]) return wb.Sheets[n];
  }
  return null;
}

async function loadMasterExcel(){
  const res = await fetch(CFG.excelPath, { cache: "no-store" });
  if(!res.ok) throw new Error(`Cannot load ${CFG.excelPath}`);

  const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });

  const sRM = safeSheet(wb, "Region Mapping");
  const sNPI = safeSheet(wb, "NPI sheet", "NPI Sheet");
  const sPL  = safeSheet(wb, "Product List");
  const sAL  = safeSheet(wb, "Activity List");

  if(!sRM || !sNPI || !sPL || !sAL){
    throw new Error("Missing one or more required sheets in master_data.xlsx");
  }

  const rm = XLSX.utils.sheet_to_json(sRM, { defval: "" });
  const npi = XLSX.utils.sheet_to_json(sNPI, { defval: "" });
  const pl  = XLSX.utils.sheet_to_json(sPL,  { defval: "" });
  const al  = XLSX.utils.sheet_to_json(sAL,  { defval: "" });

  // Region mapping: Region -> Territtory (spelling preserved)
  const map = new Map();
  for(const r of rm){
    const region = String(r["Region"] || "").trim();
    const terr = String(r["Territtory"] || "").trim();
    if(!region || !terr) continue;
    if(!map.has(region)) map.set(region, []);
    map.get(region).push(terr);
  }
  for(const [k,v] of map){
    map.set(k, [...new Set(v)].sort((a,b)=>a.localeCompare(b)));
  }
  Master.regionToTerritories = map;

  // NPI meta
  const npiMeta = new Map();
  for(const r of npi){
    const product = String(r["Product"] || "").trim();
    if(!product) continue;
    const realised = toNum(r["Realised Value in Rs"]);
    const incentiveRate = toNum(r["Incentive"]);
    npiMeta.set(product, { realised, incentiveRate });
  }
  Master.npiMeta = npiMeta;
  Master.npiProducts = [...npiMeta.keys()].sort((a,b)=>a.localeCompare(b));

  // Product list meta
  const pMeta = new Map();
  for(const r of pl){
    const product = String(r["Product"] || "").trim();
    if(!product) continue;
    const realised = toNum(r["Realised Value"]);
    const category = String(r["Category"] || "").trim();
    pMeta.set(product, { realised, category });
  }
  Master.productMeta = pMeta;

  // Other products: category != NPI and not in NPI sheet
  const other = [];
  for(const [p, meta] of pMeta.entries()){
    const cat = (meta.category || "").toLowerCase();
    if(cat !== "npi" && !Master.npiMeta.has(p)) other.push(p);
  }
  Master.otherProducts = [...new Set(other)].sort((a,b)=>a.localeCompare(b));

  // All products for next week (union)
  const allSet = new Set([...Master.npiProducts, ...[...pMeta.keys()]]);
  Master.allProducts = [...allSet].sort((a,b)=>a.localeCompare(b));

  // Activity types
  const acts = [];
  for(const r of al){
    const a = String(r["Activity Type"] || "").trim();
    if(a) acts.push(a);
  }
  Master.activityTypes = [...new Set(acts)];
}

/* ---------- Init dropdowns ---------- */
function initHeaderDropdowns(){
  setOptions($("month"), CFG.months, "Select Month");
  setOptions($("week"), CFG.weeks, "Select Week");

  const regions = [...Master.regionToTerritories.keys()].sort((a,b)=>a.localeCompare(b));
  setOptions($("region"), regions, "Select Region");
  setOptions($("territory"), [], "Select Territory");

  $("region").addEventListener("change", ()=>{
    const terrs = Master.regionToTerritories.get($("region").value) || [];
    setOptions($("territory"), terrs, "Select Territory");
  });
}

/* ---------- Photos (compression) ---------- */
async function compressImage(file, maxW=1400, quality=0.78){
  const bmp = await createImageBitmap(file);
  const ratio = Math.min(1, maxW / bmp.width);
  const w = Math.round(bmp.width * ratio);
  const h = Math.round(bmp.height * ratio);

  const canvas = document.createElement("canvas");
  canvas.width = w; canvas.height = h;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(bmp, 0, 0, w, h);
  return canvas.toDataURL("image/jpeg", quality);
}

/* ---------- DOM helpers (stable inputs) ---------- */
function makeNumberInput(initial, onInput){
  const el = document.createElement("input");
  el.type = "number";
  el.min = "0";
  el.step = "any";
  el.value = (initial ?? "");
  // input event updates row + computed cells only; no rerender
  el.addEventListener("input", ()=> onInput(el.value));
  return el;
}

function makeTextInput(initial, onInput, placeholder=""){
  const el = document.createElement("input");
  el.type = "text";
  el.placeholder = placeholder;
  el.value = (initial ?? "");
  el.addEventListener("input", ()=> onInput(el.value));
  return el;
}

function makeSelect(options, initial, onChange, placeholder="Select"){
  const el = document.createElement("select");
  setOptions(el, options, placeholder);
  el.value = initial || "";
  el.addEventListener("change", ()=> onChange(el.value));
  return el;
}

function makeDelButton(onClick){
  const b = document.createElement("button");
  b.className = "icon";
  b.type = "button";
  b.textContent = "✕";
  b.addEventListener("click", onClick);
  return b;
}

/* ---------- Calculations (business rules) ----------
   - Total Incentive Opportunity = Plan × Incentive rate
   - Total Incentive Earned      = Actual × Incentive rate
*/
function recalcNpiSummary(){
  let totalOpp = 0;
  let totalEarn = 0;

  for(const r of State.npiRows){
    const meta = Master.npiMeta.get(r.product) || { incentiveRate: 0 };
    const plan = toNum(r.plan);
    const actual = toNum(r.actual);
    r.opportunity = plan * meta.incentiveRate;
    r.earned = actual * meta.incentiveRate;
    totalOpp += r.opportunity;
    totalEarn += r.earned;
  }

  $("npiEarned").textContent = rs(totalEarn);
  $("npiLose").textContent = rs(Math.max(0, totalOpp - totalEarn));
}

function recalcOtherSummary(){
  let totalRev = 0;
  for(const r of State.otherRows){
    const meta = Master.productMeta.get(r.product) || { realised: 0 };
    const actual = toNum(r.actual);
    r.revenue = actual * meta.realised;
    totalRev += r.revenue;
  }
  $("otherRevenue").textContent = rs(totalRev);
}

function recalcActivityTotals(){
  let p=0,a=0,n=0;
  for(const r of State.activityRows){
    p += toNum(r.planNo);
    a += toNum(r.actualNo);
    n += toNum(r.npiNo);
  }
  $("actPlanTotal").textContent = String(p);
  $("actActualTotal").textContent = String(a);
  $("actNpiTotal").textContent = String(n);
}

/* 🔧 Next Week Summary (PLAN ONLY)
   - Revenue  = Plan × realised
   - Incentive opportunity = Plan × incentiveRate
*/
function recalcNextWeekSummary(){
  let totalRev = 0;
  let totalOpp = 0;

  for(const r of State.nextWeekRows){
    const plan = toNum(r.plan);

    const realised =
      Master.productMeta.get(r.product)?.realised ??
      Master.npiMeta.get(r.product)?.realised ??
      0;

    const rate = Master.npiMeta.get(r.product)?.incentiveRate ?? 0;

    r.revenue = plan * realised;
    r.incentive = plan * rate;

    totalRev += r.revenue;
    totalOpp += r.incentive;
  }

  $("nwRevenue").textContent = rs(totalRev);
  $("nwOpp").textContent = rs(totalOpp);
}

/* ---------- Renders (only rerender on add/remove/clear) ---------- */
function renderNpi(){
  const tbody = $("tblNpi").querySelector("tbody");
  tbody.innerHTML = "";

  State.npiRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdProd = document.createElement("td");
    const sel = makeSelect(Master.npiProducts, row.product, (v)=>{
      row.product = v;
      recalcNpiSummary();
      oppCell.textContent = rs(row.opportunity || 0);
      earnCell.textContent = rs(row.earned || 0);
    }, "Select product");
    tdProd.appendChild(sel);
    tr.appendChild(tdProd);

    const tdPlan = document.createElement("td");
    tdPlan.className = "num";
    tdPlan.appendChild(makeNumberInput(row.plan, (v)=>{
      row.plan = v;
      recalcNpiSummary();
      oppCell.textContent = rs(row.opportunity || 0);
    }));
    tr.appendChild(tdPlan);

    const tdActual = document.createElement("td");
    tdActual.className = "num";
    tdActual.appendChild(makeNumberInput(row.actual, (v)=>{
      row.actual = v;
      recalcNpiSummary();
      earnCell.textContent = rs(row.earned || 0);
    }));
    tr.appendChild(tdActual);

    const tdOpp = document.createElement("td");
    tdOpp.className = "num";
    const oppCell = document.createElement("span");
    oppCell.textContent = rs(row.opportunity || 0);
    tdOpp.appendChild(oppCell);
    tr.appendChild(tdOpp);

    const tdEarn = document.createElement("td");
    tdEarn.className = "num";
    const earnCell = document.createElement("span");
    earnCell.textContent = rs(row.earned || 0);
    tdEarn.appendChild(earnCell);
    tr.appendChild(tdEarn);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.npiRows.splice(idx,1);
      renderNpi();
      recalcNpiSummary();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcNpiSummary();
}

function renderOther(){
  const tbody = $("tblOther").querySelector("tbody");
  tbody.innerHTML = "";

  State.otherRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdProd = document.createElement("td");
    const sel = makeSelect(Master.otherProducts, row.product, (v)=>{
      row.product = v;
      recalcOtherSummary();
      revenueCell.textContent = rs(row.revenue || 0);
    }, "Select product");
    tdProd.appendChild(sel);
    tr.appendChild(tdProd);

    const tdPlan = document.createElement("td");
    tdPlan.className = "num";
    tdPlan.appendChild(makeNumberInput(row.plan, (v)=>{ row.plan = v; }));
    tr.appendChild(tdPlan);

    const tdActual = document.createElement("td");
    tdActual.className = "num";
    tdActual.appendChild(makeNumberInput(row.actual, (v)=>{
      row.actual = v;
      recalcOtherSummary();
      revenueCell.textContent = rs(row.revenue || 0);
    }));
    tr.appendChild(tdActual);

    const tdRev = document.createElement("td");
    tdRev.className = "num";
    const revenueCell = document.createElement("span");
    revenueCell.textContent = rs(row.revenue || 0);
    tdRev.appendChild(revenueCell);
    tr.appendChild(tdRev);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.otherRows.splice(idx,1);
      renderOther();
      recalcOtherSummary();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcOtherSummary();
}

function renderActivities(){
  const tbody = $("tblActivities").querySelector("tbody");
  tbody.innerHTML = "";

  State.activityRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdAct = document.createElement("td");
    tdAct.appendChild(makeSelect(Master.activityTypes, row.activity, (v)=>{ row.activity = v; }, "Select activity"));
    tr.appendChild(tdAct);

    const tdP = document.createElement("td"); tdP.className="num";
    tdP.appendChild(makeNumberInput(row.planNo, (v)=>{ row.planNo=v; recalcActivityTotals(); }));
    tr.appendChild(tdP);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(makeNumberInput(row.actualNo, (v)=>{ row.actualNo=v; recalcActivityTotals(); }));
    tr.appendChild(tdA);

    const tdN = document.createElement("td"); tdN.className="num";
    tdN.appendChild(makeNumberInput(row.npiNo, (v)=>{ row.npiNo=v; recalcActivityTotals(); }));
    tr.appendChild(tdN);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.activityRows.splice(idx,1);
      renderActivities();
      recalcActivityTotals();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcActivityTotals();
}

function renderPhotos(){
  const tbody = $("tblPhotos").querySelector("tbody");
  tbody.innerHTML = "";

  State.photoRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdAct = document.createElement("td");
    tdAct.appendChild(makeSelect(Master.activityTypes, row.activity, (v)=>{ row.activity=v; renderPhotoPreview(); }, "Select activity"));
    tr.appendChild(tdAct);

    const tdUp = document.createElement("td");
    const inp = document.createElement("input");
    inp.type = "file";
    inp.accept = "image/*";
    inp.addEventListener("change", async (e)=>{
      const file = e.target.files?.[0];
      if(!file) return;
      row.fileName = file.name;
      row.dataUrl = await compressImage(file);
      renderPhotoPreview();
    });
    tdUp.appendChild(inp);
    tr.appendChild(tdUp);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.photoRows.splice(idx,1);
      renderPhotos();
      renderPhotoPreview();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  renderPhotoPreview();
}

function renderPhotoPreview(){
  const grid = $("photoPreview");
  grid.innerHTML = "";
  State.photoRows
    .filter(p=>p.dataUrl)
    .slice(0, CFG.maxActPhotos)
    .forEach((p)=>{
      const card = document.createElement("div");
      card.className = "photoCard";
      card.innerHTML = `<img alt="photo"><div class="photoMeta"></div>`;
      card.querySelector("img").src = p.dataUrl;
      card.querySelector(".photoMeta").textContent = p.activity || "Activity";
      grid.appendChild(card);
    });
}

/* 🔧 Next Week render (PLAN ONLY): no Actual input, no row.actual */
function renderNextWeek(){
  const tbody = $("tblNextWeek").querySelector("tbody");
  tbody.innerHTML = "";

  State.nextWeekRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdProd = document.createElement("td");
    tdProd.appendChild(makeSelect(Master.allProducts, row.product, (v)=>{
      row.product = v;
      recalcNextWeekSummary();
      revCell.textContent = rs(row.revenue || 0);
      incCell.textContent = rs(row.incentive || 0);
    }, "Select product"));
    tr.appendChild(tdProd);

    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(makeNumberInput(row.plan, (v)=>{
      row.plan = v;
      recalcNextWeekSummary();
      revCell.textContent = rs(row.revenue || 0);
      incCell.textContent = rs(row.incentive || 0);
    }));
    tr.appendChild(tdPlan);

    // ❌ Actual column removed completely

    const tdRev = document.createElement("td"); tdRev.className="num";
    const revCell = document.createElement("span");
    revCell.textContent = rs(row.revenue || 0);
    tdRev.appendChild(revCell);
    tr.appendChild(tdRev);

    const tdInc = document.createElement("td"); tdInc.className="num";
    const incCell = document.createElement("span");
    incCell.textContent = rs(row.incentive || 0);
    tdInc.appendChild(incCell);
    tr.appendChild(tdInc);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.nextWeekRows.splice(idx,1);
      renderNextWeek();
      recalcNextWeekSummary();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcNextWeekSummary();
}

function renderActPlan(){
  const tbody = $("tblActPlan").querySelector("tbody");
  tbody.innerHTML = "";

  State.actPlanRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdAct = document.createElement("td");
    tdAct.appendChild(makeSelect(Master.activityTypes, row.activity, (v)=>{ row.activity=v; }, "Select activity"));
    tr.appendChild(tdAct);

    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(makeNumberInput(row.planNo, (v)=>{ row.planNo=v; }));
    tr.appendChild(tdPlan);

    const tdVill = document.createElement("td");
    tdVill.appendChild(makeTextInput(row.villages, (v)=>{
      row.villages = v;
      const count = v.split(",").map(s=>s.trim()).filter(Boolean).length;
      row.villageNo = count;
      villageNoCell.textContent = String(count);
    }, "Village1, Village2"));
    tr.appendChild(tdVill);

    const tdCount = document.createElement("td"); tdCount.className="num";
    const villageNoCell = document.createElement("span");
    villageNoCell.textContent = String(row.villageNo || 0);
    tdCount.appendChild(villageNoCell);
    tr.appendChild(tdCount);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.actPlanRows.splice(idx,1);
      renderActPlan();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });
}

function renderSpPhotos(){
  const tbody = $("tblSpPhotos").querySelector("tbody");
  tbody.innerHTML = "";

  State.spPhotoRows.forEach((row, idx)=>{
    const tr = document.createElement("tr");

    const tdIdx = document.createElement("td");
    tdIdx.textContent = String(idx+1);
    tr.appendChild(tdIdx);

    const tdAct = document.createElement("td");
    tdAct.appendChild(makeSelect(Master.activityTypes, row.activity, (v)=>{ row.activity=v; renderSpPhotoPreview(); }, "Select activity"));
    tr.appendChild(tdAct);

    const tdUp = document.createElement("td");
    const inp = document.createElement("input");
    inp.type = "file";
    inp.accept = "image/*";
    inp.addEventListener("change", async (e)=>{
      const file = e.target.files?.[0];
      if(!file) return;
      row.fileName = file.name;
      row.dataUrl = await compressImage(file);
      renderSpPhotoPreview();
    });
    tdUp.appendChild(inp);
    tr.appendChild(tdUp);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelButton(()=>{
      State.spPhotoRows.splice(idx,1);
      renderSpPhotos();
      renderSpPhotoPreview();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  renderSpPhotoPreview();
}

function renderSpPhotoPreview(){
  const grid = $("spPhotoPreview");
  grid.innerHTML = "";
  State.spPhotoRows
    .filter(p=>p.dataUrl)
    .slice(0, CFG.maxSpPhotos)
    .forEach((p)=>{
      const card = document.createElement("div");
      card.className = "photoCard";
      card.innerHTML = `<img alt="photo"><div class="photoMeta"></div>`;
      card.querySelector("img").src = p.dataUrl;
      card.querySelector(".photoMeta").textContent = p.activity || "Special";
      grid.appendChild(card);
    });
}

/* ---------- Buttons ---------- */
function wireButtons(){
  // Clear all
  $("btnClearAll").addEventListener("click", ()=>{
    // form
    $("mdoName").value = "";
    $("hq").value = "";
    $("region").value = "";
    setOptions($("territory"), [], "Select Territory");
    $("month").value = "";
    $("week").value = "";

    // state
    State.npiRows = [];
    State.otherRows = [];
    State.activityRows = [];
    State.photoRows = [];
    State.nextWeekRows = [];     // plan-only
    State.actPlanRows = [];
    State.spDesc = "";
    State.spPhotoRows = [];

    $("spDesc").value = "";

    renderNpi();
    renderOther();
    renderActivities();
    renderPhotos();
    renderNextWeek();
    renderActPlan();
    renderSpPhotos();

    clearFatal();
    window.scrollTo({top:0, behavior:"smooth"});
  });

  // NPI
  $("btnNpiAdd").addEventListener("click", ()=>{
    if(State.npiRows.length >= CFG.maxNpiRows) return alert(`Max ${CFG.maxNpiRows} rows allowed.`);
    State.npiRows.push({ product:"", plan:"", actual:"", opportunity:0, earned:0 });
    renderNpi();
  });
  $("btnNpiClear").addEventListener("click", ()=>{
    State.npiRows = [];
    renderNpi();
  });

  // Other
  $("btnOtherAdd").addEventListener("click", ()=>{
    if(State.otherRows.length >= CFG.maxOtherRows) return alert(`Max ${CFG.maxOtherRows} rows allowed.`);
    State.otherRows.push({ product:"", plan:"", actual:"", revenue:0 });
    renderOther();
  });
  $("btnOtherClear").addEventListener("click", ()=>{
    State.otherRows = [];
    renderOther();
  });

  // Activities update
  $("btnActAdd").addEventListener("click", ()=>{
    State.activityRows.push({ activity:"", planNo:"", actualNo:"", npiNo:"" });
    renderActivities();
  });
  $("btnActClear").addEventListener("click", ()=>{
    State.activityRows = [];
    renderActivities();
  });

  // Photos
  $("btnPhotoAdd").addEventListener("click", ()=>{
    if(State.photoRows.length >= CFG.maxActPhotos) return alert(`Max ${CFG.maxActPhotos} photos allowed.`);
    State.photoRows.push({ activity:"", dataUrl:"", fileName:"" });
    renderPhotos();
  });
  $("btnPhotoClear").addEventListener("click", ()=>{
    State.photoRows = [];
    renderPhotos();
  });

  // Next week (PLAN ONLY)
  $("btnNwAdd").addEventListener("click", ()=>{
    if(State.nextWeekRows.length >= CFG.maxNextWeekRows) return alert(`Max ${CFG.maxNextWeekRows} rows allowed.`);
    // 🔧 no "actual" field
    State.nextWeekRows.push({ product:"", plan:"", revenue:0, incentive:0 });
    renderNextWeek();
  });
  $("btnNwClear").addEventListener("click", ()=>{
    State.nextWeekRows = [];
    renderNextWeek();
  });

  // Activity plan
  $("btnApAdd").addEventListener("click", ()=>{
    State.actPlanRows.push({ activity:"", planNo:"", villages:"", villageNo:0 });
    renderActPlan();
  });
  $("btnApClear").addEventListener("click", ()=>{
    State.actPlanRows = [];
    renderActPlan();
  });

  // Special
  $("spDesc").addEventListener("input", ()=>{ State.spDesc = $("spDesc").value || ""; });
  $("btnSpClear").addEventListener("click", ()=>{
    State.spDesc = "";
    $("spDesc").value = "";
    State.spPhotoRows = [];
    renderSpPhotos();
  });

  // Special photos
  $("btnSpPhotoAdd").addEventListener("click", ()=>{
    if(State.spPhotoRows.length >= CFG.maxSpPhotos) return alert(`Max ${CFG.maxSpPhotos} photos allowed.`);
    State.spPhotoRows.push({ activity:"", dataUrl:"", fileName:"" });
    renderSpPhotos();
  });
  $("btnSpPhotoClear").addEventListener("click", ()=>{
    State.spPhotoRows = [];
    renderSpPhotos();
  });

  // PDF
  $("btnPdf").addEventListener("click", ()=>{
    const payload = {
      masterLoaded: Master.npiProducts.length > 0 || Master.otherProducts.length > 0,
      mdo: {
        name: $("mdoName").value || "",
        hq: $("hq").value || "",
        region: $("region").value || "",
        territory: $("territory").value || "",
        month: $("month").value || "",
        week: $("week").value ? `Week ${$("week").value}` : ""
      },
      npiRows: State.npiRows,
      otherRows: State.otherRows,
      activityRows: State.activityRows,
      photos: State.photoRows.filter(p=>p.dataUrl).slice(0, CFG.maxActPhotos),
      nextWeekRows: State.nextWeekRows, // 🔧 plan-only rows
      actPlanRows: State.actPlanRows,
      spDesc: $("spDesc").value || "",
      spPhotos: State.spPhotoRows.filter(p=>p.dataUrl).slice(0, CFG.maxSpPhotos)
    };

    // Ensure latest summaries
    recalcNpiSummary();
    recalcOtherSummary();
    recalcActivityTotals();
    recalcNextWeekSummary();

    window.generatePerformancePdf(payload, Master, rs);
  });
}

/* ---------- Boot ---------- */
async function boot(){
  wireButtons(); // ✅ ALWAYS wire buttons first

  // Init dropdowns with empty placeholders (safe even if master fails)
  setOptions($("region"), [], "Select Region");
  setOptions($("territory"), [], "Select Territory");
  setOptions($("month"), CFG.months, "Select Month");
  setOptions($("week"), CFG.weeks, "Select Week");

  // Render empty tables
  renderNpi();
  renderOther();
  renderActivities();
  renderPhotos();
  renderNextWeek();
  renderActPlan();
  renderSpPhotos();

  clearFatal();
  setStatus("loading", "Loading master data…");

  try{
    await loadMasterExcel();
    initHeaderDropdowns();

    setStatus("ok", "Master data loaded");
    clearFatal();
  }catch(err){
    console.error(err);
    setStatus("warn", "Master data not loaded");
    showFatal(`Admin Excel not loaded. Check <b>${CFG.excelPath}</b> (case-sensitive on GitHub Pages). Buttons & PDF still work.`);
  }

  // Re-render now that dropdown data exists (if loaded)
  renderNpi();
  renderOther();
  renderActivities();
  renderPhotos();
  renderNextWeek();
  renderActPlan();
  renderSpPhotos();
}

document.addEventListener("DOMContentLoaded", boot);
