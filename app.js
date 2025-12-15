/* ===========================
   Performance Report - app.js
   - Dropdowns from data/master_data.xlsx (admin excel)
   - All buttons wired after DOM ready
   - Cursor stable: commit on BLUR (no render on every keystroke)
   - No server storage
=========================== */

const $ = (id) => document.getElementById(id);

const CFG = {
  excelPath: "data/master_data.xlsx",
  months: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  weeks: ["1","2","3","4","5"],
  maxNpiRows: 9,
  maxOtherRows: 10,
  maxNextWeekRows: 10,
  maxActivityPhotos: 16,
  maxSpecialPhotos: 4
};

const Master = {
  regionToTerritories: new Map(),
  npiProducts: [],
  npiMeta: new Map(),        // product -> { realised, incentiveRate }
  otherProducts: [],
  productMeta: new Map(),    // product -> { realised, category }
  activityTypes: [],
  allProducts: []
};

const State = {
  mdoName: "",
  hq: "",
  region: "",
  territory: "",
  month: "",
  week: "",

  // 2) NPI
  npiRows: [],      // {product, plan, actual, opportunity, earned}

  // 3) Other products
  otherRows: [],    // {product, plan, actual, revenue}

  // 4) Activities update
  actRows: [],      // {typeObj, planNo, actualNo, npiNo}

  // 5) Activities photos
  photoRows: [],    // {activity, dataUrl, fileName}

  // 6) Next week plan
  nwRows: [],       // {product, plan, actual, revenue, incentiveEarned}

  // 7) Activities plan
  apRows: [],       // {typeObj, planNo, villages, villageNo}

  // 8) Special
  spDesc: "",
  spPhotoRows: []   // {activity, dataUrl, fileName}
};

function toNum(x){
  if (x === null || x === undefined) return 0;
  const s = String(x).replace(/,/g,"").trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function rs(n){
  return Math.round(Number(n||0)).toLocaleString("en-IN");
}
function showFatal(msg){
  $("fatalError").classList.remove("hidden");
  $("fatalError").innerHTML = msg || "Error";
}

function safeSheet(wb, ...names){
  for(const n of names){
    if(wb.Sheets[n]) return wb.Sheets[n];
  }
  return null;
}

function setOptions(sel, arr, placeholder){
  sel.innerHTML = "";
  const ph = document.createElement("option");
  ph.value = "";
  ph.textContent = placeholder || "Select";
  sel.appendChild(ph);
  for(const v of arr){
    const op = document.createElement("option");
    op.value = v;
    op.textContent = v;
    sel.appendChild(op);
  }
}

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

/* Activity selector: DD + Custom */
function typeSelector(typeObj, onChange){
  const wrap = document.createElement("div");
  wrap.style.display = "grid";
  wrap.style.gap = "6px";

  const sel = document.createElement("select");
  const custom = document.createElement("input");
  custom.type = "text";
  custom.placeholder = "Custom activity";
  custom.style.display = "none";

  setOptions(sel, Master.activityTypes, "Select activity");
  const optC = document.createElement("option");
  optC.value = "__CUSTOM__";
  optC.textContent = "Custom";
  sel.appendChild(optC);

  if(typeObj?.mode === "custom"){
    sel.value = "__CUSTOM__";
    custom.style.display = "block";
    custom.value = typeObj.value || "";
  } else {
    sel.value = typeObj?.value || "";
  }

  sel.addEventListener("change", ()=>{
    if(sel.value === "__CUSTOM__"){
      custom.style.display = "block";
      onChange({mode:"custom", value: custom.value || ""});
    }else{
      custom.style.display = "none";
      custom.value = "";
      onChange({mode:"preset", value: sel.value});
    }
  });

  custom.addEventListener("input", ()=>{
    onChange({mode:"custom", value: custom.value || ""});
  });

  wrap.appendChild(sel);
  wrap.appendChild(custom);
  return wrap;
}
function typeLabel(obj){
  if(!obj) return "";
  return obj.mode === "custom" ? (obj.value || "Custom") : (obj.value || "");
}

function delBtn(onClick){
  const b = document.createElement("button");
  b.className = "iconbtn";
  b.textContent = "✕";
  b.title = "Remove";
  b.type = "button";
  b.addEventListener("click", onClick);
  return b;
}

/* Cursor-stable numeric input: commit on BLUR */
function numInput(value, onCommit){
  const i = document.createElement("input");
  i.type = "number";
  i.min = "0";
  i.step = "any";
  i.value = value ?? "";
  i.addEventListener("blur", ()=> onCommit(i.value));
  return i;
}

/* ---------- Load admin excel ---------- */
async function loadMasterExcel(){
  if(typeof XLSX === "undefined"){
    throw new Error("XLSX library failed to load.");
  }
  const res = await fetch(CFG.excelPath, {cache:"no-store"});
  if(!res.ok){
    throw new Error(`Cannot load ${CFG.excelPath}.`);
  }
  const wb = XLSX.read(await res.arrayBuffer(), {type:"array"});

  const sRM = safeSheet(wb, "Region Mapping");
  const sNPI = safeSheet(wb, "NPI sheet", "NPI Sheet");
  const sPL  = safeSheet(wb, "Product List");
  const sAL  = safeSheet(wb, "Activity List");

  if(!sRM) throw new Error(`Missing sheet: Region Mapping`);
  if(!sNPI) throw new Error(`Missing sheet: NPI sheet / NPI Sheet`);
  if(!sPL) throw new Error(`Missing sheet: Product List`);
  if(!sAL) throw new Error(`Missing sheet: Activity List`);

  const rm = XLSX.utils.sheet_to_json(sRM, {defval:""});
  const npi = XLSX.utils.sheet_to_json(sNPI, {defval:""});
  const pl  = XLSX.utils.sheet_to_json(sPL, {defval:""});
  const al  = XLSX.utils.sheet_to_json(sAL, {defval:""});

  // Region -> Territory mapping (column "Territtory" spelling preserved)
  const rmap = new Map();
  for(const r of rm){
    const region = String(r["Region"]||"").trim();
    const terr = String(r["Territtory"]||"").trim();
    if(!region || !terr) continue;
    if(!rmap.has(region)) rmap.set(region, []);
    rmap.get(region).push(terr);
  }
  for(const [k,v] of rmap){
    rmap.set(k, [...new Set(v)].sort((a,b)=>a.localeCompare(b)));
  }
  Master.regionToTerritories = rmap;

  // NPI meta
  const npiMeta = new Map();
  for(const r of npi){
    const product = String(r["Product"]||"").trim();
    if(!product) continue;

    const realised = toNum(r["Realised Value in Rs"]);
    const incentiveRate = toNum(r["Incentive"]); // rate per L/Kg

    npiMeta.set(product, { realised, incentiveRate });
  }
  Master.npiMeta = npiMeta;
  Master.npiProducts = [...npiMeta.keys()].sort((a,b)=>a.localeCompare(b));

  // Product list meta + otherProducts
  const pMeta = new Map();
  const otherProducts = [];
  const allProductsSet = new Set(Master.npiProducts);

  for(const r of pl){
    const product = String(r["Product"]||"").trim();
    if(!product) continue;
    const realised = toNum(r["Realised Value"]);
    const category = String(r["Category"]||"").trim();
    pMeta.set(product, { realised, category });
    allProductsSet.add(product);
  }
  Master.productMeta = pMeta;

  for(const [p,meta] of pMeta.entries()){
    const cat = (meta.category||"").toLowerCase();
    if(cat !== "npi" && !Master.npiMeta.has(p)){
      otherProducts.push(p);
    }
  }
  Master.otherProducts = [...new Set(otherProducts)].sort((a,b)=>a.localeCompare(b));
  Master.allProducts = [...allProductsSet].sort((a,b)=>a.localeCompare(b));

  // Activity list
  const acts = [];
  for(const r of al){
    const a = String(r["Activity Type"]||"").trim();
    if(a) acts.push(a);
  }
  Master.activityTypes = [...new Set(acts)];
}

/* ---------- MDO dropdowns ---------- */
function initMdoDropdowns(){
  const regions = [...Master.regionToTerritories.keys()].sort((a,b)=>a.localeCompare(b));
  setOptions($("region"), regions, "Select Region");
  setOptions($("territory"), [], "Select Territory");
  setOptions($("month"), CFG.months, "Select Month");
  setOptions($("week"), CFG.weeks, "Select Week");

  $("region").addEventListener("change", ()=>{
    const terrs = Master.regionToTerritories.get($("region").value) || [];
    setOptions($("territory"), terrs, "Select Territory");
    $("territory").value = "";
  });
}

/* ---------- Calculations ---------- */
function recalcSummaries(){
  // NPI totals
  let oppTotal = 0;
  let earnedTotal = 0;

  for(const r of State.npiRows){
    const meta = Master.npiMeta.get(r.product) || { incentiveRate: 0 };
    const plan = toNum(r.plan);
    const actual = toNum(r.actual);

    // CONFIRMED BY YOU:
    // Opportunity = Plan × Incentive rate
    // Earned      = Actual × Incentive rate
    r.opportunity = plan * meta.incentiveRate;
    r.earned = actual * meta.incentiveRate;

    oppTotal += r.opportunity;
    earnedTotal += r.earned;
  }

  const lose = Math.max(0, oppTotal - earnedTotal);
  $("npiTotalEarned").textContent = rs(earnedTotal);
  $("npiTotalLose").textContent = rs(lose);

  // Other revenue total (Actual × realised value)
  let otherTotal = 0;
  for(const r of State.otherRows){
    const meta = Master.productMeta.get(r.product) || { realised: 0 };
    const actual = toNum(r.actual);
    r.revenue = actual * meta.realised;
    otherTotal += r.revenue;
  }
  $("otherTotalRevenue").textContent = rs(otherTotal);

  // Activities totals
  let p=0,a=0,n=0;
  for(const r of State.actRows){
    p += toNum(r.planNo);
    a += toNum(r.actualNo);
    n += toNum(r.npiNo);
  }
  $("actPlanTotal").textContent = String(p);
  $("actActualTotal").textContent = String(a);
  $("actNpiTotal").textContent = String(n);

  // Next week totals:
  // Revenue = Plan × realised (from Product List OR NPI sheet)
  // Incentive opportunity = Plan × incentiveRate (if available in NPI meta)
  let nwRevenue = 0;
  let nwOpp = 0;

  for(const r of State.nwRows){
    const plan = toNum(r.plan);

    const realised =
      (Master.productMeta.get(r.product)?.realised) ??
      (Master.npiMeta.get(r.product)?.realised) ??
      0;

    const rate = Master.npiMeta.get(r.product)?.incentiveRate ?? 0;

    r.revenue = plan * realised;
    r.incentiveEarned = plan * rate;

    nwRevenue += r.revenue;
    nwOpp += r.incentiveEarned;
  }

  $("nwRevenue").textContent = rs(nwRevenue);
  $("nwOpp").textContent = rs(nwOpp);
}

/* ---------- Renders ---------- */
function renderNpi(){
  const tbody = $("npiTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.npiRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    // Product dropdown from NPI sheet
    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.npiProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcSummaries();
      renderNpi();
    });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    // Plan
    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(numInput(r.plan, v=>{
      r.plan = v;
      recalcSummaries();
      renderNpi(); // ok on blur
    }));
    tr.appendChild(tdPlan);

    // Actual
    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actual, v=>{
      r.actual = v;
      recalcSummaries();
      renderNpi(); // ok on blur
    }));
    tr.appendChild(tdA);

    // Opportunity
    const tdOpp = document.createElement("td"); tdOpp.className="num";
    tdOpp.textContent = rs(r.opportunity || 0);
    tr.appendChild(tdOpp);

    // Earned
    const tdEarn = document.createElement("td"); tdEarn.className="num";
    tdEarn.textContent = rs(r.earned || 0);
    tr.appendChild(tdEarn);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{
      State.npiRows.splice(idx,1);
      recalcSummaries();
      renderNpi();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcSummaries();
}

function renderOther(){
  const tbody = $("otherTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.otherRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    // Product dropdown from Product List (non-NPI)
    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.otherProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcSummaries();
      renderOther();
    });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(numInput(r.plan, v=>{ r.plan = v; }));
    tr.appendChild(tdPlan);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actual, v=>{
      r.actual = v;
      recalcSummaries();
      renderOther(); // ok on blur
    }));
    tr.appendChild(tdA);

    const tdR = document.createElement("td"); tdR.className="num";
    tdR.textContent = rs(r.revenue || 0);
    tr.appendChild(tdR);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{
      State.otherRows.splice(idx,1);
      recalcSummaries();
      renderOther();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcSummaries();
}

function renderActivities(){
  const tbody = $("activityTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.actRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdT = document.createElement("td");
    tdT.appendChild(typeSelector(r.typeObj, (v)=>{ r.typeObj=v; }));
    tr.appendChild(tdT);

    const tdP = document.createElement("td"); tdP.className="num";
    tdP.appendChild(numInput(r.planNo, v=>{ r.planNo=v; recalcSummaries(); }));
    tr.appendChild(tdP);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actualNo, v=>{ r.actualNo=v; recalcSummaries(); }));
    tr.appendChild(tdA);

    const tdN = document.createElement("td"); tdN.className="num";
    tdN.appendChild(numInput(r.npiNo, v=>{ r.npiNo=v; recalcSummaries(); }));
    tr.appendChild(tdN);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{
      State.actRows.splice(idx,1);
      recalcSummaries();
      renderActivities();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcSummaries();
}

function renderPhotoRows(){
  const tbody = $("photoTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.photoRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdA = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.activityTypes, "Select activity");
    sel.value = r.activity || "";
    sel.addEventListener("change", ()=>{ r.activity = sel.value; renderPhotoPreview(); });
    tdA.appendChild(sel);
    tr.appendChild(tdA);

    const tdU = document.createElement("td");
    const up = document.createElement("input");
    up.type = "file";
    up.accept = "image/*";
    up.addEventListener("change", async (e)=>{
      const file = e.target.files?.[0];
      if(!file) return;
      r.fileName = file.name;
      r.dataUrl = await compressImage(file);
      renderPhotoPreview();
    });
    tdU.appendChild(up);
    tr.appendChild(tdU);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{
      State.photoRows.splice(idx,1);
      renderPhotoRows();
      renderPhotoPreview();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  renderPhotoPreview();
}

function renderPhotoPreview(){
  const grid = $("photoPreviewGrid");
  grid.innerHTML = "";

  const slice = State.photoRows.filter(p=>p.dataUrl).slice(0, CFG.maxActivityPhotos);
  slice.forEach((p)=>{
    const card = document.createElement("div");
    card.className = "photo-card";
    card.innerHTML = `
      <img alt="photo"/>
      <div class="photo-meta">
        <div class="photo-caption"></div>
      </div>
    `;
    card.querySelector("img").src = p.dataUrl;
    card.querySelector(".photo-caption").textContent = p.activity || "Activity";
    grid.appendChild(card);
  });
}

function renderNextWeek(){
  const tbody = $("nextWeekTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.nwRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.allProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcSummaries();
      renderNextWeek();
    });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(numInput(r.plan, v=>{
      r.plan = v;
      recalcSummaries();
      renderNextWeek();
    }));
    tr.appendChild(tdPlan);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actual, v=>{ r.actual = v; }));
    tr.appendChild(tdA);

    const tdR = document.createElement("td"); tdR.className="num";
    tdR.textContent = rs(r.revenue || 0);
    tr.appendChild(tdR);

    const tdI = document.createElement("td"); tdI.className="num";
    tdI.textContent = rs(r.incentiveEarned || 0);
    tr.appendChild(tdI);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{
      State.nwRows.splice(idx,1);
      recalcSummaries();
      renderNextWeek();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcSummaries();
}

function renderActivityPlan(){
  const tbody = $("activityPlanTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.apRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdT = document.createElement("td");
    tdT.appendChild(typeSelector(r.typeObj, (v)=>{ r.typeObj=v; }));
    tr.appendChild(tdT);

    const tdP = document.createElement("td"); tdP.className="num";
    tdP.appendChild(numInput(r.planNo, v=>{ r.planNo=v; }));
    tr.appendChild(tdP);

    const tdV = document.createElement("td");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.placeholder = "Village1, Village2, Village3";
    inp.value = r.villages || "";
    inp.addEventListener("blur", ()=>{
      r.villages = inp.value;
      renderActivityPlan();
    });
    tdV.appendChild(inp);
    tr.appendChild(tdV);

    const tdC = document.createElement("td"); tdC.className="num";
    const count = String(r.villages||"")
      .split(",")
      .map(s=>s.trim())
      .filter(Boolean).length;
    r.villageNo = count;
    tdC.textContent = String(count);
    tr.appendChild(tdC);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{ State.apRows.splice(idx,1); renderActivityPlan(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });
}

function renderSpecialPhotos(){
  const tbody = $("spPhotoTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.spPhotoRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdA = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.activityTypes, "Select activity");
    sel.value = r.activity || "";
    sel.addEventListener("change", ()=>{ r.activity = sel.value; renderSpecialPreview(); });
    tdA.appendChild(sel);
    tr.appendChild(tdA);

    const tdU = document.createElement("td");
    const up = document.createElement("input");
    up.type = "file";
    up.accept = "image/*";
    up.addEventListener("change", async (e)=>{
      const file = e.target.files?.[0];
      if(!file) return;
      r.fileName = file.name;
      r.dataUrl = await compressImage(file);
      renderSpecialPreview();
    });
    tdU.appendChild(up);
    tr.appendChild(tdU);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{
      State.spPhotoRows.splice(idx,1);
      renderSpecialPhotos();
      renderSpecialPreview();
    }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  renderSpecialPreview();
}

function renderSpecialPreview(){
  const grid = $("spPhotoPreviewGrid");
  grid.innerHTML = "";

  const slice = State.spPhotoRows.filter(p=>p.dataUrl).slice(0, CFG.maxSpecialPhotos);
  slice.forEach((p)=>{
    const card = document.createElement("div");
    card.className = "photo-card";
    card.innerHTML = `
      <img alt="photo"/>
      <div class="photo-meta">
        <div class="photo-caption"></div>
      </div>
    `;
    card.querySelector("img").src = p.dataUrl;
    card.querySelector(".photo-caption").textContent = p.activity || "Special";
    grid.appendChild(card);
  });
}

/* ---------- Buttons + wiring (THIS IS WHAT FIXES YOUR “NOT WORKING”) ---------- */
function wireButtons(){
  $("btnClearAll").addEventListener("click", ()=>{
    State.mdoName=""; State.hq=""; State.region=""; State.territory=""; State.month=""; State.week="";
    State.npiRows=[]; State.otherRows=[]; State.actRows=[];
    State.photoRows=[]; State.nwRows=[]; State.apRows=[];
    State.spDesc=""; State.spPhotoRows=[];

    $("mdoName").value="";
    $("hq").value="";
    $("region").value="";
    setOptions($("territory"), [], "Select Territory");
    $("month").value="";
    $("week").value="";
    $("spDesc").value="";

    renderNpi();
    renderOther();
    renderActivities();
    renderPhotoRows();
    renderNextWeek();
    renderActivityPlan();
    renderSpecialPhotos();
    recalcSummaries();
    window.scrollTo({top:0, behavior:"smooth"});
  });

  $("npiAdd").addEventListener("click", ()=>{
    if(State.npiRows.length >= CFG.maxNpiRows){ alert(`Max ${CFG.maxNpiRows} rows allowed.`); return; }
    State.npiRows.push({product:"", plan:"", actual:"", opportunity:0, earned:0});
    renderNpi();
  });
  $("npiClear").addEventListener("click", ()=>{ State.npiRows = []; renderNpi(); });

  $("otherAdd").addEventListener("click", ()=>{
    if(State.otherRows.length >= CFG.maxOtherRows){ alert(`Max ${CFG.maxOtherRows} rows allowed.`); return; }
    State.otherRows.push({product:"", plan:"", actual:"", revenue:0});
    renderOther();
  });
  $("otherClear").addEventListener("click", ()=>{ State.otherRows = []; renderOther(); });

  $("actAdd").addEventListener("click", ()=>{
    State.actRows.push({typeObj:{mode:"preset",value:""}, planNo:"", actualNo:"", npiNo:""});
    renderActivities();
  });
  $("actClear").addEventListener("click", ()=>{ State.actRows = []; renderActivities(); });

  $("photoAdd").addEventListener("click", ()=>{
    if(State.photoRows.length >= CFG.maxActivityPhotos){ alert(`Max ${CFG.maxActivityPhotos} photos allowed.`); return; }
    State.photoRows.push({activity:"", dataUrl:"", fileName:""});
    renderPhotoRows();
  });
  $("photoClear").addEventListener("click", ()=>{ State.photoRows = []; renderPhotoRows(); });

  $("nwAdd").addEventListener("click", ()=>{
    if(State.nwRows.length >= CFG.maxNextWeekRows){ alert(`Max ${CFG.maxNextWeekRows} rows allowed.`); return; }
    State.nwRows.push({product:"", plan:"", actual:"", revenue:0, incentiveEarned:0});
    renderNextWeek();
  });
  $("nwClear").addEventListener("click", ()=>{ State.nwRows = []; renderNextWeek(); recalcSummaries(); });

  $("apAdd").addEventListener("click", ()=>{
    State.apRows.push({typeObj:{mode:"preset",value:""}, planNo:"", villages:"", villageNo:0});
    renderActivityPlan();
  });
  $("apClear").addEventListener("click", ()=>{ State.apRows = []; renderActivityPlan(); });

  $("spDesc").addEventListener("input", ()=>{ State.spDesc = $("spDesc").value || ""; });

  $("spPhotoAdd").addEventListener("click", ()=>{
    if(State.spPhotoRows.length >= CFG.maxSpecialPhotos){ alert(`Max ${CFG.maxSpecialPhotos} photos allowed.`); return; }
    State.spPhotoRows.push({activity:"", dataUrl:"", fileName:""});
    renderSpecialPhotos();
  });
  $("spPhotoClear").addEventListener("click", ()=>{ State.spPhotoRows = []; renderSpecialPhotos(); });

  $("spClearAll").addEventListener("click", ()=>{
    State.spDesc = "";
    $("spDesc").value = "";
    State.spPhotoRows = [];
    renderSpecialPhotos();
  });

  $("btnPdf").addEventListener("click", ()=>{
    // capture latest MDO fields
    State.mdoName = $("mdoName").value || "";
    State.hq = $("hq").value || "";
    State.region = $("region").value || "";
    State.territory = $("territory").value || "";
    State.month = $("month").value || "";
    State.week = $("week").value || "";
    State.spDesc = $("spDesc").value || "";

    recalcSummaries();
    window.generateA4Pdf({ Master, State, typeLabel, rs });
  });
}

/* ---------- Boot (ensures DOM exists -> buttons always bind) ---------- */
async function boot(){
  try{
    await loadMasterExcel();
  }catch(err){
    console.error(err);
    showFatal(`Unable to load master data. Please check <b>${CFG.excelPath}</b> and refresh.`);
    return;
  }

  initMdoDropdowns();

  renderNpi();
  renderOther();
  renderActivities();
  renderPhotoRows();
  renderNextWeek();
  renderActivityPlan();
  renderSpecialPhotos();
  recalcSummaries();

  wireButtons();
}

document.addEventListener("DOMContentLoaded", boot);
