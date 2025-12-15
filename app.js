/* =========================================================
  Backend Excel (admin maintained):
  data/master_data.xlsx

  Sheets used:
  - Region Mapping: Region, Territtory
  - NPI sheet / NPI Sheet: Product, Realised Value in Rs, Incentive, Category
  - Product List: Product, Realised Value, Category
  - Activity List: Activity Type
========================================================= */

const $ = (id) => document.getElementById(id);

const CFG = {
  excelPath: "data/master_data.xlsx",
  months: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  weeks: ["1","2","3","4","5"],
  maxActivityPhotos: 16,
  maxSpecialPhotos: 4,
  storageKey: "performance_report_draft_v1"
};

const Master = {
  regionToTerritories: new Map(),
  npiProducts: [],
  npiMeta: new Map(),        // product -> { realised, incentive }
  otherProducts: [],
  productMeta: new Map(),    // product -> { realised, category }
  activityTypes: [],
  allProducts: []
};

const State = {
  // MDO
  mdoName: "",
  hq: "",
  region: "",
  territory: "",
  month: "",
  week: "",

  // tables
  npiRows: [],      // {product, plan, actual, incentiveEarned}
  otherRows: [],    // {product, plan, actual, revenue}
  actRows: [],      // {typeObj, planNo, actualNo, npiNo}

  // photos (rows)
  photoRows: [],    // {activity, dataUrl, fileName}
  nwRows: [],       // {product, plan, actual, revenue, incentive}
  apRows: [],       // {typeObj, planNo, villages, villageNo}

  // special
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
  if(msg) $("fatalError").innerHTML = msg;
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

function saveDraft(){
  const payload = {
    ...State,
    mdoName: $("mdoName").value || "",
    hq: $("hq").value || "",
    region: $("region").value || "",
    territory: $("territory").value || "",
    month: $("month").value || "",
    week: $("week").value || "",
    spDesc: $("spDesc").value || ""
  };
  localStorage.setItem(CFG.storageKey, JSON.stringify(payload));
}

function loadDraft(){
  const raw = localStorage.getItem(CFG.storageKey);
  if(!raw) return;
  try{
    const d = JSON.parse(raw);
    Object.assign(State, d);
  }catch{}
}

function clearAll(){
  localStorage.removeItem(CFG.storageKey);
  location.reload();
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

function typeSelector(typeObj, onChange){
  // DD from Master.activityTypes + "Custom"
  const wrap = document.createElement("div");
  wrap.style.display = "grid";
  wrap.style.gap = "6px";

  const sel = document.createElement("select");
  const custom = document.createElement("input");
  custom.type = "text";
  custom.placeholder = "Custom activity";
  custom.style.display = "none";

  const opts = [...Master.activityTypes];
  setOptions(sel, opts, "Select activity");
  // add Custom option
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

/* ---------------- Excel Load ---------------- */

async function loadMasterExcel(){
  if(typeof XLSX === "undefined"){
    throw new Error("XLSX library failed to load.");
  }
  const res = await fetch(CFG.excelPath, {cache:"no-store"});
  if(!res.ok){
    throw new Error(`Cannot load ${CFG.excelPath}. Make sure it exists in your repo.`);
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

  // Region mapping
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
    const incentive = toNum(r["Incentive"]);
    npiMeta.set(product, { realised, incentive });
  }
  Master.npiMeta = npiMeta;
  Master.npiProducts = [...npiMeta.keys()].sort((a,b)=>a.localeCompare(b));

  // Product list meta
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

  // other products = product list where category != NPI and not present in NPI sheet
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

/* ---------------- UI init ---------------- */

function initMdoDropdowns(){
  const regions = [...Master.regionToTerritories.keys()].sort((a,b)=>a.localeCompare(b));
  setOptions($("region"), regions, "Select Region");

  $("region").addEventListener("change", ()=>{
    const terrs = Master.regionToTerritories.get($("region").value) || [];
    setOptions($("territory"), terrs, "Select Territory");
    $("territory").value = "";
    saveDraft();
  });
  $("territory").addEventListener("change", saveDraft);

  setOptions($("month"), CFG.months, "Select Month");
  setOptions($("week"), CFG.weeks, "Select Week");

  ["mdoName","hq","month","week"].forEach(id=>{
    $(id).addEventListener("input", saveDraft);
    $(id).addEventListener("change", saveDraft);
  });
}

function applyDraft(){
  $("mdoName").value = State.mdoName || "";
  $("hq").value = State.hq || "";
  $("month").value = State.month || "";
  $("week").value = State.week || "";
  $("spDesc").value = State.spDesc || "";

  $("region").value = State.region || "";
  const terrs = Master.regionToTerritories.get($("region").value) || [];
  setOptions($("territory"), terrs, "Select Territory");
  $("territory").value = State.territory || "";
}

function delBtn(onClick){
  const b = document.createElement("button");
  b.className = "iconbtn";
  b.textContent = "✕";
  b.title = "Remove";
  b.addEventListener("click", onClick);
  return b;
}

function numInput(value, onInput){
  const i = document.createElement("input");
  i.type = "number";
  i.min = "0";
  i.step = "any";
  i.value = value ?? "";
  i.addEventListener("input", ()=>onInput(i.value));
  return i;
}

/* ---------------- Renders ---------------- */

function recalcSummaries(){
  // NPI total incentive
  let npiTotal = 0;
  for(const r of State.npiRows){
    const meta = Master.npiMeta.get(r.product) || { incentive: 0 };
    const actual = toNum(r.actual);
    r.incentiveEarned = actual * meta.incentive;
    npiTotal += r.incentiveEarned;
  }
  $("npiTotalIncentive").textContent = rs(npiTotal);

  // Other total revenue
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

  // Next week totals
  let nwRev=0, nwInc=0;
  for(const r of State.nwRows){
    const actual = toNum(r.actual);
    const realised =
      (Master.productMeta.get(r.product)?.realised) ??
      (Master.npiMeta.get(r.product)?.realised) ??
      0;

    const incRate = Master.npiMeta.get(r.product)?.incentive ?? 0;

    r.revenue = actual * realised;
    r.incentive = actual * incRate;

    nwRev += r.revenue;
    nwInc += r.incentive;
  }
  $("nwTotalRevenue").textContent = rs(nwRev);
  $("nwTotalIncentive").textContent = rs(nwInc);

  // Activities plan village count computed per row in render

  saveDraft();
}

function renderNpi(){
  const tbody = $("npiTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.npiRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    // product dd
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

    // plan
    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(numInput(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(tdPlan);

    // actual
    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actual, v=>{ r.actual=v; recalcSummaries(); renderNpi(); }));
    tr.appendChild(tdA);

    // incentive earned
    const tdInc = document.createElement("td"); tdInc.className="num";
    tdInc.textContent = rs(r.incentiveEarned || 0);
    tr.appendChild(tdInc);

    // delete
    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{ State.npiRows.splice(idx,1); renderNpi(); recalcSummaries(); }));
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
    tdPlan.appendChild(numInput(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(tdPlan);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actual, v=>{ r.actual=v; recalcSummaries(); renderOther(); }));
    tr.appendChild(tdA);

    const tdR = document.createElement("td"); tdR.className="num";
    tdR.textContent = rs(r.revenue || 0);
    tr.appendChild(tdR);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{ State.otherRows.splice(idx,1); renderOther(); recalcSummaries(); }));
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
    tdT.appendChild(typeSelector(r.typeObj, (v)=>{ r.typeObj=v; saveDraft(); }));
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
    tdDel.appendChild(delBtn(()=>{ State.actRows.splice(idx,1); renderActivities(); recalcSummaries(); }));
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
    sel.addEventListener("change", ()=>{ r.activity = sel.value; saveDraft(); renderPhotoPreview(); });
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
      saveDraft();
      renderPhotoPreview();
    });
    tdU.appendChild(up);
    tr.appendChild(tdU);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{ State.photoRows.splice(idx,1); renderPhotoRows(); renderPhotoPreview(); saveDraft(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  renderPhotoPreview();
}

function renderPhotoPreview(){
  const grid = $("photoPreviewGrid");
  grid.innerHTML = "";

  const slice = State.photoRows.filter(p=>p.dataUrl).slice(0, CFG.maxActivityPhotos);
  slice.forEach((p, idx)=>{
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
    const del = delBtn(()=>{
      // remove the preview row by finding its row index in State.photoRows
      const realIdx = State.photoRows.findIndex(x=>x.dataUrl===p.dataUrl && x.fileName===p.fileName);
      if(realIdx >= 0){
        State.photoRows.splice(realIdx,1);
        renderPhotoRows();
        saveDraft();
      }
    });
    card.querySelector(".photo-meta").appendChild(del);
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
    tdPlan.appendChild(numInput(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(tdPlan);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numInput(r.actual, v=>{ r.actual=v; recalcSummaries(); renderNextWeek(); }));
    tr.appendChild(tdA);

    const tdR = document.createElement("td"); tdR.className="num";
    tdR.textContent = rs(r.revenue || 0);
    tr.appendChild(tdR);

    const tdI = document.createElement("td"); tdI.className="num";
    tdI.textContent = rs(r.incentive || 0);
    tr.appendChild(tdI);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{ State.nwRows.splice(idx,1); renderNextWeek(); recalcSummaries(); }));
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
    tdT.appendChild(typeSelector(r.typeObj, (v)=>{ r.typeObj=v; saveDraft(); }));
    tr.appendChild(tdT);

    const tdP = document.createElement("td"); tdP.className="num";
    tdP.appendChild(numInput(r.planNo, v=>{ r.planNo=v; saveDraft(); }));
    tr.appendChild(tdP);

    const tdV = document.createElement("td");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.placeholder = "Village1, Village2, Village3";
    inp.value = r.villages || "";
    inp.addEventListener("input", ()=>{
      r.villages = inp.value;
      renderActivityPlan();
      saveDraft();
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
    tdDel.appendChild(delBtn(()=>{ State.apRows.splice(idx,1); renderActivityPlan(); saveDraft(); }));
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
    sel.addEventListener("change", ()=>{ r.activity = sel.value; saveDraft(); renderSpecialPreview(); });
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
      saveDraft();
      renderSpecialPreview();
    });
    tdU.appendChild(up);
    tr.appendChild(tdU);

    const tdDel = document.createElement("td");
    tdDel.appendChild(delBtn(()=>{ State.spPhotoRows.splice(idx,1); renderSpecialPhotos(); renderSpecialPreview(); saveDraft(); }));
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

/* ---------------- Buttons ---------------- */

function wireButtons(){
  $("btnClearAll").addEventListener("click", clearAll);

  // NPI
  $("npiAdd").addEventListener("click", ()=>{ State.npiRows.push({product:"", plan:"", actual:"", incentiveEarned:0}); renderNpi(); });
  $("npiClear").addEventListener("click", ()=>{ State.npiRows = []; renderNpi(); });

  // Other products
  $("otherAdd").addEventListener("click", ()=>{ State.otherRows.push({product:"", plan:"", actual:"", revenue:0}); renderOther(); });
  $("otherClear").addEventListener("click", ()=>{ State.otherRows = []; renderOther(); });

  // Activities update
  $("actAdd").addEventListener("click", ()=>{ State.actRows.push({typeObj:{mode:"preset",value:""}, planNo:"", actualNo:"", npiNo:""}); renderActivities(); });
  $("actClear").addEventListener("click", ()=>{ State.actRows = []; renderActivities(); });

  // Activity photos
  $("photoAdd").addEventListener("click", ()=>{
    if(State.photoRows.length >= CFG.maxActivityPhotos) return;
    State.photoRows.push({activity:"", dataUrl:"", fileName:""});
    renderPhotoRows();
    saveDraft();
  });
  $("photoClear").addEventListener("click", ()=>{ State.photoRows = []; renderPhotoRows(); saveDraft(); });

  // Next week
  $("nwAdd").addEventListener("click", ()=>{ State.nwRows.push({product:"", plan:"", actual:"", revenue:0, incentive:0}); renderNextWeek(); });
  $("nwClear").addEventListener("click", ()=>{ State.nwRows = []; renderNextWeek(); });

  // Activity plan
  $("apAdd").addEventListener("click", ()=>{ State.apRows.push({typeObj:{mode:"preset",value:""}, planNo:"", villages:"", villageNo:0}); renderActivityPlan(); });
  $("apClear").addEventListener("click", ()=>{ State.apRows = []; renderActivityPlan(); });

  // Special
  $("spDesc").addEventListener("input", saveDraft);

  $("spPhotoAdd").addEventListener("click", ()=>{
    if(State.spPhotoRows.length >= CFG.maxSpecialPhotos) return;
    State.spPhotoRows.push({activity:"", dataUrl:"", fileName:""});
    renderSpecialPhotos();
    saveDraft();
  });
  $("spPhotoClear").addEventListener("click", ()=>{ State.spPhotoRows = []; renderSpecialPhotos(); saveDraft(); });

  $("spClearAll").addEventListener("click", ()=>{
    State.spDesc = "";
    $("spDesc").value = "";
    State.spPhotoRows = [];
    renderSpecialPhotos();
    saveDraft();
  });

  // PDF
  $("btnPdf").addEventListener("click", ()=>{
    saveDraft();
    window.generateA4Pdf({
      Master,
      State: {
        ...State,
        mdoName: $("mdoName").value || "",
        hq: $("hq").value || "",
        region: $("region").value || "",
        territory: $("territory").value || "",
        month: $("month").value || "",
        week: $("week").value || "",
        spDesc: $("spDesc").value || ""
      },
      typeLabel,
      rs
    });
  });
}

/* ---------------- Boot ---------------- */

async function boot(){
  loadDraft();

  try{
    await loadMasterExcel();
  }catch(err){
    console.error(err);
    showFatal(`Unable to load master data. Please check <b>${CFG.excelPath}</b> and refresh.`);
    return;
  }

  initMdoDropdowns();
  applyDraft();

  // initial renders
  renderNpi();
  renderOther();
  renderActivities();
  renderPhotoRows();
  renderNextWeek();
  renderActivityPlan();
  renderSpecialPhotos();

  wireButtons();
  recalcSummaries();
}

document.addEventListener("DOMContentLoaded", boot);
