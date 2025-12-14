/* =========================================================
  BACKEND MASTER DATA (ADMIN):
  data/master_data.xlsx

  Sheet 1: "Region Mapping"
    Columns: Region, Territtory   (note spelling)

  Sheet 2: "NPI Sheet"
    Columns: Product, Realised Value in Rs, Portfolio, Incentive, Category

  Sheet 3: "Product List"
    Columns: Product, Realised Value, Portfolio, Incentive, Category
========================================================= */

const $ = (id) => document.getElementById(id);

const CFG = {
  excelPath: "data/master_data.xlsx",
  months: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  weeks: ["1","2","3","4","5"],
  activityTypes: ["One to One Farmer","Village Farmer Meeting","Field Day","OFM","LFM","Demo","Special Campaign","Custom"],
  limits: { activityPhotos: 16, specialPhotos: 4 },
  storageKey: "mdo_weekly_report_draft_v3"
};

const Master = {
  loaded: false,
  regionToTerritories: new Map(),  // region -> [territory...]
  npiProducts: [],                 // [product...]
  npiMeta: new Map(),              // product -> { realised, incentive }
  otherProducts: [],               // [product...] from Product List where Category != NPI
  productMeta: new Map(),          // product -> { realised, category }
  allProducts: []                  // union of product list + NPI
};

const State = {
  mdo: { name:"", headquarter:"", region:"", territory:"", month:"", week:"" },
  npiRows: [],           // {product, plan, ach, revenue, incentiveEarned}
  otherRows: [],         // {product, plan, ach, revenue}
  activityRows: [],      // {typeObj, planned, achieved, npiFocused}
  activityPhotos: [],    // {type, dataUrl}
  nextWeekRows: [],      // {product, placement, liquidation, revenue}
  activityPlanRows: [],  // {typeObj, planned, villages}
  special: {
    desc: "",
    photos: [],          // {type, dataUrl}
    rows: []             // {product, placement, liquidation, revenue}
  }
};

function showError(msg){
  const box = $("errorBox");
  box.textContent = msg;
  box.classList.remove("hidden");
}

function hideLoading(){
  $("loading").classList.add("hidden");
}

function moneyINR(n){
  const v = Math.round(Number(n || 0));
  return "₹ " + v.toLocaleString("en-IN");
}

function toNum(x){
  if (x === null || x === undefined) return 0;
  const s = String(x).replace(/,/g,"").trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function setOptions(sel, arr, placeholder="Select"){
  sel.innerHTML = "";
  const ph = document.createElement("option");
  ph.value = "";
  ph.textContent = placeholder;
  sel.appendChild(ph);
  for(const it of arr){
    const op = document.createElement("option");
    op.value = it;
    op.textContent = it;
    sel.appendChild(op);
  }
}

function saveDraft(){
  State.mdo.name = $("mdoName").value || "";
  State.mdo.headquarter = $("headquarter").value || "";
  State.mdo.region = $("regionSelect").value || "";
  State.mdo.territory = $("territorySelect").value || "";
  State.mdo.month = $("monthSelect").value || "";
  State.mdo.week = $("weekSelect").value || "";
  State.special.desc = $("specialDesc").value || "";
  localStorage.setItem(CFG.storageKey, JSON.stringify(State));
}

function loadDraft(){
  const raw = localStorage.getItem(CFG.storageKey);
  if(!raw) return;
  try{
    const parsed = JSON.parse(raw);
    // soft merge
    Object.assign(State, parsed);
    State.mdo = Object.assign({name:"",headquarter:"",region:"",territory:"",month:"",week:""}, State.mdo||{});
    State.npiRows = Array.isArray(State.npiRows) ? State.npiRows : [];
    State.otherRows = Array.isArray(State.otherRows) ? State.otherRows : [];
    State.activityRows = Array.isArray(State.activityRows) ? State.activityRows : [];
    State.activityPhotos = Array.isArray(State.activityPhotos) ? State.activityPhotos : [];
    State.nextWeekRows = Array.isArray(State.nextWeekRows) ? State.nextWeekRows : [];
    State.activityPlanRows = Array.isArray(State.activityPlanRows) ? State.activityPlanRows : [];
    State.special = State.special || {desc:"",photos:[],rows:[]};
    State.special.photos = Array.isArray(State.special.photos) ? State.special.photos : [];
    State.special.rows = Array.isArray(State.special.rows) ? State.special.rows : [];
  }catch(e){}
}

function clearAll(){
  localStorage.removeItem(CFG.storageKey);
  location.reload();
}

/* ---------------- Excel loading (backend) ---------------- */

async function loadMasterExcel(){
  const res = await fetch(CFG.excelPath, { cache: "no-store" });
  if(!res.ok){
    throw new Error(`Cannot load ${CFG.excelPath}. Ensure the file exists in /data/ and GitHub Pages is serving it.`);
  }
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, {type:"array"});

  const sheetRM = wb.Sheets["Region Mapping"];
  const sheetNPI = wb.Sheets["NPI Sheet"];
  const sheetPL = wb.Sheets["Product List"];

  if(!sheetRM) throw new Error(`Sheet missing: "Region Mapping"`);
  if(!sheetNPI) throw new Error(`Sheet missing: "NPI Sheet"`);
  if(!sheetPL) throw new Error(`Sheet missing: "Product List"`);

  const rm = XLSX.utils.sheet_to_json(sheetRM, { defval:"" });
  const npi = XLSX.utils.sheet_to_json(sheetNPI, { defval:"" });
  const pl  = XLSX.utils.sheet_to_json(sheetPL, { defval:"" });

  // Region -> territories
  const map = new Map();
  for(const r of rm){
    const region = String(r["Region"]||"").trim();
    const terr = String(r["Territtory"]||"").trim();
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
    const product = String(r["Product"]||"").trim();
    if(!product) continue;
    npiMeta.set(product, {
      realised: toNum(r["Realised Value in Rs"]),
      incentive: toNum(r["Incentive"])
    });
  }
  Master.npiMeta = npiMeta;
  Master.npiProducts = [...npiMeta.keys()].sort((a,b)=>a.localeCompare(b));

  // Product list meta
  const productMeta = new Map();
  const other = [];
  const all = new Set();

  for(const r of pl){
    const product = String(r["Product"]||"").trim();
    if(!product) continue;
    const category = String(r["Category"]||"").trim();
    const realised = toNum(r["Realised Value"]);
    productMeta.set(product, { realised, category });
    all.add(product);
    if(category.toLowerCase() !== "npi") other.push(product);
  }
  // add NPI products into all
  for(const p of Master.npiProducts) all.add(p);

  Master.productMeta = productMeta;
  Master.otherProducts = [...new Set(other)].sort((a,b)=>a.localeCompare(b));
  Master.allProducts = [...all].sort((a,b)=>a.localeCompare(b));

  Master.loaded = true;
}

/* ---------------- UI helpers ---------------- */

function makeDelBtn(onClick){
  const b = document.createElement("button");
  b.className = "iconbtn";
  b.textContent = "✕";
  b.title = "Remove";
  b.addEventListener("click", onClick);
  return b;
}

function numberInput(value, onInput){
  const i = document.createElement("input");
  i.type = "number";
  i.min = "0";
  i.step = "any";
  i.value = value ?? "";
  i.addEventListener("input", ()=>onInput(i.value));
  return i;
}

function activitySelector(typeObj, onChange){
  // typeObj: {mode:"preset"|"custom", value:""}
  const wrap = document.createElement("div");
  wrap.style.display = "grid";
  wrap.style.gap = "6px";

  const sel = document.createElement("select");
  setOptions(sel, CFG.activityTypes, "Select activity");
  const custom = document.createElement("input");
  custom.type = "text";
  custom.placeholder = "Custom activity name";
  custom.classList.add("hidden");

  if(typeObj?.mode === "custom"){
    sel.value = "Custom";
    custom.classList.remove("hidden");
    custom.value = typeObj.value || "";
  }else{
    sel.value = typeObj?.value || "";
  }

  sel.addEventListener("change", ()=>{
    if(sel.value === "Custom"){
      custom.classList.remove("hidden");
      onChange({mode:"custom", value: custom.value || ""});
    }else{
      custom.classList.add("hidden");
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

function getActivityLabel(obj){
  if(!obj) return "";
  if(obj.mode === "custom") return obj.value || "Custom";
  return obj.value || "";
}

/* ---------------- Render functions ---------------- */

function recalcAll(){
  // NPI totals
  let npiRev = 0, npiInc = 0;
  for(const r of State.npiRows){
    const meta = Master.npiMeta.get(r.product) || {realised:0,incentive:0};
    const ach = toNum(r.ach);
    r.revenue = ach * meta.realised;
    r.incentiveEarned = ach * meta.incentive;
    npiRev += r.revenue;
    npiInc += r.incentiveEarned;
  }
  $("npiTotalRevenue").textContent = moneyINR(npiRev);
  $("npiTotalIncentive").textContent = moneyINR(npiInc);

  // Other totals
  let otherRev = 0;
  for(const r of State.otherRows){
    const meta = Master.productMeta.get(r.product) || {realised:0};
    const ach = toNum(r.ach);
    r.revenue = ach * (meta.realised || 0);
    otherRev += r.revenue;
  }
  $("otherTotalRevenue").textContent = moneyINR(otherRev);

  // Next week totals
  let nwRev = 0;
  for(const r of State.nextWeekRows){
    const liq = toNum(r.liquidation);
    const fromPL = Master.productMeta.get(r.product);
    const fromNPI = Master.npiMeta.get(r.product);
    const realised = (fromPL && fromPL.realised) ? fromPL.realised : (fromNPI ? fromNPI.realised : 0);
    r.revenue = liq * realised;
    nwRev += r.revenue;
  }
  $("nwTotalRevenue").textContent = moneyINR(nwRev);

  // Special totals
  let spRev = 0;
  for(const r of State.special.rows){
    const liq = toNum(r.liquidation);
    const fromPL = Master.productMeta.get(r.product);
    const fromNPI = Master.npiMeta.get(r.product);
    const realised = (fromPL && fromPL.realised) ? fromPL.realised : (fromNPI ? fromNPI.realised : 0);
    r.revenue = liq * realised;
    spRev += r.revenue;
  }
  $("specialTotalRevenue").textContent = moneyINR(spRev);

  // Activity totals
  let p=0,a=0,n=0;
  for(const r of State.activityRows){
    p += toNum(r.planned);
    a += toNum(r.achieved);
    n += toNum(r.npiFocused);
  }
  $("actPlanTotal").textContent = String(p);
  $("actAchTotal").textContent = String(a);
  $("actNpiTotal").textContent = String(n);

  saveDraft();
}

function renderNpiTable(){
  const tbody = $("npiTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.npiRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    // product select
    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.npiProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{ r.product = sel.value; recalcAll(); renderNpiTable(); });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    // plan
    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(numberInput(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(tdPlan);

    // achievement
    const tdAch = document.createElement("td"); tdAch.className="num";
    tdAch.appendChild(numberInput(r.ach, v=>{ r.ach=v; recalcAll(); renderNpiTable(); }));
    tr.appendChild(tdAch);

    // revenue / incentive earned
    const tdRev = document.createElement("td"); tdRev.className="num"; tdRev.textContent = moneyINR(r.revenue||0);
    const tdInc = document.createElement("td"); tdInc.className="num"; tdInc.textContent = moneyINR(r.incentiveEarned||0);
    tr.appendChild(tdRev);
    tr.appendChild(tdInc);

    // delete
    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelBtn(()=>{ State.npiRows.splice(idx,1); renderNpiTable(); recalcAll(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcAll();
}

function renderOtherTable(){
  const tbody = $("otherTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.otherRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.otherProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{ r.product = sel.value; recalcAll(); renderOtherTable(); });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    const tdPlan = document.createElement("td"); tdPlan.className="num";
    tdPlan.appendChild(numberInput(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(tdPlan);

    const tdAch = document.createElement("td"); tdAch.className="num";
    tdAch.appendChild(numberInput(r.ach, v=>{ r.ach=v; recalcAll(); renderOtherTable(); }));
    tr.appendChild(tdAch);

    const tdRev = document.createElement("td"); tdRev.className="num"; tdRev.textContent = moneyINR(r.revenue||0);
    tr.appendChild(tdRev);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelBtn(()=>{ State.otherRows.splice(idx,1); renderOtherTable(); recalcAll(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcAll();
}

function renderActivityTable(){
  const tbody = $("activityTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.activityRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdT = document.createElement("td");
    tdT.appendChild(activitySelector(r.typeObj, (v)=>{ r.typeObj=v; saveDraft(); }));
    tr.appendChild(tdT);

    const tdP = document.createElement("td"); tdP.className="num";
    tdP.appendChild(numberInput(r.planned, v=>{ r.planned=v; recalcAll(); }));
    tr.appendChild(tdP);

    const tdA = document.createElement("td"); tdA.className="num";
    tdA.appendChild(numberInput(r.achieved, v=>{ r.achieved=v; recalcAll(); }));
    tr.appendChild(tdA);

    const tdN = document.createElement("td"); tdN.className="num";
    tdN.appendChild(numberInput(r.npiFocused, v=>{ r.npiFocused=v; recalcAll(); }));
    tr.appendChild(tdN);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelBtn(()=>{ State.activityRows.splice(idx,1); renderActivityTable(); recalcAll(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcAll();
}

function renderNextWeekTable(){
  const tbody = $("nextWeekTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.nextWeekRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.allProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{ r.product = sel.value; recalcAll(); renderNextWeekTable(); });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    const tdPl = document.createElement("td"); tdPl.className="num";
    tdPl.appendChild(numberInput(r.placement, v=>{ r.placement=v; saveDraft(); }));
    tr.appendChild(tdPl);

    const tdLiq = document.createElement("td"); tdLiq.className="num";
    tdLiq.appendChild(numberInput(r.liquidation, v=>{ r.liquidation=v; recalcAll(); renderNextWeekTable(); }));
    tr.appendChild(tdLiq);

    const tdRev = document.createElement("td"); tdRev.className="num"; tdRev.textContent = moneyINR(r.revenue||0);
    tr.appendChild(tdRev);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelBtn(()=>{ State.nextWeekRows.splice(idx,1); renderNextWeekTable(); recalcAll(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcAll();
}

function renderActivityPlanTable(){
  const tbody = $("activityPlanTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.activityPlanRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdT = document.createElement("td");
    tdT.appendChild(activitySelector(r.typeObj, (v)=>{ r.typeObj=v; saveDraft(); }));
    tr.appendChild(tdT);

    const tdP = document.createElement("td"); tdP.className="num";
    tdP.appendChild(numberInput(r.planned, v=>{ r.planned=v; saveDraft(); }));
    tr.appendChild(tdP);

    const tdV = document.createElement("td");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.placeholder = "Village1, Village2, Village3";
    inp.value = r.villages || "";
    inp.addEventListener("input", ()=>{ r.villages = inp.value; renderActivityPlanTable(); saveDraft(); });
    tdV.appendChild(inp);
    tr.appendChild(tdV);

    const tdC = document.createElement("td"); tdC.className="num";
    const count = String(r.villages||"").split(",").map(s=>s.trim()).filter(Boolean).length;
    tdC.textContent = String(count);
    tr.appendChild(tdC);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelBtn(()=>{ State.activityPlanRows.splice(idx,1); renderActivityPlanTable(); saveDraft(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });
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

function renderActivityPhotos(){
  $("photoCount").textContent = String(State.activityPhotos.length);
  const grid = $("photoGrid");
  grid.innerHTML = "";
  State.activityPhotos.forEach((p, idx)=>{
    const card = document.createElement("div");
    card.className = "photo-card";
    card.innerHTML = `
      <img alt="photo"/>
      <div class="photo-meta">
        <div class="photo-caption"></div>
      </div>
    `;
    card.querySelector("img").src = p.dataUrl;
    card.querySelector(".photo-caption").textContent = p.type || "Activity";

    const del = makeDelBtn(()=>{ State.activityPhotos.splice(idx,1); renderActivityPhotos(); saveDraft(); });
    card.querySelector(".photo-meta").appendChild(del);

    grid.appendChild(card);
  });
}

function renderSpecial(){
  $("specialDesc").value = State.special.desc || "";

  $("specialPhotoCount").textContent = String(State.special.photos.length);
  const grid = $("specialPhotoGrid");
  grid.innerHTML = "";
  State.special.photos.forEach((p, idx)=>{
    const card = document.createElement("div");
    card.className = "photo-card";
    card.innerHTML = `
      <img alt="photo"/>
      <div class="photo-meta">
        <div class="photo-caption"></div>
      </div>
    `;
    card.querySelector("img").src = p.dataUrl;
    card.querySelector(".photo-caption").textContent = p.type || "Special";

    const del = makeDelBtn(()=>{ State.special.photos.splice(idx,1); renderSpecial(); saveDraft(); });
    card.querySelector(".photo-meta").appendChild(del);

    grid.appendChild(card);
  });

  // table
  const tbody = $("specialTable").querySelector("tbody");
  tbody.innerHTML = "";
  State.special.rows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const tdP = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, Master.allProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{ r.product = sel.value; recalcAll(); renderSpecial(); });
    tdP.appendChild(sel);
    tr.appendChild(tdP);

    const tdPl = document.createElement("td"); tdPl.className="num";
    tdPl.appendChild(numberInput(r.placement, v=>{ r.placement=v; saveDraft(); }));
    tr.appendChild(tdPl);

    const tdL = document.createElement("td"); tdL.className="num";
    tdL.appendChild(numberInput(r.liquidation, v=>{ r.liquidation=v; recalcAll(); renderSpecial(); }));
    tr.appendChild(tdL);

    const tdR = document.createElement("td"); tdR.className="num"; tdR.textContent = moneyINR(r.revenue||0);
    tr.appendChild(tdR);

    const tdDel = document.createElement("td");
    tdDel.appendChild(makeDelBtn(()=>{ State.special.rows.splice(idx,1); renderSpecial(); recalcAll(); }));
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  });

  recalcAll();
}

/* ---------------- Boot & events ---------------- */

function applyDraftToUI(){
  $("mdoName").value = State.mdo.name || "";
  $("headquarter").value = State.mdo.headquarter || "";
  $("monthSelect").value = State.mdo.month || "";
  $("weekSelect").value = State.mdo.week || "";
  $("specialDesc").value = State.special.desc || "";
}

function bindMdoDropdowns(){
  const regions = [...Master.regionToTerritories.keys()].sort((a,b)=>a.localeCompare(b));
  setOptions($("regionSelect"), regions, "Select region");

  $("regionSelect").value = State.mdo.region || "";
  const terrs = Master.regionToTerritories.get($("regionSelect").value) || [];
  setOptions($("territorySelect"), terrs, "Select territory");
  $("territorySelect").value = State.mdo.territory || "";

  $("regionSelect").addEventListener("change", ()=>{
    const t = Master.regionToTerritories.get($("regionSelect").value) || [];
    setOptions($("territorySelect"), t, "Select territory");
    $("territorySelect").value = "";
    saveDraft();
  });
  $("territorySelect").addEventListener("change", saveDraft);
}

function bindStaticDropdowns(){
  setOptions($("monthSelect"), CFG.months, "Select month");
  setOptions($("weekSelect"), CFG.weeks, "Select week");
  setOptions($("photoTypeSelect"), CFG.activityTypes, "Select activity type");
  setOptions($("specialPhotoTypeSelect"), CFG.activityTypes, "Select activity type");
}

function wireButtons(){
  $("btnSave").addEventListener("click", saveDraft);
  $("btnClearAll").addEventListener("click", clearAll);

  $("npiAdd").addEventListener("click", ()=>{ State.npiRows.push({product:"",plan:"",ach:"",revenue:0,incentiveEarned:0}); renderNpiTable(); });
  $("npiClear").addEventListener("click", ()=>{ State.npiRows=[]; renderNpiTable(); });

  $("otherAdd").addEventListener("click", ()=>{ State.otherRows.push({product:"",plan:"",ach:"",revenue:0}); renderOtherTable(); });
  $("otherClear").addEventListener("click", ()=>{ State.otherRows=[]; renderOtherTable(); });

  $("actAdd").addEventListener("click", ()=>{ State.activityRows.push({typeObj:{mode:"preset",value:""},planned:"",achieved:"",npiFocused:""}); renderActivityTable(); });
  $("actClear").addEventListener("click", ()=>{ State.activityRows=[]; renderActivityTable(); });

  $("photoClear").addEventListener("click", ()=>{ State.activityPhotos=[]; renderActivityPhotos(); saveDraft(); });

  $("nwAdd").addEventListener("click", ()=>{ State.nextWeekRows.push({product:"",placement:"",liquidation:"",revenue:0}); renderNextWeekTable(); });
  $("nwClear").addEventListener("click", ()=>{ State.nextWeekRows=[]; renderNextWeekTable(); });

  $("apAdd").addEventListener("click", ()=>{ State.activityPlanRows.push({typeObj:{mode:"preset",value:""},planned:"",villages:""}); renderActivityPlanTable(); });
  $("apClear").addEventListener("click", ()=>{ State.activityPlanRows=[]; renderActivityPlanTable(); });

  $("specialAdd").addEventListener("click", ()=>{ if(State.special.rows.length < 5){ State.special.rows.push({product:"",placement:"",liquidation:"",revenue:0}); renderSpecial(); } });
  $("specialClear").addEventListener("click", ()=>{ State.special.rows=[]; renderSpecial(); });

  $("specialClearAll").addEventListener("click", ()=>{
    State.special.desc="";
    State.special.photos=[];
    State.special.rows=[];
    renderSpecial();
  });

  // basic input saving
  ["mdoName","headquarter","monthSelect","weekSelect","specialDesc"].forEach(id=>{
    $(id).addEventListener("input", saveDraft);
    $(id).addEventListener("change", saveDraft);
  });

  // custom activity type for photo toolbars
  function toggleCustom(selectEl, inputEl){
    if(selectEl.value === "Custom"){
      inputEl.classList.remove("hidden");
    }else{
      inputEl.classList.add("hidden");
      inputEl.value = "";
    }
  }
  $("photoTypeSelect").addEventListener("change", ()=>toggleCustom($("photoTypeSelect"), $("photoCustomType")));
  $("specialPhotoTypeSelect").addEventListener("change", ()=>toggleCustom($("specialPhotoTypeSelect"), $("specialPhotoCustomType")));

  // upload activity photos
  $("photoUpload").addEventListener("change", async (e)=>{
    const files = [...(e.target.files||[])];
    let type = $("photoTypeSelect").value || "Activity";
    if(type === "Custom") type = $("photoCustomType").value || "Custom";

    for(const f of files){
      if(State.activityPhotos.length >= CFG.limits.activityPhotos) break;
      const dataUrl = await compressImage(f);
      State.activityPhotos.push({type, dataUrl});
    }
    e.target.value = "";
    renderActivityPhotos();
    saveDraft();
  });

  // upload special photos
  $("specialPhotoUpload").addEventListener("change", async (e)=>{
    const files = [...(e.target.files||[])];
    let type = $("specialPhotoTypeSelect").value || "Special";
    if(type === "Custom") type = $("specialPhotoCustomType").value || "Custom";

    for(const f of files){
      if(State.special.photos.length >= CFG.limits.specialPhotos) break;
      const dataUrl = await compressImage(f);
      State.special.photos.push({type, dataUrl});
    }
    e.target.value = "";
    renderSpecial();
    saveDraft();
  });
}

async function boot(){
  loadDraft();
  bindStaticDropdowns();

  try{
    await loadMasterExcel();
  }catch(err){
    hideLoading();
    showError(err.message);
    return;
  }

  // fill dropdowns now that master data is loaded
  bindMdoDropdowns();
  applyDraftToUI();

  // render all
  renderNpiTable();
  renderOtherTable();
  renderActivityTable();
  renderActivityPhotos();
  renderNextWeekTable();
  renderActivityPlanTable();
  renderSpecial();

  // PDF hook (pdf.js defines window.generateA4Pdf)
  $("btnPdf").addEventListener("click", async ()=>{
    await window.generateA4Pdf({
      State, Master, CFG,
      moneyINR, getActivityLabel
    });
  });

  hideLoading();
}

boot();
