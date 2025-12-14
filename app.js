/* ========= Excel schema (your file) =========
Sheets: Region Mapping | NPI Sheet | Product List
Columns:
  Region Mapping: Region, Territtory
  NPI Sheet: Product, Realised Value in Rs, Incentive, Category
  Product List: Product, Realised Value, Category
============================================= */

const CFG = {
  months: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  weeks: ["1","2","3","4","5"],
  activityTypes: ["One to One Farmer","Village Farmer Meeting","Field Day","OFM","LFM","Demo","Special Campaign","Custom"],
  limits: { photos: 16, specialPhotos: 4 },
  storageKey: "mdo_report_draft_v2"
};

const $ = (id)=>document.getElementById(id);

const DataStore = {
  loaded: false,
  regionToTerr: new Map(),       // region => territories[]
  npiProducts: [],              // product[]
  npiMap: new Map(),            // product => { realised, incentive }
  otherProducts: [],            // product[] (Category !== NPI)
  productMap: new Map(),        // product => { realised, category }
};

const State = {
  mdo: { name:"", headquarter:"", region:"", territory:"", month:"", week:"" },
  npiRows: [],
  otherRows: [],
  activityRows: [],
  photos: [], // {type, dataUrl}
  nextWeekRows: [],
  activityPlanRows: [],
  special: {
    desc: "",
    photos: [],
    rows: []
  }
};

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

function setOptions(sel, options, placeholder="Select"){
  sel.innerHTML = "";
  const ph = document.createElement("option");
  ph.value = "";
  ph.textContent = placeholder;
  sel.appendChild(ph);
  for(const o of options){
    const op = document.createElement("option");
    op.value = o;
    op.textContent = o;
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
    // shallow merge
    Object.assign(State, parsed);
    State.mdo = Object.assign({name:"",headquarter:"",region:"",territory:"",month:"",week:""}, State.mdo||{});
    State.npiRows = Array.isArray(State.npiRows) ? State.npiRows : [];
    State.otherRows = Array.isArray(State.otherRows) ? State.otherRows : [];
    State.activityRows = Array.isArray(State.activityRows) ? State.activityRows : [];
    State.photos = Array.isArray(State.photos) ? State.photos : [];
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

async function loadExcelFile(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, {type:"array"});

  const rm = XLSX.utils.sheet_to_json(wb.Sheets["Region Mapping"], {defval:""});
  const npi = XLSX.utils.sheet_to_json(wb.Sheets["NPI Sheet"], {defval:""});
  const pl = XLSX.utils.sheet_to_json(wb.Sheets["Product List"], {defval:""});

  // Region mapping
  const map = new Map();
  for(const r of rm){
    const region = String(r["Region"]||"").trim();
    const terr = String(r["Territtory"]||"").trim();
    if(!region || !terr) continue;
    if(!map.has(region)) map.set(region, []);
    map.get(region).push(terr);
  }
  for(const [k, v] of map){
    map.set(k, [...new Set(v)].sort((a,b)=>a.localeCompare(b)));
  }
  DataStore.regionToTerr = map;

  // NPI
  const npiMap = new Map();
  for(const r of npi){
    const p = String(r["Product"]||"").trim();
    if(!p) continue;
    const realised = toNum(r["Realised Value in Rs"]);
    const incentive = toNum(r["Incentive"]);
    npiMap.set(p, { realised, incentive });
  }
  DataStore.npiMap = npiMap;
  DataStore.npiProducts = [...npiMap.keys()].sort((a,b)=>a.localeCompare(b));

  // Product List
  const pMap = new Map();
  const other = [];
  for(const r of pl){
    const p = String(r["Product"]||"").trim();
    if(!p) continue;
    const realised = toNum(r["Realised Value"]);
    const category = String(r["Category"]||"").trim();
    pMap.set(p, { realised, category });

    if(category.toLowerCase() !== "npi") other.push(p);
  }
  DataStore.productMap = pMap;
  DataStore.otherProducts = [...new Set(other)].sort((a,b)=>a.localeCompare(b));

  DataStore.loaded = true;
}

function makeDelBtn(onClick){
  const b = document.createElement("button");
  b.className = "iconBtn";
  b.textContent = "✕";
  b.title = "Remove";
  b.addEventListener("click", onClick);
  return b;
}

function inputNumber(value, onInput, placeholder="0"){
  const i = document.createElement("input");
  i.type = "number";
  i.step = "any";
  i.min = "0";
  i.value = value ?? "";
  i.placeholder = placeholder;
  i.addEventListener("input", ()=>onInput(i.value));
  return i;
}

function selectWithCustom(valueObj, onChange){
  // valueObj = {mode:"preset"|"custom", value:""}
  const wrap = document.createElement("div");
  wrap.style.display = "grid";
  wrap.style.gridTemplateColumns = "1fr";
  wrap.style.gap = "6px";

  const sel = document.createElement("select");
  setOptions(sel, CFG.activityTypes, "Select activity");
  const customInput = document.createElement("input");
  customInput.type = "text";
  customInput.placeholder = "Enter custom activity name";
  customInput.classList.add("hidden");

  if(valueObj?.mode === "custom"){
    sel.value = "Custom";
    customInput.classList.remove("hidden");
    customInput.value = valueObj.value || "";
  }else{
    sel.value = valueObj?.value || "";
  }

  sel.addEventListener("change", ()=>{
    if(sel.value === "Custom"){
      customInput.classList.remove("hidden");
      onChange({mode:"custom", value: customInput.value || ""});
    }else{
      customInput.classList.add("hidden");
      customInput.value = "";
      onChange({mode:"preset", value: sel.value});
    }
  });

  customInput.addEventListener("input", ()=>{
    onChange({mode:"custom", value: customInput.value || ""});
  });

  wrap.appendChild(sel);
  wrap.appendChild(customInput);
  return wrap;
}

function getActivityLabel(obj){
  if(!obj) return "";
  if(obj.mode === "custom") return obj.value || "Custom";
  return obj.value || "";
}

/* ======== Build Tables ======== */

function renderNPITable(){
  const tbody = $("npiTable").querySelector("tbody");
  tbody.innerHTML = "";

  State.npiRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    const td0 = document.createElement("td");
    td0.textContent = String(idx+1);
    tr.appendChild(td0);

    const td1 = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, DataStore.npiProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcAll();
    });
    td1.appendChild(sel);
    tr.appendChild(td1);

    const td2 = document.createElement("td"); td2.className="num";
    td2.appendChild(inputNumber(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(td2);

    const td3 = document.createElement("td"); td3.className="num";
    td3.appendChild(inputNumber(r.ach, v=>{ r.ach=v; recalcAll(); }));
    tr.appendChild(td3);

    const td4 = document.createElement("td"); td4.className="num";
    td4.textContent = moneyINR(r.revenue || 0);
    tr.appendChild(td4);

    const td5 = document.createElement("td"); td5.className="num";
    td5.textContent = moneyINR(r.incentiveEarned || 0);
    tr.appendChild(td5);

    const td6 = document.createElement("td");
    td6.appendChild(makeDelBtn(()=>{
      State.npiRows.splice(idx,1);
      renderNPITable();
      recalcAll();
    }));
    tr.appendChild(td6);

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

    const td1 = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, DataStore.otherProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcAll();
    });
    td1.appendChild(sel);
    tr.appendChild(td1);

    const td2 = document.createElement("td"); td2.className="num";
    td2.appendChild(inputNumber(r.plan, v=>{ r.plan=v; saveDraft(); }));
    tr.appendChild(td2);

    const td3 = document.createElement("td"); td3.className="num";
    td3.appendChild(inputNumber(r.ach, v=>{ r.ach=v; recalcAll(); }));
    tr.appendChild(td3);

    const td4 = document.createElement("td"); td4.className="num";
    td4.textContent = moneyINR(r.revenue || 0);
    tr.appendChild(td4);

    const td5 = document.createElement("td");
    td5.appendChild(makeDelBtn(()=>{
      State.otherRows.splice(idx,1);
      renderOtherTable();
      recalcAll();
    }));
    tr.appendChild(td5);

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

    const td1 = document.createElement("td");
    td1.appendChild(selectWithCustom(r.typeObj, (v)=>{ r.typeObj=v; saveDraft(); }));
    tr.appendChild(td1);

    const td2 = document.createElement("td"); td2.className="num";
    td2.appendChild(inputNumber(r.plan, v=>{ r.plan=v; recalcActivities(); saveDraft(); }, "0"));
    tr.appendChild(td2);

    const td3 = document.createElement("td"); td3.className="num";
    td3.appendChild(inputNumber(r.ach, v=>{ r.ach=v; recalcActivities(); saveDraft(); }, "0"));
    tr.appendChild(td3);

    const td4 = document.createElement("td"); td4.className="num";
    td4.appendChild(inputNumber(r.npiFocused, v=>{ r.npiFocused=v; recalcActivities(); saveDraft(); }, "0"));
    tr.appendChild(td4);

    const td5 = document.createElement("td");
    td5.appendChild(makeDelBtn(()=>{
      State.activityRows.splice(idx,1);
      renderActivityTable();
      recalcActivities();
      saveDraft();
    }));
    tr.appendChild(td5);

    tbody.appendChild(tr);
  });

  recalcActivities();
}

function renderPhotos(){
  $("photoCount").textContent = String(State.photos.length);
  const grid = $("photoGrid");
  grid.innerHTML = "";

  State.photos.forEach((p, idx)=>{
    const card = document.createElement("div");
    card.className = "photoCard";
    card.innerHTML = `
      <img alt="photo"/>
      <div class="photoMeta">
        <div class="photoCaption"></div>
        <button class="iconBtn" title="Remove">✕</button>
      </div>
    `;
    card.querySelector("img").src = p.dataUrl;
    card.querySelector(".photoCaption").textContent = p.type || "Activity";
    card.querySelector("button").addEventListener("click", ()=>{
      State.photos.splice(idx,1);
      renderPhotos();
      saveDraft();
    });
    grid.appendChild(card);
  });
}

function renderNextWeekTable(){
  const tbody = $("nextWeekTable").querySelector("tbody");
  tbody.innerHTML = "";

  const allProducts = [...new Set([
    ...DataStore.otherProducts,
    ...DataStore.npiProducts
  ])].sort((a,b)=>a.localeCompare(b));

  State.nextWeekRows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const td1 = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, allProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcAll();
    });
    td1.appendChild(sel);
    tr.appendChild(td1);

    const td2 = document.createElement("td"); td2.className="num";
    td2.appendChild(inputNumber(r.placement, v=>{ r.placement=v; saveDraft(); }));
    tr.appendChild(td2);

    const td3 = document.createElement("td"); td3.className="num";
    td3.appendChild(inputNumber(r.liquidation, v=>{ r.liquidation=v; recalcAll(); }));
    tr.appendChild(td3);

    const td4 = document.createElement("td"); td4.className="num";
    td4.textContent = moneyINR(r.revenue || 0);
    tr.appendChild(td4);

    const td5 = document.createElement("td");
    td5.appendChild(makeDelBtn(()=>{
      State.nextWeekRows.splice(idx,1);
      renderNextWeekTable();
      recalcAll();
    }));
    tr.appendChild(td5);

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

    const td1 = document.createElement("td");
    td1.appendChild(selectWithCustom(r.typeObj, (v)=>{ r.typeObj=v; saveDraft(); }));
    tr.appendChild(td1);

    const td2 = document.createElement("td"); td2.className="num";
    td2.appendChild(inputNumber(r.planned, v=>{ r.planned=v; saveDraft(); }, "0"));
    tr.appendChild(td2);

    const td3 = document.createElement("td");
    const inpVill = document.createElement("input");
    inpVill.type = "text";
    inpVill.placeholder = "Village1, Village2, Village3";
    inpVill.value = r.villages || "";
    inpVill.addEventListener("input", ()=>{
      r.villages = inpVill.value;
      renderActivityPlanTable(); // refresh village count cell
      saveDraft();
    });
    td3.appendChild(inpVill);
    tr.appendChild(td3);

    const td4 = document.createElement("td"); td4.className="num";
    const villages = String(r.villages||"").split(",").map(s=>s.trim()).filter(Boolean);
    td4.textContent = String(villages.length);
    tr.appendChild(td4);

    const td5 = document.createElement("td");
    td5.appendChild(makeDelBtn(()=>{
      State.activityPlanRows.splice(idx,1);
      renderActivityPlanTable();
      saveDraft();
    }));
    tr.appendChild(td5);

    tbody.appendChild(tr);
  });
}

function renderSpecial(){
  $("specialDesc").value = State.special.desc || "";

  // photos
  $("specialPhotoCount").textContent = String(State.special.photos.length);
  const grid = $("specialPhotoGrid");
  grid.innerHTML = "";
  State.special.photos.forEach((p, idx)=>{
    const card = document.createElement("div");
    card.className = "photoCard";
    card.innerHTML = `
      <img alt="photo"/>
      <div class="photoMeta">
        <div class="photoCaption"></div>
        <button class="iconBtn" title="Remove">✕</button>
      </div>
    `;
    card.querySelector("img").src = p.dataUrl;
    card.querySelector(".photoCaption").textContent = p.type || "Special";
    card.querySelector("button").addEventListener("click", ()=>{
      State.special.photos.splice(idx,1);
      renderSpecial();
      saveDraft();
    });
    grid.appendChild(card);
  });

  // table
  const tbody = $("specialTable").querySelector("tbody");
  tbody.innerHTML = "";

  const allProducts = [...new Set([
    ...DataStore.otherProducts,
    ...DataStore.npiProducts
  ])].sort((a,b)=>a.localeCompare(b));

  State.special.rows.forEach((r, idx)=>{
    const tr = document.createElement("tr");

    tr.appendChild(Object.assign(document.createElement("td"), {textContent:String(idx+1)}));

    const td1 = document.createElement("td");
    const sel = document.createElement("select");
    setOptions(sel, allProducts, "Select product");
    sel.value = r.product || "";
    sel.addEventListener("change", ()=>{
      r.product = sel.value;
      recalcAll();
    });
    td1.appendChild(sel);
    tr.appendChild(td1);

    const td2 = document.createElement("td"); td2.className="num";
    td2.appendChild(inputNumber(r.placement, v=>{ r.placement=v; saveDraft(); }));
    tr.appendChild(td2);

    const td3 = document.createElement("td"); td3.className="num";
    td3.appendChild(inputNumber(r.liquidation, v=>{ r.liquidation=v; recalcAll(); }));
    tr.appendChild(td3);

    const td4 = document.createElement("td"); td4.className="num";
    td4.textContent = moneyINR(r.revenue || 0);
    tr.appendChild(td4);

    const td5 = document.createElement("td");
    td5.appendChild(makeDelBtn(()=>{
      State.special.rows.splice(idx,1);
      renderSpecial();
      recalcAll();
    }));
    tr.appendChild(td5);

    tbody.appendChild(tr);
  });

  recalcAll();
}

/* ======== Calculations ======== */

function recalcActivities(){
  let p=0,a=0,n=0;
  for(const r of State.activityRows){
    p += toNum(r.plan);
    a += toNum(r.ach);
    n += toNum(r.npiFocused);
  }
  $("actPlanTotal").textContent = String(p);
  $("actAchTotal").textContent = String(a);
  $("actNpiTotal").textContent = String(n);
}

function recalcAll(){
  // NPI
  let npiRev = 0;
  let npiInc = 0;
  for(const r of State.npiRows){
    const meta = DataStore.npiMap.get(r.product) || {realised:0,incentive:0};
    const ach = toNum(r.ach);
    const revenue = ach * meta.realised;
    const incentiveEarned = ach * meta.incentive;
    r.revenue = revenue;
    r.incentiveEarned = incentiveEarned;
    npiRev += revenue;
    npiInc += incentiveEarned;
  }
  $("npiTotalRevenue").textContent = moneyINR(npiRev);
  $("npiTotalIncentive").textContent = moneyINR(npiInc);

  // update visible cells quickly by re-rendering table bodies only if needed
  // (simple approach: rerender tables)
  // but avoid infinite loops: only update money cells if tables exist
  const npiTbody = $("npiTable").querySelector("tbody");
  [...npiTbody.rows].forEach((tr, idx)=>{
    const r = State.npiRows[idx];
    if(!r) return;
    tr.children[4].textContent = moneyINR(r.revenue || 0);
    tr.children[5].textContent = moneyINR(r.incentiveEarned || 0);
  });

  // Other products
  let otherRev = 0;
  for(const r of State.otherRows){
    const meta = DataStore.productMap.get(r.product) || {realised:0};
    const ach = toNum(r.ach);
    const revenue = ach * (meta.realised || 0);
    r.revenue = revenue;
    otherRev += revenue;
  }
  $("otherTotalRevenue").textContent = moneyINR(otherRev);

  const otherTbody = $("otherTable").querySelector("tbody");
  [...otherTbody.rows].forEach((tr, idx)=>{
    const r = State.otherRows[idx];
    if(!r) return;
    tr.children[4].textContent = moneyINR(r.revenue || 0);
  });

  // Next week revenue
  let nwRev = 0;
  for(const r of State.nextWeekRows){
    const ach = toNum(r.liquidation);
    // realised value can come from Product List or NPI Sheet
    const pl = DataStore.productMap.get(r.product);
    const npi = DataStore.npiMap.get(r.product);
    const realised = (pl && pl.realised) ? pl.realised : (npi ? npi.realised : 0);
    const revenue = ach * realised;
    r.revenue = revenue;
    nwRev += revenue;
  }
  $("nwTotalRevenue").textContent = moneyINR(nwRev);

  const nwTbody = $("nextWeekTable").querySelector("tbody");
  [...nwTbody.rows].forEach((tr, idx)=>{
    const r = State.nextWeekRows[idx];
    if(!r) return;
    tr.children[4].textContent = moneyINR(r.revenue || 0);
  });

  // Special revenue
  let spRev = 0;
  for(const r of State.special.rows){
    const ach = toNum(r.liquidation);
    const pl = DataStore.productMap.get(r.product);
    const npi = DataStore.npiMap.get(r.product);
    const realised = (pl && pl.realised) ? pl.realised : (npi ? npi.realised : 0);
    const revenue = ach * realised;
    r.revenue = revenue;
    spRev += revenue;
  }
  $("specialTotalRevenue").textContent = moneyINR(spRev);

  const spTbody = $("specialTable").querySelector("tbody");
  [...spTbody.rows].forEach((tr, idx)=>{
    const r = State.special.rows[idx];
    if(!r) return;
    tr.children[4].textContent = moneyINR(r.revenue || 0);
  });

  recalcActivities();
  saveDraft();
}

/* ======== Photos (compression) ======== */
async function compressImage(file, maxW=1280, quality=0.75){
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

/* ======== Wire up UI ======== */

function fillStaticDropdowns(){
  setOptions($("monthSelect"), CFG.months, "Select month");
  setOptions($("weekSelect"), CFG.weeks, "Select week");

  setOptions($("photoTypeSelect"), CFG.activityTypes, "Select activity type");
  setOptions($("specialPhotoTypeSelect"), CFG.activityTypes, "Select activity type");
}

function applyDraftToUI(){
  $("mdoName").value = State.mdo.name || "";
  $("headquarter").value = State.mdo.headquarter || "";
  $("monthSelect").value = State.mdo.month || "";
  $("weekSelect").value = State.mdo.week || "";
  $("specialDesc").value = State.special.desc || "";
}

function fillRegionDropdowns(){
  const regions = [...DataStore.regionToTerr.keys()].sort((a,b)=>a.localeCompare(b));
  setOptions($("regionSelect"), regions, "Select region");

  $("regionSelect").value = State.mdo.region || "";
  const terrs = DataStore.regionToTerr.get($("regionSelect").value) || [];
  setOptions($("territorySelect"), terrs, "Select territory");
  $("territorySelect").value = State.mdo.territory || "";

  $("regionSelect").addEventListener("change", ()=>{
    const terrs2 = DataStore.regionToTerr.get($("regionSelect").value) || [];
    setOptions($("territorySelect"), terrs2, "Select territory");
    $("territorySelect").value = "";
    saveDraft();
  });
  $("territorySelect").addEventListener("change", saveDraft);
}

function wireButtons(){
  $("btnSave").addEventListener("click", saveDraft);
  $("btnClearAll").addEventListener("click", clearAll);

  $("npiAdd").addEventListener("click", ()=>{ State.npiRows.push({product:"",plan:"",ach:"",revenue:0,incentiveEarned:0}); renderNPITable(); });
  $("npiClear").addEventListener("click", ()=>{ State.npiRows = []; renderNPITable(); recalcAll(); });

  $("otherAdd").addEventListener("click", ()=>{ State.otherRows.push({product:"",plan:"",ach:"",revenue:0}); renderOtherTable(); });
  $("otherClear").addEventListener("click", ()=>{ State.otherRows = []; renderOtherTable(); recalcAll(); });

  $("actAdd").addEventListener("click", ()=>{ State.activityRows.push({typeObj:{mode:"preset",value:""},plan:"",ach:"",npiFocused:""}); renderActivityTable(); saveDraft(); });
  $("actClear").addEventListener("click", ()=>{ State.activityRows = []; renderActivityTable(); recalcActivities(); saveDraft(); });

  $("photoClear").addEventListener("click", ()=>{ State.photos = []; renderPhotos(); saveDraft(); });

  $("nwAdd").addEventListener("click", ()=>{ State.nextWeekRows.push({product:"",placement:"",liquidation:"",revenue:0}); renderNextWeekTable(); });
  $("nwClear").addEventListener("click", ()=>{ State.nextWeekRows = []; renderNextWeekTable(); recalcAll(); });

  $("apAdd").addEventListener("click", ()=>{ State.activityPlanRows.push({typeObj:{mode:"preset",value:""},planned:"",villages:""}); renderActivityPlanTable(); saveDraft(); });
  $("apClear").addEventListener("click", ()=>{ State.activityPlanRows = []; renderActivityPlanTable(); saveDraft(); });

  $("specialAdd").addEventListener("click", ()=>{ State.special.rows.push({product:"",placement:"",liquidation:"",revenue:0}); renderSpecial(); });
  $("specialClear").addEventListener("click", ()=>{ State.special.rows = []; renderSpecial(); recalcAll(); });

  $("specialClearAll").addEventListener("click", ()=>{
    State.special.desc = "";
    State.special.photos = [];
    State.special.rows = [];
    renderSpecial();
    recalcAll();
  });

  // Inputs
  ["mdoName","headquarter","monthSelect","weekSelect","specialDesc"].forEach(id=>{
    $(id).addEventListener("input", saveDraft);
    $(id).addEventListener("change", saveDraft);
  });

  // Photo type custom handling (top controls)
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

  // Upload photos
  $("photoUpload").addEventListener("change", async (e)=>{
    const files = [...(e.target.files||[])];
    let t = $("photoTypeSelect").value || "Activity";
    if(t === "Custom") t = ($("photoCustomType").value || "Custom");
    for(const f of files){
      if(State.photos.length >= CFG.limits.photos) break;
      const dataUrl = await compressImage(f);
      State.photos.push({type: t, dataUrl});
    }
    e.target.value = "";
    renderPhotos();
    saveDraft();
  });

  $("specialPhotoUpload").addEventListener("change", async (e)=>{
    const files = [...(e.target.files||[])];
    let t = $("specialPhotoTypeSelect").value || "Special";
    if(t === "Custom") t = ($("specialPhotoCustomType").value || "Custom");
    for(const f of files){
      if(State.special.photos.length >= CFG.limits.specialPhotos) break;
      const dataUrl = await compressImage(f);
      State.special.photos.push({type: t, dataUrl});
    }
    e.target.value = "";
    renderSpecial();
    saveDraft();
  });

  // PDF
  $("btnPdf").addEventListener("click", async ()=>{
    if(!DataStore.loaded){
      alert("Please load Excel first.");
      return;
    }
    await window.generateA4Pdf({State, DataStore, CFG, moneyINR, getActivityLabel});
  });
}

function bootAfterExcel(){
  fillRegionDropdowns();
  applyDraftToUI();

  renderNPITable();
  renderOtherTable();
  renderActivityTable();
  renderPhotos();
  renderNextWeekTable();
  renderActivityPlanTable();
  renderSpecial();

  recalcAll();
}

function init(){
  loadDraft();
  fillStaticDropdowns();
  wireButtons();

  $("excelFile").addEventListener("change", async (e)=>{
    const file = e.target.files?.[0];
    if(!file) return;
    try{
      await loadExcelFile(file);
      bootAfterExcel();
      alert("Excel loaded ✅");
    }catch(err){
      console.error(err);
      alert("Failed to load Excel. Please use the provided file format.\n" + err.message);
    }finally{
      e.target.value = "";
    }
  });
}

init();
