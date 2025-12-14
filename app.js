const DATA = { regions:{}, npi:{}, products:{}, activities:[] };

const npi=[], other=[], act=[], nextWeek=[], special=[];

function money(n){ return Number(n||0).toLocaleString("en-IN"); }

async function loadExcel(){
  const wb = XLSX.read(
    await (await fetch("data/master_data.xlsx")).arrayBuffer(),
    {type:"array"}
  );

  XLSX.utils.sheet_to_json(wb.Sheets["Region Mapping"])
    .forEach(r=>{
      if(!DATA.regions[r.Region]) DATA.regions[r.Region]=[];
      DATA.regions[r.Region].push(r.Territtory);
    });

  XLSX.utils.sheet_to_json(wb.Sheets["NPI sheet"]||wb.Sheets["NPI Sheet"])
    .forEach(r=>DATA.npi[r.Product]={
      incentive:+r.Incentive,
      realised:+r["Realised Value in Rs"]
    });

  XLSX.utils.sheet_to_json(wb.Sheets["Product List"])
    .forEach(r=>DATA.products[r.Product]=+r["Realised Value"]);

  XLSX.utils.sheet_to_json(wb.Sheets["Activity List"])
    .forEach(r=>DATA.activities.push(r["Activity Type"]));

  initUI();
}

function initUI(){
  Object.keys(DATA.regions).forEach(r=>region.add(new Option(r,r)));
  region.onchange=()=>{
    territory.innerHTML="";
    DATA.regions[region.value].forEach(t=>territory.add(new Option(t,t)));
  };

  ["January","February","March","April","May","June",
   "July","August","September","October","November","December"]
    .forEach(m=>month.add(new Option(m,m)));
  ["1","2","3","4","5"].forEach(w=>week.add(new Option(w,w)));
}

/* ========== HELPERS ========== */
function sel(arr,cb){
  return `<select onchange="${cb}(this.value)">
    <option></option>${arr.map(v=>`<option>${v}</option>`).join("")}
  </select>`;
}
function num(cb){ return `<input type="number" oninput="${cb}(this.value)">`; }
function del(cb){ return `<button onclick="${cb}()">X</button>`; }
function row(i,...c){ return `<tr><td>${i+1}</td>${c.map(x=>`<td>${x}</td>`).join("")}</tr>`; }

/* ========== NPI ========== */
function addNpi(){ npi.push({}); renderNpi(); }
function clearNpi(){ npi.length=0; renderNpi(); }

function renderNpi(){
  let tot=0; npiTable.tBodies[0].innerHTML="";
  npi.forEach((r,i)=>{
    const inc=(r.a||0)*(DATA.npi[r.p]?.incentive||0);
    tot+=inc;
    npiTable.tBodies[0].innerHTML+=row(
      i,
      sel(Object.keys(DATA.npi),v=>{r.p=v;renderNpi()}),
      num(v=>r.pl=v),
      num(v=>{r.a=v;renderNpi()}),
      money(inc),
      del(()=>{npi.splice(i,1);renderNpi()})
    );
  });
  npiTotal.innerText=money(tot);
}

/* ========== OTHER ========== */
function addOther(){ other.push({}); renderOther(); }
function clearOther(){ other.length=0; renderOther(); }

function renderOther(){
  let tot=0; otherTable.tBodies[0].innerHTML="";
  other.forEach((r,i)=>{
    const rev=(r.a||0)*(DATA.products[r.p]||0);
    tot+=rev;
    otherTable.tBodies[0].innerHTML+=row(
      i,
      sel(Object.keys(DATA.products),v=>{r.p=v;renderOther()}),
      num(v=>r.pl=v),
      num(v=>{r.a=v;renderOther()}),
      money(rev),
      del(()=>{other.splice(i,1);renderOther()})
    );
  });
  otherTotal.innerText=money(tot);
}

/* ========== ACTIVITIES ========== */
function addActivity(){ act.push({}); renderActivity(); }
function clearActivity(){ act.length=0; renderActivity(); }

function renderActivity(){
  let p=0,a=0,n=0;
  activityTable.tBodies[0].innerHTML="";
  act.forEach((r,i)=>{
    p+=+r.p||0; a+=+r.a||0; n+=+r.n||0;
    activityTable.tBodies[0].innerHTML+=row(
      i,
      sel(DATA.activities,v=>r.t=v),
      num(v=>{r.p=v;renderActivity()}),
      num(v=>{r.a=v;renderActivity()}),
      num(v=>{r.n=v;renderActivity()}),
      del(()=>{act.splice(i,1);renderActivity()})
    );
  });
  actPlan.innerText=p; actActual.innerText=a; actNpi.innerText=n;
}

/* ========== NEXT WEEK ========== */
function addNext(){ nextWeek.push({}); renderNext(); }
function clearNext(){ nextWeek.length=0; renderNext(); }

function renderNext(){
  let rev=0,inc=0; nextTable.tBodies[0].innerHTML="";
  nextWeek.forEach((r,i)=>{
    const rv=(r.a||0)*((DATA.products[r.p])||(DATA.npi[r.p]?.realised)||0);
    const iv=(r.a||0)*(DATA.npi[r.p]?.incentive||0);
    rev+=rv; inc+=iv;
    nextTable.tBodies[0].innerHTML+=row(
      i,
      sel(Object.keys({...DATA.products,...DATA.npi}),v=>{r.p=v;renderNext()}),
      num(v=>r.pl=v),
      num(v=>{r.a=v;renderNext()}),
      money(rv),
      money(iv),
      del(()=>{nextWeek.splice(i,1);renderNext()})
    );
  });
  nwRevenue.innerText=money(rev);
  nwIncentive.innerText=money(inc);
}

/* ========== SPECIAL ========== */
function addSpecial(){ special.push({}); renderSpecial(); }
function clearSpecial(){ special.length=0; renderSpecial(); }

function renderSpecial(){
  specialTable.tBodies[0].innerHTML="";
  special.forEach((r,i)=>{
    specialTable.tBodies[0].innerHTML+=row(
      i,
      sel(DATA.activities,v=>r.t=v),
      `<input type="file">`,
      del(()=>{special.splice(i,1);renderSpecial()})
    );
  });
}

document.addEventListener("DOMContentLoaded",loadExcel);
