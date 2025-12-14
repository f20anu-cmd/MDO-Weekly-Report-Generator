const DATA = { regions:{}, npi:{}, products:{} };
const S = { npi:[], other:[], act:[], next:[], special:[] };

async function loadExcel(){
  const res = await fetch("data/master_data.xlsx");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf,{type:"array"});

  XLSX.utils.sheet_to_json(wb.Sheets["Region Mapping"])
    .forEach(r=>{
      if(!DATA.regions[r.Region]) DATA.regions[r.Region]=[];
      DATA.regions[r.Region].push(r.Territtory);
    });

  XLSX.utils.sheet_to_json(wb.Sheets["NPI Sheet"])
    .forEach(r=>{
      DATA.npi[r.Product]={ rv:r["Realised Value in Rs"], inc:r.Incentive };
    });

  XLSX.utils.sheet_to_json(wb.Sheets["Product List"])
    .forEach(r=>{
      DATA.products[r.Product]={ rv:r["Realised Value"], cat:r.Category };
    });

  initUI();
}

function initUI(){
  region.innerHTML="<option>Select Region</option>";
  Object.keys(DATA.regions).forEach(r=>region.innerHTML+=`<option>${r}</option>`);
  region.onchange=()=>territory.innerHTML=DATA.regions[region.value].map(t=>`<option>${t}</option>`);

  ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    .forEach(m=>month.innerHTML+=`<option>${m}</option>`);
  ["1","2","3","4","5"].forEach(w=>week.innerHTML+=`<option>${w}</option>`);
}

/* ---------- NPI ---------- */
function addNpi(){ S.npi.push({}); renderNpi(); }
function clearNpi(){ S.npi=[]; renderNpi(); }
function renderNpi(){
  let rev=0,inc=0;
  npiTable.tBodies[0].innerHTML="";
  S.npi.forEach((r,i)=>{
    const m=DATA.npi[r.p]||{rv:0,inc:0};
    const a=r.a||0;
    const tr=a*m.rv, ti=a*m.inc;
    rev+=tr; inc+=ti;
    npiTable.tBodies[0].innerHTML+=`
<tr><td>${i+1}</td>
<td><select onchange="S.npi[${i}].p=this.value;renderNpi()">
<option></option>${Object.keys(DATA.npi).map(p=>`<option ${p==r.p?"selected":""}>${p}</option>`).join("")}
</select></td>
<td><input type=number oninput="S.npi[${i}].pl=this.value"/></td>
<td><input type=number oninput="S.npi[${i}].a=this.value;renderNpi()"/></td>
<td>₹ ${tr.toLocaleString("en-IN")}</td>
<td>₹ ${ti.toLocaleString("en-IN")}</td>
<td><button onclick="S.npi.splice(${i},1);renderNpi()">✕</button></td></tr>`;
  });
  npiRev.innerText="₹ "+rev.toLocaleString("en-IN");
  npiInc.innerText="₹ "+inc.toLocaleString("en-IN");
}

/* ---------- OTHER ---------- */
function addOther(){ S.other.push({}); renderOther(); }
function clearOther(){ S.other=[]; renderOther(); }
function renderOther(){
  let rev=0;
  otherTable.tBodies[0].innerHTML="";
  S.other.forEach((r,i)=>{
    const m=DATA.products[r.p]||{rv:0};
    const tr=(r.a||0)*m.rv;
    rev+=tr;
    otherTable.tBodies[0].innerHTML+=`
<tr><td>${i+1}</td>
<td><select onchange="S.other[${i}].p=this.value;renderOther()">
<option></option>${Object.keys(DATA.products).map(p=>`<option>${p}</option>`).join("")}
</select></td>
<td><input type=number oninput="S.other[${i}].pl=this.value"/></td>
<td><input type=number oninput="S.other[${i}].a=this.value;renderOther()"/></td>
<td>₹ ${tr.toLocaleString("en-IN")}</td>
<td><button onclick="S.other.splice(${i},1);renderOther()">✕</button></td></tr>`;
  });
  otherRev.innerText="₹ "+rev.toLocaleString("en-IN");
}

window.onload=loadExcel;
