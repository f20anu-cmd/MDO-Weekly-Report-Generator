const DATA = {
  regions: {},
  npi: {},
  products: {}
};

const npiRows = [];
const otherRows = [];

function money(n){ return Number(n||0).toLocaleString("en-IN"); }

async function loadExcel(){
  const res = await fetch("data/master_data.xlsx");
  const wb = XLSX.read(await res.arrayBuffer(), {type:"array"});

  XLSX.utils.sheet_to_json(wb.Sheets["Region Mapping"])
    .forEach(r=>{
      if(!DATA.regions[r.Region]) DATA.regions[r.Region]=[];
      DATA.regions[r.Region].push(r.Territtory);
    });

  XLSX.utils.sheet_to_json(wb.Sheets["NPI sheet"] || wb.Sheets["NPI Sheet"])
    .forEach(r=>{
      DATA.npi[r.Product] = { incentive: Number(r.Incentive) };
    });

  XLSX.utils.sheet_to_json(wb.Sheets["Product List"])
    .forEach(r=>{
      DATA.products[r.Product] = Number(r["Realised Value"]);
    });

  initUI();
}

function initUI(){
  Object.keys(DATA.regions).forEach(r=>region.add(new Option(r,r)));
  region.onchange=()=>{
    territory.innerHTML="";
    DATA.regions[region.value].forEach(t=>territory.add(new Option(t,t)));
  };

  ["January","February","March","April","May","June","July","August","September","October","November","December"]
    .forEach(m=>month.add(new Option(m,m)));

  ["1","2","3","4","5"].forEach(w=>week.add(new Option(w,w)));

  npiAdd.onclick=()=>{ npiRows.push({}); renderNpi(); };
  npiClear.onclick=()=>{ npiRows.length=0; renderNpi(); };

  otherAdd.onclick=()=>{ otherRows.push({}); renderOther(); };
  otherClear.onclick=()=>{ otherRows.length=0; renderOther(); };

  generatePdf.onclick=()=>generateA4Pdf({
    mdoName:mdoName.value,
    hq:hq.value,
    region:region.value,
    territory:territory.value,
    month:month.value,
    week:week.value,
    npiRows,
    otherRows
  });
}

function renderNpi(){
  let total=0;
  npiTable.tBodies[0].innerHTML="";
  npiRows.forEach((r,i)=>{
    const inc=(r.act||0)*(DATA.npi[r.p]?.incentive||0);
    total+=inc;
    npiTable.tBodies[0].innerHTML+=`
<tr>
<td>${i+1}</td>
<td><select onchange="npiRows[${i}].p=this.value;renderNpi()">
<option></option>${Object.keys(DATA.npi).map(p=>`<option>${p}</option>`).join("")}
</select></td>
<td><input type=number oninput="npiRows[${i}].plan=this.value"></td>
<td><input type=number oninput="npiRows[${i}].act=this.value;renderNpi()"></td>
<td>${money(inc)}</td>
<td><button onclick="npiRows.splice(${i},1);renderNpi()">X</button></td>
</tr>`;
  });
  npiTotal.innerText=money(total);
}

function renderOther(){
  let total=0;
  otherTable.tBodies[0].innerHTML="";
  otherRows.forEach((r,i)=>{
    const rev=(r.act||0)*(DATA.products[r.p]||0);
    total+=rev;
    otherTable.tBodies[0].innerHTML+=`
<tr>
<td>${i+1}</td>
<td><select onchange="otherRows[${i}].p=this.value;renderOther()">
<option></option>${Object.keys(DATA.products).map(p=>`<option>${p}</option>`).join("")}
</select></td>
<td><input type=number></td>
<td><input type=number oninput="otherRows[${i}].act=this.value;renderOther()"></td>
<td>${money(rev)}</td>
<td><button onclick="otherRows.splice(${i},1);renderOther()">X</button></td>
</tr>`;
  });
  otherTotal.innerText=money(total);
}

document.addEventListener("DOMContentLoaded", loadExcel);
