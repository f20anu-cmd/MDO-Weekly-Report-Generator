let npiMaster = {};
let productMaster = {};
let photos = [];
let logo64 = null;

/* ---------------- UTIL ---------------- */
function sanitize(t) {
  return t.trim().replace(/\s+/g, "_").replace(/[^a-zA-Z0-9_]/g, "");
}

/* ---------------- LOAD LOGO ---------------- */
function loadLogo(cb) {
  const img = new Image();
  img.src = "fmc_logo.jpg";
  img.onload = () => {
    const c = document.createElement("canvas");
    c.width = img.width; c.height = img.height;
    c.getContext("2d").drawImage(img,0,0);
    logo64 = c.toDataURL("image/jpeg");
    cb();
  };
}

/* ---------------- LOAD EXCEL ---------------- */
excelInput.onchange = e => {
  const r = new FileReader();
  r.onload = ev => {
    const wb = XLSX.read(ev.target.result,{type:"binary"});
    XLSX.utils.sheet_to_json(wb.Sheets["NPI_Master"])
      .forEach(r => npiMaster[r.Product_Name]=r);
    XLSX.utils.sheet_to_json(wb.Sheets["Products_Master"])
      .forEach(r => productMaster[r.Product_Name]=r);
  };
  r.readAsBinaryString(e.target.files[0]);
};

/* ---------------- NPI ROW ---------------- */
function addNPIRow() {
  if (document.querySelectorAll("#npiTable tbody tr").length >= 9) return;
  const tr = document.createElement("tr");
  const sel = document.createElement("select");
  Object.keys(npiMaster).forEach(p=>{
    const o=document.createElement("option");o.text=o.value=p;sel.appendChild(o);
  });
  const plan=document.createElement("input");
  const ach=document.createElement("input");
  const inc=document.createElement("td");
  plan.type=ach.type="number";
  ach.oninput=()=>inc.innerText=(ach.value||0)*npiMaster[sel.value].Incentive_per_Unit;
  tr.append(Object.assign(document.createElement("td"),{appendChild:()=>sel}),plan,ach,inc);
  tr.children[0].appendChild(sel);
  document.querySelector("#npiTable tbody").appendChild(tr);
}

/* ---------------- PRODUCT ROW ---------------- */
function addProductRow() {
  if (document.querySelectorAll("#productTable tbody tr").length >= 10) return;
  const tr=document.createElement("tr");
  const sel=document.createElement("select");
  Object.keys(productMaster).forEach(p=>{
    const o=document.createElement("option");o.text=o.value=p;sel.appendChild(o);
  });
  const plan=document.createElement("input");
  const ach=document.createElement("input");
  const rev=document.createElement("td");
  plan.type=ach.type="number";
  ach.oninput=()=>rev.innerText=(ach.value||0)*productMaster[sel.value].Revenue_per_Unit;
  tr.append(Object.assign(document.createElement("td"),{appendChild:()=>sel}),plan,ach,rev);
  tr.children[0].appendChild(sel);
  document.querySelector("#productTable tbody").appendChild(tr);
}

/* ---------------- PHOTOS ---------------- */
photoInput.onchange=e=>photos=[...e.target.files].slice(0,20);

/* ---------------- PDF ---------------- */
function generatePDF() {
  loadLogo(()=>{
    const {jsPDF}=window.jspdf;
    const doc=new jsPDF();
    const fname=`${sanitize(mdoName.value)}_${sanitize(territory.value)}_${month.value}_${week.value}.pdf`;

    if(logo64) doc.addImage(logo64,"JPEG",10,8,40,15);
    doc.text("MDO WEEKLY PERFORMANCE REPORT",60,18);

    doc.text(`MDO: ${mdoName.value}`,10,30);
    doc.text(`Region: ${region.value}`,10,36);
    doc.text(`Territory: ${territory.value}`,10,42);
    doc.text(`HQ: ${hq.value}`,10,48);
    doc.text(`Month: ${month.value} | ${week.value}`,10,54);

    let npiRows=[], totalInc=0;
    document.querySelectorAll("#npiTable tbody tr").forEach(r=>{
      const v=+r.children[3].innerText||0; totalInc+=v;
      npiRows.push([r.children[0].innerText,r.children[1].value,r.children[2].value,v]);
    });

    doc.autoTable({startY:60,head:[["NPI Product","Plan","Achievement","Incentive (₹)"]],body:npiRows});

    doc.setTextColor(0,128,0);
    doc.text(`🎉 Congratulations! Total NPI Incentive Earned: ₹ ${totalInc.toLocaleString()}`,
      10, doc.lastAutoTable.finalY+10);

    doc.setTextColor(0);
    let prodRows=[], totalRev=0;
    document.querySelectorAll("#productTable tbody tr").forEach(r=>{
      const v=+r.children[3].innerText||0; totalRev+=v;
      prodRows.push([r.children[0].innerText,r.children[1].value,r.children[2].value,v]);
    });

    doc.autoTable({
      startY:doc.lastAutoTable.finalY+20,
      head:[["Product","Plan","Achievement","Revenue (₹)"]],
      body:prodRows
    });

    doc.text(`Total Revenue Generated: ₹ ${(totalRev/100000).toFixed(2)} Lakhs`,
      10, doc.lastAutoTable.finalY+10);

    doc.save(fname);
  });
}
