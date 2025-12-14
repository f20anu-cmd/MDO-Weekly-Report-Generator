const DATA = {
  regions: {},
  npi: {},
  products: {}
};

const state = {
  npiRows: [],
  otherRows: []
};

async function loadMasterData() {
  const res = await fetch("data/master_data.xlsx");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const regionSheet = XLSX.utils.sheet_to_json(wb.Sheets["Region Mapping"]);
  const npiSheet = XLSX.utils.sheet_to_json(wb.Sheets["NPI Sheet"]);
  const prodSheet = XLSX.utils.sheet_to_json(wb.Sheets["Product List"]);

  regionSheet.forEach(r => {
    if (!DATA.regions[r.Region]) DATA.regions[r.Region] = [];
    DATA.regions[r.Region].push(r.Territtory);
  });

  npiSheet.forEach(r => {
    DATA.npi[r.Product] = {
      realised: r["Realised Value in Rs"],
      incentive: r["Incentive"]
    };
  });

  prodSheet.forEach(r => {
    DATA.products[r.Product] = r["Realised Value"];
  });

  initDropdowns();
}

function initDropdowns() {
  const regionSel = document.getElementById("region");
  regionSel.innerHTML = `<option value="">Select Region</option>`;
  Object.keys(DATA.regions).forEach(r =>
    regionSel.innerHTML += `<option>${r}</option>`
  );

  regionSel.onchange = () => {
    const tSel = document.getElementById("territory");
    tSel.innerHTML = DATA.regions[regionSel.value]
      .map(t => `<option>${t}</option>`).join("");
  };

  ["January","February","March","April","May","June",
   "July","August","September","October","November","December"]
   .forEach(m => month.innerHTML += `<option>${m}</option>`);

  ["1","2","3","4","5"].forEach(w => week.innerHTML += `<option>${w}</option>`);
}

function addNpiRow() {
  state.npiRows.push({ product:"", plan:0, ach:0 });
  renderNpi();
}

function renderNpi() {
  const tbody = document.querySelector("#npiTable tbody");
  tbody.innerHTML = "";

  let totalRev = 0, totalInc = 0;

  state.npiRows.forEach((r,i) => {
    const meta = DATA.npi[r.product] || { realised:0, incentive:0 };
    const rev = r.ach * meta.realised;
    const inc = r.ach * meta.incentive;
    totalRev += rev; totalInc += inc;

    tbody.innerHTML += `
      <tr>
        <td>${i+1}</td>
        <td><select onchange="state.npiRows[${i}].product=this.value;renderNpi()">
          <option></option>
          ${Object.keys(DATA.npi).map(p=>`<option ${p===r.product?"selected":""}>${p}</option>`).join("")}
        </select></td>
        <td><input type="number" oninput="state.npiRows[${i}].plan=this.value"/></td>
        <td><input type="number" oninput="state.npiRows[${i}].ach=this.value;renderNpi()"/></td>
        <td>₹ ${rev.toLocaleString("en-IN")}</td>
        <td>₹ ${inc.toLocaleString("en-IN")}</td>
        <td><button onclick="state.npiRows.splice(${i},1);renderNpi()">✕</button></td>
      </tr>`;
  });

  npiRevenue.textContent = "₹ " + totalRev.toLocaleString("en-IN");
  npiIncentive.textContent = "₹ " + totalInc.toLocaleString("en-IN");
}

function clearNpi(){ state.npiRows=[]; renderNpi(); }

window.onload = loadMasterData;
