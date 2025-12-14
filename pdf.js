(function(){
  const { jsPDF } = window.jspdf;

  async function fetchAsBase64(url){
    const res = await fetch(url);
    if(!res.ok) throw new Error(`Font file missing: ${url}`);
    const buf = await res.arrayBuffer();
    let binary = "";
    const bytes = new Uint8Array(buf);
    const chunk = 0x8000;
    for(let i=0;i<bytes.length;i+=chunk){
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i+chunk));
    }
    return btoa(binary);
  }

  async function setupFonts(doc){
    // optional but recommended for ₹ correctness
    try{
      const reg = await fetchAsBase64("assets/fonts/NotoSans-Regular.ttf");
      const bold = await fetchAsBase64("assets/fonts/NotoSans-Bold.ttf");
      doc.addFileToVFS("NotoSans-Regular.ttf", reg);
      doc.addFileToVFS("NotoSans-Bold.ttf", bold);
      doc.addFont("NotoSans-Regular.ttf", "NotoSans", "normal");
      doc.addFont("NotoSans-Bold.ttf", "NotoSans", "bold");
      doc.setFont("NotoSans", "normal");
      return { ok:true };
    }catch(e){
      // fallback: standard font (₹ may not render correctly on some)
      doc.setFont("helvetica", "normal");
      return { ok:false };
    }
  }

  function nowDate(){
    const d = new Date();
    const dd = String(d.getDate()).padStart(2,"0");
    const mm = String(d.getMonth()+1).padStart(2,"0");
    const yy = d.getFullYear();
    return `${dd}-${mm}-${yy}`;
  }

  function header(doc, ctx){
    const { State } = ctx;
    doc.setFontSize(12);
    doc.setFont(undefined, "bold");
    doc.text("MDO Weekly Report", 15, 14);

    doc.setFont(undefined, "normal");
    doc.setFontSize(9);
    const right = 195;
    const mw = `${State.mdo.month || ""} | Week ${State.mdo.week || ""}`.trim();
    doc.text(mw, right, 14, {align:"right"});

    doc.setDrawColor(220);
    doc.line(15, 18, 195, 18);
  }

  function footer(doc, pageNo){
    doc.setFontSize(8);
    doc.setTextColor(120);
    doc.text(`Generated: ${nowDate()}`, 15, 290);
    doc.text(`Page ${pageNo}`, 195, 290, {align:"right"});
    doc.setTextColor(0);
  }

  function sectionTitle(doc, t){
    doc.setFontSize(12);
    doc.setFont(undefined, "bold");
    doc.text(t, 15, 28);
    doc.setFont(undefined, "normal");
  }

  function autoTable(doc, head, body, startY){
    doc.autoTable({
      head: [head],
      body,
      startY,
      margin: { left: 15, right: 15 },
      styles: { fontSize: 9, cellPadding: 3, valign: "middle" },
      headStyles: { fillColor: [241,245,249], textColor: 40, fontStyle:"bold" },
      alternateRowStyles: { fillColor: [250,250,250] }
    });
    return doc.lastAutoTable.finalY + 6;
  }

  function drawNpiSummary(doc, ctx, y, totalInc, totalRev){
    // print-safe “flashy” pastel block
    doc.setDrawColor(30, 64, 175);
    doc.setFillColor(224, 242, 254);
    doc.roundedRect(15, y, 180, 28, 3, 3, "FD");

    doc.setFillColor(219, 234, 254);
    doc.rect(15, y, 180, 6, "F");

    doc.setFontSize(10);
    doc.setFont(undefined, "bold");
    doc.setTextColor(11, 47, 107);
    doc.text("CONGRATULATIONS", 18, y+4.5);

    doc.setFont(undefined, "bold");
    doc.text("Total Incentive Earned (₹):", 18, y+14);
    doc.text("Total Revenue Generated (₹):", 18, y+22);

    doc.text(ctx.moneyINR(totalInc), 190, y+14, {align:"right"});
    doc.text(ctx.moneyINR(totalRev), 190, y+22, {align:"right"});

    doc.setTextColor(0);
    doc.setFont(undefined, "normal");
    return y + 34;
  }

  function addPhotosGrid(doc, title, photos, pageNoStart, ctx){
    // 8 per page (2 columns * 4 rows)
    let idx = 0;
    let pageNo = pageNoStart;

    while(idx < photos.length || (photos.length === 0 && idx === 0)){
      header(doc, ctx);
      sectionTitle(doc, title + (pageNo === pageNoStart ? "" : " (cont.)"));

      if(photos.length === 0){
        doc.setFontSize(10);
        doc.text("No photos uploaded.", 15, 40);
        footer(doc, pageNo);
        return { pageNo };
      }

      const slice = photos.slice(idx, idx + 8);
      const x0=15, y0=36;
      const w=86, h=58, gapX=8, gapY=16;
      let x=x0, y=y0, col=0;

      for(const p of slice){
        try{ doc.addImage(p.dataUrl, "JPEG", x, y, w, h); }catch(e){}
        doc.setFontSize(9);
        doc.setTextColor(80);
        doc.setFont(undefined, "bold");
        doc.text((p.type||"Activity").toString(), x, y+h+6);
        doc.setTextColor(0);
        doc.setFont(undefined, "normal");

        col++;
        if(col===2){
          col=0; x=x0; y += h + gapY;
        }else{
          x += w + gapX;
        }
      }

      footer(doc, pageNo);
      idx += 8;
      if(idx < photos.length){
        doc.addPage();
        pageNo += 1;
      }
    }
    return { pageNo };
  }

  window.generateA4Pdf = async function(ctx){
    const doc = new jsPDF({ unit:"mm", format:"a4", orientation:"portrait" });
    const fontStatus = await setupFonts(doc);

    // If fonts missing, ₹ may break on some devices
    // Still generating PDF instead of blocking.

    let pageNo = 1;
    function newPage(){
      footer(doc, pageNo);
      doc.addPage();
      pageNo += 1;
    }

    const { State, Master } = ctx;

    // Page 1: MDO Details
    header(doc, ctx);
    sectionTitle(doc, "1) MDO Details");
    doc.setFontSize(10);
    const rows = [
      ["MDO Name", State.mdo.name || ""],
      ["Headquarter", State.mdo.headquarter || ""],
      ["Region", State.mdo.region || ""],
      ["Territory", State.mdo.territory || ""],
      ["Month", State.mdo.month || ""],
      ["Week", State.mdo.week || ""],
    ];
    let y = 38;
    for(const [k,v] of rows){
      doc.setFont(undefined, "bold"); doc.text(`${k}:`, 15, y);
      doc.setFont(undefined, "normal"); doc.text(String(v||""), 60, y);
      y += 7;
    }
    newPage();

    // Page 2: NPI Performance
    header(doc, ctx);
    sectionTitle(doc, "2) NPI Performance");

    let totalNpiRevenue = 0;
    let totalNpiIncentive = 0;

    const npiBody = (State.npiRows||[]).map((r, i)=>{
      const meta = Master.npiMeta.get(r.product) || {realised:0, incentive:0};
      const plan = Number(String(r.plan||"").replace(/,/g,"")) || 0;
      const ach = Number(String(r.ach||"").replace(/,/g,"")) || 0;
      const revenue = ach * meta.realised;
      const inc = ach * meta.incentive;

      totalNpiRevenue += revenue;
      totalNpiIncentive += inc;

      return [
        String(i+1),
        r.product || "",
        String(plan || ""),
        String(ach || ""),
        ctx.moneyINR(revenue),
        ctx.moneyINR(inc),
      ];
    });

    y = autoTable(doc,
      ["#", "Product (NPI)", "Plan (L/Kg)", "Achievement (L/Kg)", "Total Revenue (₹)", "Incentive Earned (₹)"],
      npiBody,
      34
    );

    drawNpiSummary(doc, ctx, y, totalNpiIncentive, totalNpiRevenue);
    newPage();

    // Page 3: Other Product Performance
    header(doc, ctx);
    sectionTitle(doc, "3) Other Product Performance");

    let totalOther = 0;
    const otherBody = (State.otherRows||[]).map((r, i)=>{
      const meta = Master.productMeta.get(r.product) || {realised:0};
      const plan = Number(String(r.plan||"").replace(/,/g,"")) || 0;
      const ach = Number(String(r.ach||"").replace(/,/g,"")) || 0;
      const rev = ach * (meta.realised || 0);
      totalOther += rev;
      return [String(i+1), r.product||"", String(plan||""), String(ach||""), ctx.moneyINR(rev)];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Achievement (L/Kg)", "Total Revenue (₹)"],
      otherBody,
      34
    );

    doc.setFont(undefined, "bold");
    doc.setFontSize(11);
    doc.text("Total Revenue Generated (Other Products):", 15, y+6);
    doc.text(ctx.moneyINR(totalOther), 195, y+6, {align:"right"});
    doc.setFont(undefined, "normal");
    newPage();

    // Page 4: Activity Status
    header(doc, ctx);
    sectionTitle(doc, "4) Activity Status Update");

    let tp=0, ta=0, tn=0;
    const actBody = (State.activityRows||[]).map((r, i)=>{
      const planned = Number(String(r.planned||"").replace(/,/g,"")) || 0;
      const achieved = Number(String(r.achieved||"").replace(/,/g,"")) || 0;
      const npiFocused = Number(String(r.npiFocused||"").replace(/,/g,"")) || 0;
      tp += planned; ta += achieved; tn += npiFocused;
      return [String(i+1), ctx.getActivityLabel(r.typeObj)||"", String(planned), String(achieved), String(npiFocused)];
    });

    y = autoTable(doc,
      ["#", "Activity Type", "Planned", "Achieved", "NPI Focused"],
      actBody,
      34
    );

    doc.setFont(undefined, "bold");
    doc.setFontSize(10);
    doc.text(`Total Planned: ${tp}   |   Total Achieved: ${ta}   |   Total NPI Focused: ${tn}`, 15, y+6);
    doc.setFont(undefined, "normal");
    newPage();

    // Page(s): Activity Photos
    const photosRes = addPhotosGrid(doc, "5) Activity Photos", State.activityPhotos||[], pageNo, ctx);
    pageNo = photosRes.pageNo;
    // if more pages were added by addPhotosGrid, it already added pages internally
    // we need to add a new page for the next section only if we are not already on a new blank page
    doc.addPage(); pageNo += 1;

    // Next Week Product Plan
    header(doc, ctx);
    sectionTitle(doc, "6) Next Week Product Plan");
    let totalNW = 0;
    const nwBody = (State.nextWeekRows||[]).map((r,i)=>{
      const placement = Number(String(r.placement||"").replace(/,/g,"")) || 0;
      const liquidation = Number(String(r.liquidation||"").replace(/,/g,"")) || 0;

      const pl = Master.productMeta.get(r.product);
      const npi = Master.npiMeta.get(r.product);
      const realised = (pl && pl.realised) ? pl.realised : (npi ? npi.realised : 0);

      const rev = liquidation * realised;
      totalNW += rev;

      return [String(i+1), r.product||"", String(placement), String(liquidation), ctx.moneyINR(rev)];
    });

    y = autoTable(doc,
      ["#", "Product", "Placement Plan", "Liquidation Plan", "Total Revenue (₹)"],
      nwBody,
      34
    );

    doc.setFont(undefined, "bold");
    doc.setFontSize(11);
    doc.text("Total Revenue (Next Week):", 15, y+6);
    doc.text(ctx.moneyINR(totalNW), 195, y+6, {align:"right"});
    doc.setFont(undefined, "normal");
    newPage();

    // Next week activity plan
    header(doc, ctx);
    sectionTitle(doc, "7) Activities Plan for Next Week");
    const apBody = (State.activityPlanRows||[]).map((r,i)=>{
      const planned = Number(String(r.planned||"").replace(/,/g,"")) || 0;
      const villages = String(r.villages||"").split(",").map(s=>s.trim()).filter(Boolean);
      return [String(i+1), ctx.getActivityLabel(r.typeObj)||"", String(planned), String(r.villages||""), String(villages.length)];
    });

    autoTable(doc,
      ["#", "Activity Type", "Planned", "Village Names", "Village Count"],
      apBody,
      34
    );
    newPage();

    // Special achievement
    header(doc, ctx);
    sectionTitle(doc, "8) Special Achievement");

    doc.setFont(undefined, "bold");
    doc.setFontSize(10);
    doc.text("Description:", 15, 38);

    doc.setFont(undefined, "normal");
    doc.setFontSize(9);
    const desc = (State.special.desc || "—").trim();
    const wrapped = doc.splitTextToSize(desc, 180);
    doc.text(wrapped, 15, 44);

    // Special photos (up to 4)
    let yPhoto = 44 + wrapped.length*4 + 6;
    const spPhotos = (State.special.photos||[]).slice(0,4);
    if(spPhotos.length){
      const x0=15, w=55, h=38, gap=6;
      let x=x0;
      for(let i=0;i<spPhotos.length;i++){
        const p = spPhotos[i];
        try{ doc.addImage(p.dataUrl, "JPEG", x, yPhoto, w, h); }catch(e){}
        doc.setFont(undefined, "bold");
        doc.setTextColor(80);
        doc.text((p.type||"Special").toString(), x, yPhoto+h+5);
        doc.setTextColor(0);
        x += w + gap;
        if(i===2){ x=x0; yPhoto += h+16; }
      }
      doc.setFont(undefined, "normal");
    }else{
      doc.text("No special photos uploaded.", 15, yPhoto);
      yPhoto += 10;
    }

    // Special table
    let totalSp = 0;
    const spBody = (State.special.rows||[]).map((r,i)=>{
      const placement = Number(String(r.placement||"").replace(/,/g,"")) || 0;
      const liquidation = Number(String(r.liquidation||"").replace(/,/g,"")) || 0;

      const pl = Master.productMeta.get(r.product);
      const npi = Master.npiMeta.get(r.product);
      const realised = (pl && pl.realised) ? pl.realised : (npi ? npi.realised : 0);

      const rev = liquidation * realised;
      totalSp += rev;

      return [String(i+1), r.product||"", String(placement), String(liquidation), ctx.moneyINR(rev)];
    });

    y = autoTable(doc,
      ["#", "Product", "Placement Qty", "Liquidation Qty", "Total Revenue (₹)"],
      spBody,
      Math.max(yPhoto + 10, 120)
    );

    doc.setFont(undefined, "bold");
    doc.setFontSize(11);
    doc.text("Special Achievement Revenue:", 15, y+6);
    doc.text(ctx.moneyINR(totalSp), 195, y+6, {align:"right"});

    footer(doc, pageNo);

    const nameSafe = (State.mdo.name || "MDO").replace(/[^\w]+/g,"_");
    doc.save(`MDO_Weekly_Report_${nameSafe}.pdf`);
  };
})();
