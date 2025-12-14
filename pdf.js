(function(){
  const { jsPDF } = window.jspdf;

  async function fetchAsBase64(url){
    const res = await fetch(url);
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
    // Ensure these exist in repo for ₹ rendering
    const reg = await fetchAsBase64("assets/fonts/NotoSans-Regular.ttf");
    const bold = await fetchAsBase64("assets/fonts/NotoSans-Bold.ttf");

    doc.addFileToVFS("NotoSans-Regular.ttf", reg);
    doc.addFileToVFS("NotoSans-Bold.ttf", bold);
    doc.addFont("NotoSans-Regular.ttf", "NotoSans", "normal");
    doc.addFont("NotoSans-Bold.ttf", "NotoSans", "bold");
    doc.setFont("NotoSans", "normal");
  }

  function nowDate(){
    const d = new Date();
    const dd = String(d.getDate()).padStart(2,"0");
    const mm = String(d.getMonth()+1).padStart(2,"0");
    const yy = d.getFullYear();
    return `${dd}-${mm}-${yy}`;
  }

  function addHeader(doc, ctx){
    const { State } = ctx;
    const marginX = 15;
    doc.setFont("NotoSans","bold");
    doc.setFontSize(12);
    doc.text("MDO Weekly Report", marginX, 14);

    doc.setFont("NotoSans","normal");
    doc.setFontSize(9);
    const right = 195;
    const mw = `${State.mdo.month || ""} | Week ${State.mdo.week || ""}`.trim();
    doc.text(mw, right, 14, {align:"right"});

    // logo
    try{
      const img = document.querySelector(".logo");
      if(img && img.naturalWidth){
        const c = document.createElement("canvas");
        c.width = img.naturalWidth; c.height = img.naturalHeight;
        const g = c.getContext("2d");
        g.drawImage(img, 0, 0);
        const data = c.toDataURL("image/png");
        doc.addImage(data, "PNG", 15, 18, 16, 16);
      }
    }catch(e){}

    // divider
    doc.setDrawColor(220);
    doc.line(15, 36, 195, 36);
  }

  function addFooter(doc, pageNo){
    doc.setFont("NotoSans","normal");
    doc.setFontSize(8);
    doc.setTextColor(120);
    doc.text(`Generated: ${nowDate()}`, 15, 290);
    doc.text(`Page ${pageNo}`, 195, 290, {align:"right"});
    doc.setTextColor(0);
  }

  function addSectionTitle(doc, title){
    doc.setFont("NotoSans","bold");
    doc.setFontSize(12);
    doc.text(title, 15, 46);
  }

  function autoTable(doc, head, body, startY){
    doc.autoTable({
      head: [head],
      body,
      startY,
      margin: { left: 15, right: 15 },
      styles: { font:"NotoSans", fontSize: 9, cellPadding: 3, valign:"middle" },
      headStyles: { fillColor: [243, 244, 246], textColor: 40, fontStyle:"bold" },
      alternateRowStyles: { fillColor: [250, 250, 250] },
      didParseCell: (data)=>{
        // right-align numeric columns if header contains ₹ or (No.) or (L/Kg)
        const t = String(data.column.raw || "");
        if(data.section === "body"){
          // align by index later if needed
        }
      }
    });
    return doc.lastAutoTable.finalY + 6;
  }

  function moneyINRText(moneyINRFn, n){
    // moneyINRFn returns "₹ X"
    return moneyINRFn(n || 0);
  }

  function drawFlashyNpiSummary(doc, ctx, y, totalInc, totalRev){
    // Pastel box (print-safe)
    doc.setDrawColor(30, 64, 175);
    doc.setFillColor(224, 242, 254); // pastel blue
    doc.roundedRect(15, y, 180, 28, 3, 3, "FD");

    // Accent stripe
    doc.setFillColor(219, 234, 254);
    doc.rect(15, y, 180, 6, "F");

    doc.setFont("NotoSans","bold");
    doc.setFontSize(10);
    doc.setTextColor(11, 47, 107);
    doc.text("NPI PERFORMANCE SUMMARY", 18, y+4.5);

    doc.setFontSize(10);
    doc.setFont("NotoSans","bold");
    doc.text("Total Incentive Earned (₹):", 18, y+14);
    doc.text("Total Revenue Generated (₹):", 18, y+22);

    doc.setFont("NotoSans","bold");
    doc.setTextColor(11, 47, 107);
    doc.text(moneyINRText(ctx.moneyINR, totalInc), 190, y+14, {align:"right"});
    doc.text(moneyINRText(ctx.moneyINR, totalRev), 190, y+22, {align:"right"});

    doc.setTextColor(0);
    return y + 34;
  }

  async function addPhotosPage(doc, ctx, title, photos, maxPerPage){
    addHeader(doc, ctx);
    addSectionTitle(doc, title);

    if(!photos || photos.length === 0){
      doc.setFont("NotoSans","normal");
      doc.setFontSize(10);
      doc.text("No photos uploaded.", 15, 58);
      return;
    }

    const x0 = 15, y0 = 56;
    const w = 86, h = 58, gapX = 8, gapY = 16;

    let x = x0, y = y0, col = 0, onPage = 0;
    doc.setFontSize(9);
    doc.setFont("NotoSans","normal");

    for(let i=0;i<photos.length;i++){
      const p = photos[i];

      if(onPage >= maxPerPage){
        // footer handled by caller
        return i; // remaining index
      }

      try{ doc.addImage(p.dataUrl, "JPEG", x, y, w, h); }catch(e){}
      doc.setFont("NotoSans","bold");
      doc.setTextColor(80);
      doc.text((p.type || "Activity").toString(), x, y+h+6);
      doc.setTextColor(0);
      doc.setFont("NotoSans","normal");

      col++;
      onPage++;
      if(col === 2){
        col = 0;
        x = x0;
        y += h + gapY;
      }else{
        x += w + gapX;
      }
    }
    return photos.length;
  }

  window.generateA4Pdf = async function(ctx){
    const doc = new jsPDF({ unit:"mm", format:"a4", orientation:"portrait" });
    await setupFonts(doc);

    let pageNo = 1;

    function newPage(){
      addFooter(doc, pageNo);
      doc.addPage();
      pageNo += 1;
    }

    const { State, DataStore, CFG, moneyINR, getActivityLabel } = ctx;

    // ===== Page 1: MDO Details =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "1) MDO Details");

    doc.setFont("NotoSans","normal");
    doc.setFontSize(10);

    const kv = [
      ["MDO Name", State.mdo.name || ""],
      ["Headquarter", State.mdo.headquarter || ""],
      ["Region", State.mdo.region || ""],
      ["Territory", State.mdo.territory || ""],
      ["Month", State.mdo.month || ""],
      ["Week", State.mdo.week || ""],
    ];
    let y = 58;
    for(const [k,v] of kv){
      doc.setFont("NotoSans","bold"); doc.text(String(k)+":", 15, y);
      doc.setFont("NotoSans","normal"); doc.text(String(v||""), 60, y);
      y += 7;
    }

    newPage();

    // ===== Page 2: NPI Performance =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "2) NPI Performance");

    let totalNpiRevenue = 0;
    let totalNpiIncentive = 0;

    const npiBody = (State.npiRows || []).map((r, idx)=>{
      const meta = DataStore.npiMap.get(r.product) || { realised:0, incentive:0 };
      const ach = Number(String(r.ach||"").replace(/,/g,"")) || 0;
      const plan = Number(String(r.plan||"").replace(/,/g,"")) || 0;

      const revenue = ach * meta.realised;
      const incEarned = ach * meta.incentive;

      totalNpiRevenue += revenue;
      totalNpiIncentive += incEarned;

      return [
        String(idx+1),
        r.product || "",
        String(plan || ""),
        String(ach || ""),
        moneyINR(revenue).replace("₹ ","₹ "),
        moneyINR(incEarned).replace("₹ ","₹ ")
      ];
    });

    y = autoTable(doc,
      ["#", "Product (NPI)", "Plan (L/Kg)", "Achievement (L/Kg)", "Total Revenue (₹)", "Incentive Earned (₹)"],
      npiBody,
      54
    );

    // Flashy summary box (PDF print-safe)
    y = drawFlashyNpiSummary(doc, ctx, y, totalNpiIncentive, totalNpiRevenue);

    newPage();

    // ===== Page 3: Other Products =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "3) Other Product Performance");

    let totalOtherRevenue = 0;

    const otherBody = (State.otherRows || []).map((r, idx)=>{
      const meta = DataStore.productMap.get(r.product) || { realised:0 };
      const ach = Number(String(r.ach||"").replace(/,/g,"")) || 0;
      const plan = Number(String(r.plan||"").replace(/,/g,"")) || 0;

      const revenue = ach * (meta.realised || 0);
      totalOtherRevenue += revenue;

      return [
        String(idx+1),
        r.product || "",
        String(plan || ""),
        String(ach || ""),
        moneyINR(revenue).replace("₹ ","₹ ")
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Achievement (L/Kg)", "Total Revenue (₹)"],
      otherBody,
      54
    );

    doc.setFont("NotoSans","bold");
    doc.setFontSize(11);
    doc.text("Total Revenue Generated (Other Products):", 15, y+6);
    doc.text(moneyINR(totalOtherRevenue), 195, y+6, {align:"right"});

    newPage();

    // ===== Page 4: Activity Status =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "4) Activity Status Update");

    let tp=0, ta=0, tn=0;

    const actBody = (State.activityRows || []).map((r, idx)=>{
      const plan = Number(String(r.plan||"").replace(/,/g,"")) || 0;
      const ach = Number(String(r.ach||"").replace(/,/g,"")) || 0;
      const npiF = Number(String(r.npiFocused||"").replace(/,/g,"")) || 0;

      tp += plan; ta += ach; tn += npiF;

      return [String(idx+1), getActivityLabel(r.typeObj)||"", String(plan), String(ach), String(npiF)];
    });

    y = autoTable(doc,
      ["#", "Activity Type", "Planned (No.)", "Achieved (No.)", "NPI Focused (No.)"],
      actBody,
      54
    );

    doc.setFont("NotoSans","bold");
    doc.setFontSize(10);
    doc.text(`Total Planned: ${tp}   |   Total Achieved: ${ta}   |   Total NPI Focused: ${tn}`, 15, y+6);

    newPage();

    // ===== Pages 5-6: Photos =====
    const photos = State.photos || [];
    let start = 0;
    for(let p=0; p<2; p++){
      addHeader(doc, ctx);
      addSectionTitle(doc, p===0 ? "5) Activity Photos" : "5) Activity Photos (cont.)");

      const slice = photos.slice(start, start + 8); // 8 per page
      if(slice.length === 0 && p===0){
        doc.setFont("NotoSans","normal");
        doc.setFontSize(10);
        doc.text("No photos uploaded.", 15, 58);
      }else{
        // draw 2 columns
        const x0 = 15, y0 = 56;
        const w = 86, h = 58, gapX = 8, gapY = 16;
        let x = x0, y2 = y0, col = 0;

        for(const ph of slice){
          try{ doc.addImage(ph.dataUrl, "JPEG", x, y2, w, h); }catch(e){}
          doc.setFont("NotoSans","bold");
          doc.setFontSize(9);
          doc.setTextColor(80);
          doc.text((ph.type||"Activity").toString(), x, y2+h+6);
          doc.setTextColor(0);

          col++;
          if(col===2){
            col=0; x=x0; y2 += h + gapY;
          }else{
            x += w + gapX;
          }
        }
      }

      addFooter(doc, pageNo);
      if(p===0){
        doc.addPage(); pageNo += 1;
      }
      start += 8;
    }

    // ===== Page 7: Next Week Product Plan =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "6) Next Week Product Plan");

    let totalNWRevenue = 0;
    const nwBody = (State.nextWeekRows || []).map((r, idx)=>{
      const placement = Number(String(r.placement||"").replace(/,/g,"")) || 0;
      const liquidation = Number(String(r.liquidation||"").replace(/,/g,"")) || 0;

      const pl = DataStore.productMap.get(r.product);
      const npiMeta = DataStore.npiMap.get(r.product);
      const realised = (pl && pl.realised) ? pl.realised : (npiMeta ? npiMeta.realised : 0);

      const revenue = liquidation * realised;
      totalNWRevenue += revenue;

      return [String(idx+1), r.product||"", String(placement), String(liquidation), moneyINR(revenue)];
    });

    y = autoTable(doc,
      ["#", "Product", "Placement Plan (L/Kg)", "Liquidation Plan (L/Kg)", "Total Revenue (₹)"],
      nwBody,
      54
    );

    doc.setFont("NotoSans","bold");
    doc.setFontSize(11);
    doc.text("Total Revenue (Next Week):", 15, y+6);
    doc.text(moneyINR(totalNWRevenue), 195, y+6, {align:"right"});

    newPage();

    // ===== Page 8: Next Week Activity Plan =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "7) Activities Plan for Next Week");

    const apBody = (State.activityPlanRows || []).map((r, idx)=>{
      const planned = Number(String(r.planned||"").replace(/,/g,"")) || 0;
      const villages = String(r.villages||"").split(",").map(s=>s.trim()).filter(Boolean);
      return [String(idx+1), getActivityLabel(r.typeObj)||"", String(planned), String(r.villages||""), String(villages.length)];
    });

    autoTable(doc,
      ["#", "Activity Type", "Planned (No.)", "Village Names", "Village Count"],
      apBody,
      54
    );

    newPage();

    // ===== Page 9: Special Achievement =====
    addHeader(doc, ctx);
    addSectionTitle(doc, "8) Special Achievement");

    doc.setFont("NotoSans","bold");
    doc.setFontSize(10);
    doc.text("Special Achievement Description:", 15, 58);

    doc.setFont("NotoSans","normal");
    doc.setFontSize(9);
    const desc = (State.special.desc || "").trim();
    const wrapped = doc.splitTextToSize(desc || "—", 180);
    doc.text(wrapped, 15, 64);

    // photos (up to 4)
    const spPhotos = State.special.photos || [];
    let yPhoto = Math.min(64 + wrapped.length*4 + 6, 110);
    if(spPhotos.length){
      const x0=15, w=55, h=38, gap=6;
      let x=x0;
      for(let i=0;i<Math.min(4, spPhotos.length); i++){
        const p = spPhotos[i];
        try{ doc.addImage(p.dataUrl, "JPEG", x, yPhoto, w, h); }catch(e){}
        doc.setFont("NotoSans","bold"); doc.setFontSize(8);
        doc.setTextColor(80);
        doc.text((p.type||"Special").toString(), x, yPhoto+h+5);
        doc.setTextColor(0);
        x += w + gap;
        if(i===2){ x=x0; yPhoto += h+16; }
      }
    }else{
      doc.setFont("NotoSans","normal");
      doc.text("No special photos uploaded.", 15, yPhoto);
      yPhoto += 10;
    }

    // table
    let totalSpRevenue = 0;
    const spBody = (State.special.rows || []).map((r, idx)=>{
      const placement = Number(String(r.placement||"").replace(/,/g,"")) || 0;
      const liquidation = Number(String(r.liquidation||"").replace(/,/g,"")) || 0;

      const pl = DataStore.productMap.get(r.product);
      const npiMeta = DataStore.npiMap.get(r.product);
      const realised = (pl && pl.realised) ? pl.realised : (npiMeta ? npiMeta.realised : 0);

      const revenue = liquidation * realised;
      totalSpRevenue += revenue;

      return [String(idx+1), r.product||"", String(placement), String(liquidation), moneyINR(revenue)];
    });

    y = autoTable(doc,
      ["#", "Product", "Placement Qty (L/Kg)", "Liquidation Qty (L/Kg)", "Total Revenue (₹)"],
      spBody,
      Math.max(yPhoto + 10, 140)
    );

    doc.setFont("NotoSans","bold");
    doc.setFontSize(11);
    doc.text("Special Achievement Revenue:", 15, y+6);
    doc.text(moneyINR(totalSpRevenue), 195, y+6, {align:"right"});

    addFooter(doc, pageNo);

    // Save
    const safe = (State.mdo.name || "MDO").replace(/[^\w]+/g,"_");
    doc.save(`MDO_Weekly_Report_${safe}.pdf`);
  };
})();
