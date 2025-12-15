(function(){
  const { jsPDF } = window.jspdf;

  function nowDate(){
    const d = new Date();
    return d.toLocaleDateString("en-GB");
  }

  function header(doc, pageNo){
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("Performance Report", 15, 14);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text(`Page ${pageNo}`, 195, 14, {align:"right"});

    doc.setDrawColor(200);
    doc.line(15, 18, 195, 18);
  }

  function footer(doc){
    doc.setFontSize(8);
    doc.setTextColor(120);
    doc.text(`Generated on ${nowDate()}`, 15, 290);
    doc.setTextColor(0);
  }

  function sectionTitle(doc, t, y){
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text(t, 15, y);
    doc.setFont("helvetica", "normal");
    return y + 6;
  }

  function ensureSpace(doc, y, pageNo, needed=18){
    if(y + needed > 280){
      doc.addPage();
      pageNo += 1;
      header(doc, pageNo);
      footer(doc);
      y = 28;
    }
    return { y, pageNo };
  }

  function autoTable(doc, head, body, startY){
    doc.autoTable({
      head: [head],
      body,
      startY,
      margin: { left: 15, right: 15 },
      styles: { fontSize: 9, cellPadding: 3, valign: "middle" },
      headStyles: { fillColor: [13,116,200], textColor: 255, fontStyle:"bold" },
      alternateRowStyles: { fillColor: [250,250,250] }
    });
    return doc.lastAutoTable.finalY + 8;
  }

  function drawHighlight(doc, y, leftText, rightText){
    doc.setFillColor(224, 242, 254);
    doc.roundedRect(15, y, 180, 14, 3, 3, "F");

    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text(leftText, 18, y+9);
    doc.text(rightText, 190, y+9, {align:"right"});
    doc.setFont("helvetica","normal");

    return y + 18;
  }

  function addPhotosFlow(doc, title, photos, y, pageNo){
    if(!photos.length){
      ({y, pageNo} = ensureSpace(doc, y, pageNo, 18));
      y = sectionTitle(doc, title, y);
      doc.setFontSize(10);
      doc.text("No photos uploaded.", 15, y+6);
      return { y: y+18, pageNo };
    }

    ({y, pageNo} = ensureSpace(doc, y, pageNo, 18));
    y = sectionTitle(doc, title, y);

    const x0=15, w=86, h=58, gapX=8, gapY=16;
    let x=x0, col=0;

    // start photos a bit below title
    y += 2;

    for(const p of photos){
      // ensure enough space for one photo block
      ({y, pageNo} = ensureSpace(doc, y, pageNo, h + 14));

      try{ doc.addImage(p.dataUrl, "JPEG", x, y, w, h); }catch(e){}

      doc.setFont("helvetica","bold");
      doc.setFontSize(9);
      doc.setTextColor(90);
      doc.text((p.activity || "Activity").toString(), x, y+h+6);
      doc.setTextColor(0);
      doc.setFont("helvetica","normal");

      col++;
      if(col===2){
        col=0; x=x0; y += h + gapY;
      }else{
        x += w + gapX;
      }
    }

    // reset x and move y a bit
    y += 8;
    return { y, pageNo };
  }

  window.generateA4Pdf = function(payload){
    const { State, typeLabel, rs } = payload;

    const doc = new jsPDF({unit:"mm", format:"a4", orientation:"portrait"});
    let pageNo = 1;

    header(doc, pageNo);
    footer(doc);

    let y = 28;

    // 1) MDO Details
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 30));
    y = sectionTitle(doc, "1) MDO Information", y);

    doc.setFontSize(10);
    doc.setFont("helvetica","normal");
    const items = [
      ["Name", State.mdoName || ""],
      ["HQ", State.hq || ""],
      ["Region", State.region || ""],
      ["Territory", State.territory || ""],
      ["Month", State.month || ""],
      ["Week", State.week ? `Week ${State.week}` : ""],
    ];
    for(const [k,v] of items){
      ({y, pageNo} = ensureSpace(doc, y, pageNo, 8));
      doc.setFont("helvetica","bold");
      doc.text(`${k}:`, 15, y);
      doc.setFont("helvetica","normal");
      doc.text(String(v||""), 55, y);
      y += 7;
    }
    y += 4;

    // 2) NPI
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 20));
    y = sectionTitle(doc, "2) NPI Performance Update", y);

    let npiTotal = 0;
    const npiBody = (State.npiRows||[]).map((r,i)=>{
      npiTotal += (r.incentiveEarned || 0);
      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        `${rs(r.incentiveEarned||0)} Rs`
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Actual (L/Kg)", "Total Incentive Earned"],
      npiBody,
      y
    );

    ({y, pageNo} = ensureSpace(doc, y, pageNo, 18));
    y = drawHighlight(doc, y, "Congratulations you have earned", `${rs(npiTotal)} Rs !!!`);

    // 3) Other products
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 20));
    y = sectionTitle(doc, "3) Other Product Performance Update", y);

    let otherTotal = 0;
    const otherBody = (State.otherRows||[]).map((r,i)=>{
      otherTotal += (r.revenue || 0);
      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        `${rs(r.revenue||0)} Rs`
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Actual (L/Kg)", "Total Revenue"],
      otherBody,
      y
    );

    ({y, pageNo} = ensureSpace(doc, y, pageNo, 18));
    y = drawHighlight(doc, y, "TOTAL REVENUE EARNED", `${rs(otherTotal)} Rs`);

    // 4) Activities Update
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 20));
    y = sectionTitle(doc, "4) Activities Update", y);

    let tp=0, ta=0, tn=0;
    const actBody = (State.actRows||[]).map((r,i)=>{
      tp += Number(r.planNo||0);
      ta += Number(r.actualNo||0);
      tn += Number(r.npiNo||0);
      return [
        String(i+1),
        typeLabel(r.typeObj) || "",
        String(r.planNo || ""),
        String(r.actualNo || ""),
        String(r.npiNo || "")
      ];
    });

    y = autoTable(doc,
      ["#", "Activity", "Plan No", "Actual No", "NPI Focused"],
      actBody,
      y
    );

    ({y, pageNo} = ensureSpace(doc, y, pageNo, 10));
    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text(`TOTAL  Plan: ${tp}   |   Actual: ${ta}   |   NPI Focused: ${tn}`, 15, y);
    doc.setFont("helvetica","normal");
    y += 10;

    // 5) Activities Photos (flow, adds new pages only if needed)
    const actPhotos = (State.photoRows||[]).filter(p=>p.dataUrl).slice(0,16);
    ({y, pageNo} = addPhotosFlow(doc, "5) Activities Photos", actPhotos, y, pageNo));

    // 6) Next Week Plan
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 20));
    y = sectionTitle(doc, "6) Next Week Plan - Product Plan", y);

    let nwRev=0, nwInc=0;
    const nwBody = (State.nwRows||[]).map((r,i)=>{
      nwRev += (r.revenue||0);
      nwInc += (r.incentive||0);
      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        `${rs(r.revenue||0)} Rs`,
        `${rs(r.incentive||0)} Rs`
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Actual (L/Kg)", "Total Revenue", "Total Incentive Earned"],
      nwBody,
      y
    );

    ({y, pageNo} = ensureSpace(doc, y, pageNo, 36));
    y = drawHighlight(doc, y, "TOTAL REVENUE", `${rs(nwRev)} Rs`);
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 18));
    y = drawHighlight(doc, y, "Your next week incentive opportunity is", `${rs(nwInc)} Rs !!!`);

    // 7) Activities Plan
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 20));
    y = sectionTitle(doc, "7) Activities Plan", y);

    const apBody = (State.apRows||[]).map((r,i)=>{
      const vCount = (r.villageNo ?? 0);
      return [
        String(i+1),
        typeLabel(r.typeObj) || "",
        String(r.planNo || ""),
        String(r.villages || ""),
        String(vCount)
      ];
    });

    y = autoTable(doc,
      ["#", "Activity", "Plan No", "Village Names", "Village No"],
      apBody,
      y
    );

    // 8) Special Achievement
    ({y, pageNo} = ensureSpace(doc, y, pageNo, 20));
    y = sectionTitle(doc, "8) Special Achievement", y);

    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text("Description:", 15, y);
    y += 6;

    doc.setFont("helvetica","normal");
    doc.setFontSize(9);
    const desc = (State.spDesc || "—").trim();
    const wrapped = doc.splitTextToSize(desc, 180);
    ({y, pageNo} = ensureSpace(doc, y, pageNo, wrapped.length * 5 + 10));
    doc.text(wrapped, 15, y);
    y += (wrapped.length * 5) + 6;

    // Special photos (up to 4)
    const spPhotos = (State.spPhotoRows||[]).filter(p=>p.dataUrl).slice(0,4);
    ({y, pageNo} = addPhotosFlow(doc, "Special Achievement Photos", spPhotos, y, pageNo));

    const safeName = (State.mdoName || "Report").replace(/[^\w]+/g, "_");
    doc.save(`Performance_Report_${safeName}.pdf`);
  };
})();
