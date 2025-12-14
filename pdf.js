// pdf.js - A4, section-wise pages, includes tables + photos
(function(){
  const { jsPDF } = window.jspdf;

  function nowDate(){
    const d = new Date();
    return d.toLocaleDateString("en-GB");
  }

  function header(doc, title, pageNo){
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text(title, 15, 14);

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

  function sectionTitle(doc, t){
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text(t, 15, 28);
    doc.setFont("helvetica", "normal");
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

  function drawHighlight(doc, y, leftText, rightText){
    doc.setFillColor(224, 242, 254);
    doc.roundedRect(15, y, 180, 14, 3, 3, "F");

    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text(leftText, 18, y+9);
    doc.text(rightText, 190, y+9, {align:"right"});
    doc.setFont("helvetica","normal");
  }

  function addPhotosGrid(doc, title, photos, ctx){
    // 8 per page (2 cols x 4 rows)
    let idx = 0;

    while(idx < photos.length){
      ctx.newPage(title);

      const slice = photos.slice(idx, idx+8);
      const x0=15, y0=36;
      const w=86, h=58, gapX=8, gapY=16;

      let x=x0, y=y0, col=0;
      for(const p of slice){
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

      footer(doc);
      idx += 8;
    }

    if(photos.length === 0){
      ctx.newPage(title);
      doc.setFontSize(10);
      doc.text("No photos uploaded.", 15, 40);
      footer(doc);
    }
  }

  window.generateA4Pdf = function(payload){
    const { Master, State, typeLabel, rs } = payload;
    const doc = new jsPDF({unit:"mm", format:"a4", orientation:"portrait"});

    let pageNo = 1;

    const ctx = {
      newPage: (section)=>{
        if(pageNo === 1){
          header(doc, "Performance Report", pageNo);
        }else{
          doc.addPage();
          header(doc, "Performance Report", pageNo);
        }
        if(section) sectionTitle(doc, section);
        footer(doc);
        pageNo += 1;
      }
    };

    // Helper to start a new section page (and fix page numbering)
    function startSection(section){
      // first section uses page 1 without addPage
      if(pageNo === 1){
        header(doc, "Performance Report", pageNo);
        sectionTitle(doc, section);
        pageNo += 1;
      }else{
        doc.addPage();
        header(doc, "Performance Report", pageNo);
        sectionTitle(doc, section);
        pageNo += 1;
      }
    }

    // 1) MDO Details
    startSection("1) MDO Details");
    doc.setFontSize(10);
    doc.setFont("helvetica","bold");
    const items = [
      ["Name", State.mdoName || ""],
      ["HQ", State.hq || ""],
      ["Region", State.region || ""],
      ["Territory", State.territory || ""],
      ["Month", State.month || ""],
      ["Week", State.week ? `Week ${State.week}` : ""],
    ];
    let y = 38;
    for(const [k,v] of items){
      doc.text(`${k}:`, 15, y);
      doc.setFont("helvetica","normal");
      doc.text(String(v||""), 55, y);
      doc.setFont("helvetica","bold");
      y += 7;
    }
    footer(doc);

    // 2) NPI Performance
    startSection("2) NPI Performance Update");
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
      34
    );

    drawHighlight(doc, y, "Congratulations you have earned", `${rs(npiTotal)} Rs !!!`);
    footer(doc);

    // 3) Other Product Performance
    startSection("3) Other Product Performance Update");
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
      34
    );
    drawHighlight(doc, y, "TOTAL REVENUE EARNED", `${rs(otherTotal)} Rs`);
    footer(doc);

    // 4) Activities Update
    startSection("4) Activities Update");
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
      34
    );

    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text(`TOTAL  Plan: ${tp}   |   Actual: ${ta}   |   NPI Focused: ${tn}`, 15, y+6);
    doc.setFont("helvetica","normal");
    footer(doc);

    // 5) Activities Photos
    // only embed those with dataUrl; cap at 16
    const actPhotos = (State.photoRows||[]).filter(p=>p.dataUrl).slice(0,16);
    addPhotosGrid(doc, "5) Activities Photos", actPhotos, {
      newPage: (sec)=>{
        doc.addPage();
        header(doc, "Performance Report", pageNo);
        sectionTitle(doc, sec);
        pageNo += 1;
      }
    });

    // 6) Next Week Plan - Product Plan
    doc.addPage();
    header(doc, "Performance Report", pageNo);
    sectionTitle(doc, "6) Next Week Plan - Product Plan");
    pageNo += 1;

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
      34
    );

    drawHighlight(doc, y, "TOTAL REVENUE", `${rs(nwRev)} Rs`);
    drawHighlight(doc, y+18, "Your next week incentive opportunity is", `${rs(nwInc)} Rs !!!`);
    footer(doc);

    // 7) Activities Plan
    doc.addPage();
    header(doc, "Performance Report", pageNo);
    sectionTitle(doc, "7) Activities Plan");
    pageNo += 1;

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

    autoTable(doc,
      ["#", "Activity", "Plan No", "Village Names", "Village No"],
      apBody,
      34
    );
    footer(doc);

    // 8) Special Achievement
    doc.addPage();
    header(doc, "Performance Report", pageNo);
    sectionTitle(doc, "8) Special Achievement");
    pageNo += 1;

    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text("Description:", 15, 38);
    doc.setFont("helvetica","normal");
    doc.setFontSize(9);

    const desc = (State.spDesc || "—").trim();
    const wrapped = doc.splitTextToSize(desc, 180);
    doc.text(wrapped, 15, 44);

    footer(doc);

    // Special photos (up to 4)
    const spPhotos = (State.spPhotoRows||[]).filter(p=>p.dataUrl).slice(0,4);
    addPhotosGrid(doc, "Special Achievement Photos", spPhotos, {
      newPage: (sec)=>{
        doc.addPage();
        header(doc, "Performance Report", pageNo);
        sectionTitle(doc, sec);
        pageNo += 1;
      }
    });

    const safeName = (State.mdoName || "Report").replace(/[^\w]+/g, "_");
    doc.save(`Performance_Report_${safeName}.pdf`);
  };
})();
