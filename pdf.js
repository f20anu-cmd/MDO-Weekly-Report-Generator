/* Complete pdf.js (single flowing PDF, no forced page break per section)
   Exact messages preserved.
   Next Week Plan uses PLAN only (no Actual column).
   ₹ formatting fixed (no overflow / spacing issues).
*/
function formatRs(value){
  return "₹ " + Number(value || 0).toLocaleString("en-IN");
}

(function(){
  const { jsPDF } = window.jspdf;

  function ensureSpace(doc, y, needed, pageState){
    if(y + needed <= 280) return y;
    doc.addPage();
    pageState.page += 1;
    header(doc, pageState);
    return 28;
  }

  function header(doc, pageState){
    doc.setFont("helvetica","bold");
    doc.setFontSize(12);
    doc.text("Performance Report", 15, 14);

    doc.setFont("helvetica","normal");
    doc.setFontSize(9);
    doc.text(`Page ${pageState.page}`, 195, 14, {align:"right"});

    doc.setDrawColor(200);
    doc.line(15, 18, 195, 18);
  }

  function sectionTitle(doc, title, y, pageState){
    y = ensureSpace(doc, y, 14, pageState);
    doc.setFont("helvetica","bold");
    doc.setFontSize(11);
    doc.text(title, 15, y);
    doc.setFont("helvetica","normal");
    return y + 6;
  }

  function highlightBox(doc, left, right, y, pageState, style){
    y = ensureSpace(doc, y, 18, pageState);

    if(style === "success") doc.setFillColor(233,255,245);
    else if(style === "danger") doc.setFillColor(255,232,234);
    else doc.setFillColor(224,242,254);

    doc.roundedRect(15, y, 180, 14, 3, 3, "F");
    doc.setFont("helvetica","bold");
    doc.setFontSize(10);

    doc.text(left, 18, y+9);
    doc.text(right, 192, y+9, {align:"right"});

    doc.setFont("helvetica","normal");
    return y + 18;
  }

  function autoTable(doc, head, body, startY){
    doc.autoTable({
      head: [head],
      body,
      startY,
      margin: { left: 15, right: 15 },
      styles: { fontSize: 9, cellPadding: 3, valign:"middle" },
      headStyles: { fillColor:[13,116,200], textColor:255, fontStyle:"bold" },
      alternateRowStyles:{ fillColor:[250,250,250] }
    });
    return doc.lastAutoTable.finalY + 8;
  }

 function addPhotos(doc, title, photos, y, pageState){
  y = sectionTitle(doc, title, y, pageState);

  if(!photos.length){
    y = ensureSpace(doc, y, 10, pageState);
    doc.setFontSize(9);
    doc.text("No photos uploaded.", 15, y);
    return y + 12;
  }

  const x0 = 15;
  const imgW = 80;
  const imgH = 55;
  const gapX = 10;
  const gapY = 22;

  let x = x0;
  let col = 0;

  for(const p of photos){
    y = ensureSpace(doc, y, imgH + gapY, pageState);

    try{
      doc.addImage(p.dataUrl, "JPEG", x, y, imgW, imgH);
    }catch(e){
      // ignore invalid images
    }

    doc.setFont("helvetica","bold");
    doc.setFontSize(9);
    doc.setTextColor(90);
    doc.text(String(p.activity || "Activity"), x, y + imgH + 6);
    doc.setTextColor(0);
    doc.setFont("helvetica","normal");

    col++;

    if(col === 2){
      // completed a full row
      col = 0;
      x = x0;
      y += imgH + gapY;
    }else{
      // move to second column
      x += imgW + gapX;
    }
  }

  // ✅ CRITICAL FIX:
  // If the last row had only ONE image, move Y down
  if(col === 1){
    y += imgH + gapY;
  }

  // extra breathing room before next section
  return y + 6;
}


  window.generatePerformancePdf = function(payload, Master){
    const doc = new jsPDF({unit:"mm", format:"a4"});
    const pageState = { page: 1 };

    header(doc, pageState);
    let y = 28;

    /* 1) MDO */
    y = sectionTitle(doc, "1) MDO Information", y, pageState);
    const info = [
      ["Name", payload.mdo?.name],
      ["HQ", payload.mdo?.hq],
      ["Region", payload.mdo?.region],
      ["Territory", payload.mdo?.territory],
      ["Month", payload.mdo?.month],
      ["Week", payload.mdo?.week]
    ];
    doc.setFontSize(10);
    for(const [k,v] of info){
      y = ensureSpace(doc, y, 8, pageState);
      doc.setFont("helvetica","bold");
      doc.text(`${k}:`, 15, y);
      doc.setFont("helvetica","normal");
      doc.text(String(v || ""), 55, y);
      y += 7;
    }

    /* 2) NPI */
    y = sectionTitle(doc, "2) NPI Performance Update", y, pageState);
    let npiOpp=0, npiEarn=0;

    const npiBody = (payload.npiRows||[]).map((r,i)=>{
      const rate = Master.npiMeta.get(r.product)?.incentiveRate || 0;
      const plan = Number(r.plan||0);
      const actual = Number(r.actual||0);
      const opp = plan * rate;
      const earn = actual * rate;
      npiOpp += opp; npiEarn += earn;

      return [
        i+1, r.product||"", r.plan||"", r.actual||"",
        formatRs(opp), formatRs(earn)
      ];
    });

    y = autoTable(doc,
      ["#","Product","Plan (L/Kg)","Actual (L/Kg)","Total Incentive Opportunity","Total Incentive Earned"],
      npiBody, y
    );

    y = highlightBox(doc,
      "CONGRATULATIONS YOU HAVE EARNED",
      `${formatRs(npiEarn)} !!!`,
      y, pageState, "success"
    );
    y = highlightBox(doc,
      "YOU LOSE",
      formatRs(Math.max(0, npiOpp-npiEarn)),
      y, pageState, "danger"
    );

    /* 3) Other Products */
    y = sectionTitle(doc, "3) Other Product Performance Update", y, pageState);
    let otherTotal=0;

    const otherBody=(payload.otherRows||[]).map((r,i)=>{
      const realised=Master.productMeta.get(r.product)?.realised||0;
      const actual=Number(r.actual||0);
      const rev=actual*realised;
      otherTotal+=rev;
      return [i+1,r.product||"",r.plan||"",r.actual||"",formatRs(rev)];
    });

    y = autoTable(doc,
      ["#","Product","Plan (L/Kg)","Actual (L/Kg)","Total Revenue"],
      otherBody, y
    );

    y = highlightBox(doc,"TOTAL REVENUE EARNED",formatRs(otherTotal),y,pageState,"info");

    /* 4) Activities */
    y = sectionTitle(doc, "4) Activities Update", y, pageState);
    const actBody=(payload.activityRows||[]).map((r,i)=>[
      i+1,r.activity||"",r.planNo||"",r.actualNo||"",r.npiNo||""
    ]);
    y = autoTable(doc,
      ["#","Activity","Plan No","Actual No","NPI Focused Activity No"],
      actBody, y
    );

    /* 5) Photos */
    y = addPhotos(doc,"5) Activities Photos",payload.photos||[],y,pageState);

    /* 6) Next Week – PLAN ONLY */
    y = sectionTitle(doc, "6) Next Week Plan – Product Plan", y, pageState);
    let nwRev=0,nwOpp=0;

    const nwBody=(payload.nextWeekRows||[]).map((r,i)=>{
      const plan=Number(r.plan||0);
      const realised=Master.productMeta.get(r.product)?.realised||0;
      const rate=Master.npiMeta.get(r.product)?.incentiveRate||0;
      const rev=plan*realised;
      const opp=plan*rate;
      nwRev+=rev; nwOpp+=opp;
      return [i+1,r.product||"",r.plan||"",formatRs(rev),formatRs(opp)];
    });

    y = autoTable(doc,
      ["#","Product","Plan (L/Kg)","Total Revenue","Total Incentive Earned"],
      nwBody, y
    );

    y = highlightBox(doc,
      "YOUR NEXT WEEK INCENTIVE OPPORTUNITY",
      `${formatRs(nwOpp)} !!!`,
      y,pageState,"success"
    );
    y = highlightBox(doc,"TOTAL REVENUE",formatRs(nwRev),y,pageState,"info");

    /* 7) Activities Plan */
    y = sectionTitle(doc, "7) Activities Plan", y, pageState);
    const planBody=(payload.actPlanRows||[]).map((r,i)=>{
      const villages=(r.villages||"").split(",").filter(Boolean);
      return [i+1,r.activity||"",r.planNo||"",villages.join(", "),villages.length];
    });
    y = autoTable(doc,
      ["#","Activity","Plan No","Village Names","Village No"],
      planBody,y
    );

    /* 8) Special Achievement */
    y = sectionTitle(doc, "8) Special Achievement", y, pageState);
    doc.setFontSize(9);
    doc.text(doc.splitTextToSize(payload.spDesc||"—",180),15,y);

    y = addPhotos(doc,"Special Achievement Photos",payload.spPhotos||[],y+10,pageState);

    const safeName=(payload.mdo?.name||"Report").replace(/[^\w]+/g,"_");
    doc.save(`Performance_Report_${safeName}.pdf`);
  };
})();
