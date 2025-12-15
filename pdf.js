/* Complete pdf.js – FINAL STABLE VERSION
   ✔ No photo overlap
   ✔ Correct Rs formatting
   ✔ No ₹ inside autoTable (prevents digit split bug)
   ✔ Exact wording preserved
*/

function formatRsTable(value){
  return "Rs " + Number(value || 0).toLocaleString("en-IN");
}

function formatRsHighlight(value){
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
      styles: {
        fontSize: 9,
        cellPadding: 3,
        valign: "middle",
        font: "helvetica"
      },
      headStyles: {
        fillColor: [13,116,200],
        textColor: 255,
        fontStyle: "bold"
      },
      alternateRowStyles: { fillColor: [250,250,250] }
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

    const imgW = 80, imgH = 55, gapX = 10, gapY = 22;
    let x = 15, col = 0;

    for(const p of photos){
      y = ensureSpace(doc, y, imgH + gapY, pageState);

      try{
        doc.addImage(p.dataUrl, "JPEG", x, y, imgW, imgH);
      }catch{}

      doc.setFont("helvetica","bold");
      doc.setFontSize(9);
      doc.text(p.activity || "Activity", x, y + imgH + 6);
      doc.setFont("helvetica","normal");

      col++;
      if(col === 2){
        col = 0;
        x = 15;
        y += imgH + gapY;
      }else{
        x += imgW + gapX;
      }
    }

    if(col === 1){
      y += imgH + gapY;
    }

    return y + 6;
  }

  window.generatePerformancePdf = function(payload, Master){
    const doc = new jsPDF({unit:"mm", format:"a4"});
    const pageState = { page: 1 };

    header(doc, pageState);
    let y = 28;

    /* 2) NPI */
    y = sectionTitle(doc, "2) NPI Performance Update", y, pageState);
    let npiOpp = 0, npiEarn = 0;

    const npiBody = (payload.npiRows || []).map((r,i)=>{
      const rate = Master.npiMeta.get(r.product)?.incentiveRate || 0;
      const plan = Number(r.plan || 0);
      const actual = Number(r.actual || 0);
      const opp = plan * rate;
      const earn = actual * rate;
      npiOpp += opp;
      npiEarn += earn;

      return [
        i+1,
        r.product || "",
        r.plan || "",
        r.actual || "",
        formatRsTable(opp),
        formatRsTable(earn)
      ];
    });

    y = autoTable(doc,
      ["#","Product","Plan (L/Kg)","Actual (L/Kg)","Total Incentive Opportunity","Total Incentive Earned"],
      npiBody, y
    );

    y = highlightBox(
      doc,
      "CONGRATULATIONS YOU HAVE EARNED",
      `${formatRsHighlight(npiEarn)} !!!`,
      y, pageState, "success"
    );

    y = highlightBox(
      doc,
      "YOU LOSE",
      formatRsHighlight(Math.max(0, npiOpp - npiEarn)),
      y, pageState, "danger"
    );

    /* 6) NEXT WEEK PLAN (PLAN ONLY) */
    y = sectionTitle(doc, "6) Next Week Plan – Product Plan", y, pageState);
    let nwRev = 0, nwOpp = 0;

    const nwBody = (payload.nextWeekRows || []).map((r,i)=>{
      const plan = Number(r.plan || 0);
      const realised = Master.productMeta.get(r.product)?.realised || 0;
      const rate = Master.npiMeta.get(r.product)?.incentiveRate || 0;
      const rev = plan * realised;
      const opp = plan * rate;
      nwRev += rev;
      nwOpp += opp;

      return [
        i+1,
        r.product || "",
        r.plan || "",
        formatRsTable(rev),
        formatRsTable(opp)
      ];
    });

    y = autoTable(doc,
      ["#","Product","Plan (L/Kg)","Total Revenue","Total Incentive Earned"],
      nwBody, y
    );

    y = highlightBox(
      doc,
      "YOUR NEXT WEEK INCENTIVE OPPORTUNITY",
      `${formatRsHighlight(nwOpp)} !!!`,
      y, pageState, "success"
    );

    y = highlightBox(
      doc,
      "TOTAL REVENUE",
      formatRsHighlight(nwRev),
      y, pageState, "info"
    );

    const safeName = (payload.mdo?.name || "Report").replace(/[^\w]+/g,"_");
    doc.save(`Performance_Report_${safeName}.pdf`);
  };
})();
