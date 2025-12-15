/* Complete pdf.js – FULL VERSION (ALL SECTIONS)
   ✔ Single flowing PDF (no forced page break per section)
   ✔ No photo overlap (odd/even safe)
   ✔ Rs formatting safe: autoTable uses "Rs", highlight uses "₹"
   ✔ Next Week Plan = PLAN ONLY (no Actual)
   ✔ Realised fallback: Product List -> NPI Meta
   ✔ Comma-safe numeric parsing
   ✔ Exact key messages preserved:
     - CONGRATULATIONS YOU HAVE EARNED ₹ ___ !!!
     - YOU LOSE ₹ ___
     - YOUR NEXT WEEK INCENTIVE OPPORTUNITY ₹ ___ !!!
     - TOTAL REVENUE ₹ ___
*/

function formatRsTable(value){
  return "Rs " + Number(value || 0).toLocaleString("en-IN");
}
function formatRsHighlight(value){
  return "₹ " + Number(value || 0).toLocaleString("en-IN");
}
function toNumSafe(v){
  return Number(String(v ?? 0).replace(/,/g,"").trim()) || 0;
}

(function(){
  const { jsPDF } = window.jspdf;

  function ensureSpace(doc, y, needed, pageState){
    if (y + needed <= 280) return y;
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
    doc.text(`Page ${pageState.page}`, 195, 14, { align:"right" });

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
    doc.text(right, 192, y+9, { align:"right" });

    doc.setFont("helvetica","normal");
    return y + 18;
  }

  function autoTable(doc, head, body, startY){
    doc.autoTable({
      head: [head],
      body,
      startY,
      margin: { left: 15, right: 15 },
      styles: { fontSize: 9, cellPadding: 3, valign:"middle", font: "helvetica" },
      headStyles: { fillColor: [13,116,200], textColor: 255, fontStyle: "bold" },
      alternateRowStyles: { fillColor: [250,250,250] }
    });
    return doc.lastAutoTable.finalY + 8;
  }

  function addPhotos(doc, title, photos, y, pageState){
    y = sectionTitle(doc, title, y, pageState);

    if(!photos || !photos.length){
      y = ensureSpace(doc, y, 10, pageState);
      doc.setFontSize(9);
      doc.text("No photos uploaded.", 15, y);
      return y + 12;
    }

    const imgW = 80, imgH = 55, gapX = 10, gapY = 22;
    let x = 15, col = 0;

    for(const p of photos){
      y = ensureSpace(doc, y, imgH + gapY, pageState);

      try { doc.addImage(p.dataUrl, "JPEG", x, y, imgW, imgH); } catch {}

      doc.setFont("helvetica","bold");
      doc.setFontSize(9);
      doc.text(String(p.activity || "Activity"), x, y + imgH + 6);
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

    // ✅ critical fix for odd image counts (e.g., 13th photo)
    if(col === 1){
      y += imgH + gapY;
    }

    return y + 6;
  }

  window.generatePerformancePdf = function(payload, Master){
    const doc = new jsPDF({ unit:"mm", format:"a4" });
    const pageState = { page: 1 };

    header(doc, pageState);
    let y = 28;

    /* 1) MDO Information */
    y = sectionTitle(doc, "1) MDO Information", y, pageState);
    doc.setFontSize(10);

    const info = [
      ["Name", payload?.mdo?.name],
      ["HQ", payload?.mdo?.hq],
      ["Region", payload?.mdo?.region],
      ["Territory", payload?.mdo?.territory],
      ["Month", payload?.mdo?.month],
      ["Week", payload?.mdo?.week]
    ];

    for(const [k,v] of info){
      y = ensureSpace(doc, y, 8, pageState);
      doc.setFont("helvetica","bold");
      doc.text(`${k}:`, 15, y);
      doc.setFont("helvetica","normal");
      doc.text(String(v || ""), 55, y);
      y += 7;
    }
    y += 3;

    /* 2) NPI Performance Update */
    y = sectionTitle(doc, "2) NPI Performance Update", y, pageState);

    let npiOppTotal = 0, npiEarnTotal = 0;
    const npiBody = (payload.npiRows || []).map((r,i)=>{
      const rate = Master.npiMeta.get(r.product)?.incentiveRate || 0;
      const plan = toNumSafe(r.plan);
      const actual = toNumSafe(r.actual);

      const opp = plan * rate;
      const earn = actual * rate;

      npiOppTotal += opp;
      npiEarnTotal += earn;

      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
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
      `${formatRsHighlight(npiEarnTotal)} !!!`,
      y, pageState, "success"
    );

    y = highlightBox(
      doc,
      "YOU LOSE",
      formatRsHighlight(Math.max(0, npiOppTotal - npiEarnTotal)),
      y, pageState, "danger"
    );

    /* 3) Other Product Performance Update */
    y = sectionTitle(doc, "3) Other Product Performance Update", y, pageState);

    let otherTotal = 0;
    const otherBody = (payload.otherRows || []).map((r,i)=>{
      const realised = Master.productMeta.get(r.product)?.realised || 0;
      const actual = toNumSafe(r.actual);
      const rev = actual * realised;
      otherTotal += rev;

      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        formatRsTable(rev)
      ];
    });

    y = autoTable(doc,
      ["#","Product","Plan (L/Kg)","Actual (L/Kg)","Total Revenue"],
      otherBody, y
    );

    y = highlightBox(doc, "TOTAL REVENUE EARNED", formatRsHighlight(otherTotal), y, pageState, "info");

    /* 4) Activities Update */
    y = sectionTitle(doc, "4) Activities Update", y, pageState);

    let ap=0, aa=0, an=0;
    const actBody = (payload.activityRows || []).map((r,i)=>{
      ap += toNumSafe(r.planNo);
      aa += toNumSafe(r.actualNo);
      an += toNumSafe(r.npiNo);
      return [String(i+1), r.activity || "", String(r.planNo||""), String(r.actualNo||""), String(r.npiNo||"")];
    });

    y = autoTable(doc,
      ["#","Activity","Plan No","Actual No","NPI Focused Activity No"],
      actBody, y
    );

    y = ensureSpace(doc, y, 10, pageState);
    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text(`TOTAL  Plan: ${ap}   |   Actual: ${aa}   |   NPI Focused: ${an}`, 15, y);
    doc.setFont("helvetica","normal");
    y += 10;

    /* 5) Activities Photos */
    y = addPhotos(doc, "5) Activities Photos", (payload.photos || []), y, pageState);

    /* 6) Next Week Plan – Product Plan (PLAN ONLY) */
    y = sectionTitle(doc, "6) Next Week Plan – Product Plan", y, pageState);

    let nwRev = 0, nwOpp = 0;
    const nwBody = (payload.nextWeekRows || []).map((r,i)=>{
      const plan = toNumSafe(r.plan);

      const realised =
        Master.productMeta.get(r.product)?.realised ??
        Master.npiMeta.get(r.product)?.realised ??
        0;

      const rate = Master.npiMeta.get(r.product)?.incentiveRate ?? 0;

      const rev = plan * realised;
      const opp = plan * rate;

      nwRev += rev;
      nwOpp += opp;

      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
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

    y = highlightBox(doc, "TOTAL REVENUE", formatRsHighlight(nwRev), y, pageState, "info");

    /* 7) Activities Plan */
    y = sectionTitle(doc, "7) Activities Plan", y, pageState);

    const planBody = (payload.actPlanRows || []).map((r,i)=>{
      const villages = String(r.villages || "");
      const vCount = villages.split(",").map(s=>s.trim()).filter(Boolean).length;
      return [String(i+1), r.activity || "", String(r.planNo || ""), villages, String(vCount)];
    });

    y = autoTable(doc,
      ["#","Activity","Plan No","Village Names","Village No"],
      planBody, y
    );

    /* 8) Special Achievement */
    y = sectionTitle(doc, "8) Special Achievement", y, pageState);

    y = ensureSpace(doc, y, 10, pageState);
    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text("Description:", 15, y);
    y += 6;

    doc.setFont("helvetica","normal");
    doc.setFontSize(9);
    const desc = (payload.spDesc || "—").trim();
    const wrapped = doc.splitTextToSize(desc, 180);
    y = ensureSpace(doc, y, wrapped.length*5 + 10, pageState);
    doc.text(wrapped, 15, y);
    y += wrapped.length*5 + 8;

    y = addPhotos(doc, "Special Achievement Photos", (payload.spPhotos || []), y, pageState);

    const safeName = (payload.mdo?.name || "Report").replace(/[^\w]+/g,"_");
    doc.save(`Performance_Report_${safeName}.pdf`);
  };
})();
