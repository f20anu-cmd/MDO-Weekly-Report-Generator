/* Complete pdf.js (single flowing PDF, no forced page break per section)
   Includes required exact messages:
   - CONGRATULATIONS YOU HAVE EARNED ₹ ___ !!!
   - YOU LOSE ₹ ___
   - YOUR NEXT WEEK INCENTIVE OPPORTUNITY ₹ ___ !!!
   - TOTAL REVENUE ₹ ___
   Includes photo grids without overlap.
*/

(function(){
  const { jsPDF } = window.jspdf;

  function money(rsFn, n){ return `₹ ${rsFn(n)}`; }

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

  function highlightBox(doc, textLeft, textRight, y, pageState, style){
    y = ensureSpace(doc, y, 18, pageState);
    if(style === "success") doc.setFillColor(233, 255, 245);
    else if(style === "danger") doc.setFillColor(255, 232, 234);
    else doc.setFillColor(224, 242, 254);

    doc.roundedRect(15, y, 180, 14, 3, 3, "F");
    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.setTextColor(20);

    doc.text(textLeft, 18, y+9);
    doc.text(textRight, 192, y+9, {align:"right"});

    doc.setTextColor(0);
    doc.setFont("helvetica","normal");
    return y + 18;
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

  function addPhotos(doc, title, photos, y, pageState){
    y = sectionTitle(doc, title, y, pageState);

    if(!photos.length){
      y = ensureSpace(doc, y, 12, pageState);
      doc.setFontSize(10);
      doc.text("No photos uploaded.", 15, y+6);
      return y + 14;
    }

    const x0 = 15;
    const imgW = 80;
    const imgH = 55;
    const gapX = 10;
    const gapY = 22;

    let x = x0;
    let col = 0;

    for(const p of photos){
      y = ensureSpace(doc, y, imgH + 22, pageState);

      try{
        doc.addImage(p.dataUrl, "JPEG", x, y, imgW, imgH);
      }catch(e){
        // ignore bad image
      }

      doc.setFont("helvetica","bold");
      doc.setFontSize(9);
      doc.setTextColor(90);
      doc.text(String(p.activity || "Activity"), x, y + imgH + 6);
      doc.setTextColor(0);
      doc.setFont("helvetica","normal");

      col++;
      if(col === 2){
        col = 0;
        x = x0;
        y += imgH + gapY;
      }else{
        x += imgW + gapX;
      }
    }

    // reserve space after the final row so next section never overlaps
    return y + 10;
  }

  window.generatePerformancePdf = function(payload, Master, rsFn){
    const doc = new jsPDF({unit:"mm", format:"a4", orientation:"portrait"});
    const pageState = { page: 1 };

    header(doc, pageState);
    let y = 28;

    // 1) MDO
    y = sectionTitle(doc, "1) MDO Information", y, pageState);
    doc.setFontSize(10);
    const info = [
      ["Name", payload.mdo.name],
      ["HQ", payload.mdo.hq],
      ["Region", payload.mdo.region],
      ["Territory", payload.mdo.territory],
      ["Month", payload.mdo.month],
      ["Week", payload.mdo.week]
    ];
    for(const [k,v] of info){
      y = ensureSpace(doc, y, 8, pageState);
      doc.setFont("helvetica","bold");
      doc.text(`${k}:`, 15, y);
      doc.setFont("helvetica","normal");
      doc.text(String(v || ""), 55, y);
      y += 7;
    }
    y += 4;

    // 2) NPI
    y = sectionTitle(doc, "2) NPI Performance Update", y, pageState);

    let npiOppTotal = 0, npiEarnTotal = 0;
    const npiBody = (payload.npiRows || []).map((r, i)=>{
      const meta = Master.npiMeta.get(r.product) || { incentiveRate: 0 };
      const plan = Number(String(r.plan||0).replace(/,/g,"")) || 0;
      const actual = Number(String(r.actual||0).replace(/,/g,"")) || 0;
      const opp = plan * meta.incentiveRate;
      const earn = actual * meta.incentiveRate;
      npiOppTotal += opp;
      npiEarnTotal += earn;
      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        money(rsFn, opp),
        money(rsFn, earn)
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Actual (L/Kg)", "Total Incentive Opportunity", "Total Incentive Earned"],
      npiBody,
      y
    );

    const npiLose = Math.max(0, npiOppTotal - npiEarnTotal);

    // EXACT wording requested (both in UI and PDF)
    y = highlightBox(
      doc,
      "CONGRATULATIONS YOU HAVE EARNED",
      `${money(rsFn, npiEarnTotal)} !!!`,
      y, pageState, "success"
    );
    y = highlightBox(
      doc,
      "YOU LOSE",
      `${money(rsFn, npiLose)}`,
      y, pageState, "danger"
    );

    // 3) Other products
    y = sectionTitle(doc, "3) Other Product Performance Update", y, pageState);

    let otherTotal = 0;
    const otherBody = (payload.otherRows || []).map((r,i)=>{
      const meta = Master.productMeta.get(r.product) || { realised: 0 };
      const actual = Number(String(r.actual||0).replace(/,/g,"")) || 0;
      const rev = actual * meta.realised;
      otherTotal += rev;
      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        money(rsFn, rev)
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Actual (L/Kg)", "Total Revenue"],
      otherBody,
      y
    );

    y = highlightBox(doc, "TOTAL REVENUE EARNED", money(rsFn, otherTotal), y, pageState, "info");

    // 4) Activities update
    y = sectionTitle(doc, "4) Activities Update", y, pageState);

    let ap=0, aa=0, an=0;
    const actBody = (payload.activityRows || []).map((r,i)=>{
      ap += Number(r.planNo||0) || 0;
      aa += Number(r.actualNo||0) || 0;
      an += Number(r.npiNo||0) || 0;
      return [
        String(i+1),
        r.activity || "",
        String(r.planNo || ""),
        String(r.actualNo || ""),
        String(r.npiNo || "")
      ];
    });

    y = autoTable(doc,
      ["#", "Activity", "Plan No", "Actual No", "NPI Focused Activity No"],
      actBody,
      y
    );

    y = ensureSpace(doc, y, 10, pageState);
    doc.setFont("helvetica","bold");
    doc.setFontSize(10);
    doc.text(`TOTAL  Plan: ${ap}   |   Actual: ${aa}   |   NPI Focused: ${an}`, 15, y);
    doc.setFont("helvetica","normal");
    y += 10;

    // 5) Photos
    y = addPhotos(doc, "5) Activities Photos", payload.photos || [], y, pageState);

    // 6) Next week
    y = sectionTitle(doc, "6) Next Week Plan – Product Plan", y, pageState);

    let nwRev=0, nwOpp=0;
    const nwBody = (payload.nextWeekRows || []).map((r,i)=>{
      const plan = Number(String(r.plan||0).replace(/,/g,"")) || 0;

      const realised =
        Master.productMeta.get(r.product)?.realised ??
        Master.npiMeta.get(r.product)?.realised ??
        0;

      const rate = Master.npiMeta.get(r.product)?.incentiveRate ?? 0;

      const rev = plan * realised;
      const opp2 = plan * rate;

      nwRev += rev;
      nwOpp += opp2;

      return [
        String(i+1),
        r.product || "",
        String(r.plan || ""),
        String(r.actual || ""),
        money(rsFn, rev),
        money(rsFn, opp2)
      ];
    });

    y = autoTable(doc,
      ["#", "Product", "Plan (L/Kg)", "Actual (L/Kg)", "Total Revenue", "Total Incentive Earned"],
      nwBody,
      y
    );

    // EXACT wording requested (both in UI and PDF)
    y = highlightBox(
      doc,
      "YOUR NEXT WEEK INCENTIVE OPPORTUNITY",
      `${money(rsFn, nwOpp)} !!!`,
      y, pageState, "success"
    );
    y = highlightBox(
      doc,
      "TOTAL REVENUE",
      money(rsFn, nwRev),
      y, pageState, "info"
    );

    // 7) Activities plan
    y = sectionTitle(doc, "7) Activities Plan", y, pageState);

    const planBody = (payload.actPlanRows || []).map((r,i)=>{
      const villages = String(r.villages || "");
      const vCount = villages.split(",").map(s=>s.trim()).filter(Boolean).length;
      return [
        String(i+1),
        r.activity || "",
        String(r.planNo || ""),
        villages,
        String(vCount)
      ];
    });

    y = autoTable(doc,
      ["#", "Activity", "Plan No", "Village Names", "Village No"],
      planBody,
      y
    );

    // 8) Special achievement
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

    y = addPhotos(doc, "Special Achievement Photos", payload.spPhotos || [], y, pageState);

    const safeName = (payload.mdo?.name || "Report").replace(/[^\w]+/g, "_");
    doc.save(`Performance_Report_${safeName}.pdf`);
  };
})();
