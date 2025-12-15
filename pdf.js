/* pdf.js */
(function(){
  const { jsPDF } = window.jspdf;

  function fmtINR(n){
    const x = Number(n || 0);
    return x.toLocaleString("en-IN");
  }
  function num(v){
    const x = Number(v);
    return Number.isFinite(x) ? x : 0;
  }

  window.generateA4Pdf = function(S){
    const doc = new jsPDF("p","mm","a4");
    let y = 14;

    doc.setFontSize(16);
    doc.text("Performance Report", 14, y);
    y += 8;

    doc.setFontSize(11);
    doc.text("Weekly Field Performance", 14, y);
    y += 10;

    // 1) MDO
    doc.setFontSize(13);
    doc.text("1) MDO Information", 14, y); y += 6;

    doc.setFontSize(10);
    const mdoRows = [
      ["Name", document.getElementById("mdoName").value || ""],
      ["HQ", document.getElementById("hq").value || ""],
      ["Region", document.getElementById("region").value || ""],
      ["Territory", document.getElementById("territory").value || ""],
      ["Month", document.getElementById("month").value || ""],
      ["Week", document.getElementById("week").value || ""],
    ];
    doc.autoTable({
      startY: y,
      head: [["Field", "Value"]],
      body: mdoRows,
      theme: "grid",
      styles: { fontSize: 9 },
      headStyles: { fillColor: [13,116,200] },
      margin: { left: 14, right: 14 }
    });
    y = doc.lastAutoTable.finalY + 8;

    // Helper to add tables safely
    function addSectionTable(title, head, body, afterGap=8){
      doc.setFontSize(13);
      doc.text(title, 14, y);
      y += 4;
      doc.autoTable({
        startY: y,
        head: [head],
        body,
        theme: "grid",
        styles: { fontSize: 9 },
        headStyles: { fillColor: [13,116,200] },
        margin: { left: 14, right: 14 },
      });
      y = doc.lastAutoTable.finalY + afterGap;
      if (y > 260){ doc.addPage(); y = 14; }
    }

    // 2) NPI
    let npiEarned = 0;
    const npiBody = (S.npiRows||[]).map((r, i)=>{
      const rate = (S.npiProducts||[]).find(p=>p.name===r.product)?.rate || 0;
      const opp = num(r.plan) * num(rate);
      const earned = num(r.actual) * num(rate);
      npiEarned += earned;
      return [String(i+1), r.product||"", String(r.plan||""), String(r.actual||""), fmtINR(opp), fmtINR(earned)];
    });
    addSectionTable(
      "2) NPI Performance Update",
      ["#", "Product", "Plan", "Actual", "Opportunity", "Earned"],
      npiBody.length ? npiBody : [["", "", "", "", "", ""]]
    );

    doc.setFontSize(11);
    doc.text(`Congratulations you have earned ${fmtINR(npiEarned)} Rs !!!`, 14, y);
    y += 10;

    // 3) Other
    let otherRevenue = 0;
    const otherBody = (S.otherRows||[]).map((r,i)=>{
      const rate = (S.otherProducts||[]).find(p=>p.name===r.product)?.revenuePerUnit || 0;
      const rev = num(r.actual) * num(rate);
      otherRevenue += rev;
      return [String(i+1), r.product||"", String(r.plan||""), String(r.actual||""), fmtINR(rev)];
    });
    addSectionTable(
      "3) Other Product Performance Update",
      ["#", "Product", "Plan", "Actual", "Revenue"],
      otherBody.length ? otherBody : [["", "", "", "", ""]]
    );
    doc.setFontSize(11);
    doc.text(`TOTAL REVENUE EARNED: ${fmtINR(otherRevenue)} Rs`, 14, y);
    y += 10;

    // 4) Activities
    const actBody = (S.actRows||[]).map((r,i)=>[
      String(i+1), r.activity||"", String(r.plan||""), String(r.actual||""), String(r.npiFocused||"")
    ]);
    addSectionTable(
      "4) Activities Update",
      ["#", "Activity", "Plan", "Actual", "NPI Focused"],
      actBody.length ? actBody : [["", "", "", "", ""]]
    );

    // 5) Next week product plan
    let nwIncentive = 0;
    const nwBody = (S.nwRows||[]).map((r,i)=>{
      const npiRate = (S.npiProducts||[]).find(p=>p.name===r.product)?.rate || 0;
      const otherRate = (S.otherProducts||[]).find(p=>p.name===r.product)?.revenuePerUnit || 0;
      const revenue = num(r.plan)*num(otherRate);
      const incentive = num(r.plan)*num(npiRate);
      nwIncentive += incentive;
      return [String(i+1), r.product||"", String(r.plan||""), fmtINR(revenue), fmtINR(incentive)];
    });
    addSectionTable(
      "5) Next Week Plan - Product Plan",
      ["#", "Product", "Plan", "Revenue", "Incentive Opportunity"],
      nwBody.length ? nwBody : [["", "", "", "", ""]]
    );
    doc.setFontSize(11);
    doc.text(`Your next week incentive opportunity is ${fmtINR(nwIncentive)} Rs !!!`, 14, y);
    y += 10;

    // 6) Next week activity plan
    const nwaBody = (S.nwaRows||[]).map((r,i)=>[
      String(i+1), r.activity||"", String(r.count||""), r.remarks||""
    ]);
    addSectionTable(
      "6) Next Week Activity Plan",
      ["#", "Activity", "Count", "Remarks"],
      nwaBody.length ? nwaBody : [["", "", "", ""]]
    );

    // 7) Special achievement
    const special = (document.getElementById("specialText").value || "").trim();
    doc.setFontSize(13);
    doc.text("7) Special Achievement", 14, y); y += 6;
    doc.setFontSize(10);
    const lines = doc.splitTextToSize(special || "-", 180);
    doc.text(lines, 14, y);
    y += (lines.length * 5) + 8;
    if (y > 260){ doc.addPage(); y = 14; }

    // 8) Photos (LAST) - no overlap, auto page breaks
    if ((S.photos||[]).some(p=>p.dataUrl)){
      doc.setFontSize(13);
      doc.text("8) Activities Photos", 14, y);
      y += 6;

      for (const p of (S.photos||[])){
        if (!p.dataUrl) continue;
        if (y > 240){ doc.addPage(); y = 14; }
        doc.setFontSize(10);
        doc.text(p.activity || "Photo", 14, y);
        y += 4;
        doc.addImage(p.dataUrl, "JPEG", 14, y, 60, 45);
        y += 55;
      }
    }

    doc.save("Performance_Report.pdf");
  };
})();
