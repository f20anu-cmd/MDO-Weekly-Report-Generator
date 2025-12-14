(function(){
  const { jsPDF } = window.jspdf;

  window.generateA4Pdf = function(d){
    const doc = new jsPDF("p","mm","a4");
    let y = 15;

    function line(txt){
      doc.text(txt, 15, y); y+=7;
    }

    doc.setFont("times","bold");
    line("Performance Report");
    doc.setFont("times","normal");
    line(`Name: ${d.mdoName}`);
    line(`HQ: ${d.hq}`);
    line(`Region: ${d.region}`);
    line(`Territory: ${d.territory}`);
    line(`Month: ${d.month}   Week: ${d.week}`);

    y+=5;
    doc.setFont("times","bold");
    line("NPI PERFORMANCE UPDATE");
    doc.autoTable({
      startY:y,
      head:[["Product","Plan","Actual","Incentive"]],
      body:d.npiRows.map(r=>[r.p,r.plan,r.act,(r.act||0)*0]),
      styles:{font:"times"}
    });

    y=doc.lastAutoTable.finalY+6;
    doc.text(`Total Incentive Earned: ${document.getElementById("npiTotal").innerText} Rs`,15,y);

    y+=10;
    doc.setFont("times","bold");
    line("OTHER PRODUCT PERFORMANCE UPDATE");
    doc.autoTable({
      startY:y,
      head:[["Product","Plan","Actual","Revenue"]],
      body:d.otherRows.map(r=>[r.p,r.plan,r.act,(r.act||0)*0]),
      styles:{font:"times"}
    });

    doc.save("Performance_Report.pdf");
  };
})();
