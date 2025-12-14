(function(){
  const { jsPDF } = window.jspdf;

  window.generatePDF = function(){
    const doc=new jsPDF("p","mm","a4");
    let y=15;
    const t=s=>{doc.text(s,15,y);y+=6;};

    t("Performance Report");
    t(`Name: ${mdoName.value}`);
    t(`HQ: ${hq.value}`);
    t(`Region: ${region.value}`);
    t(`Territory: ${territory.value}`);
    t(`Month: ${month.value}   Week: ${week.value}`);

    y+=4;
    doc.autoTable({startY:y,html:"#npiTable"});
    y=doc.lastAutoTable.finalY+4;
    t(`Total Incentive: ${npiTotal.innerText} Rs`);

    y+=4;
    doc.autoTable({startY:y,html:"#otherTable"});
    y=doc.lastAutoTable.finalY+4;
    t(`Total Revenue: ${otherTotal.innerText} Rs`);

    y+=4;
    doc.autoTable({startY:y,html:"#activityTable"});

    y+=4;
    doc.autoTable({startY:y,html:"#nextTable"});
    y=doc.lastAutoTable.finalY+4;
    t(`Next Week Revenue: ${nwRevenue.innerText} Rs`);
    t(`Next Week Incentive: ${nwIncentive.innerText} Rs`);

    y+=4;
    t("Special Achievement:");
    doc.text(specialText.value||"-",15,y,{maxWidth:180});

    doc.save("Performance_Report.pdf");
  };
})();
