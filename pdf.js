document.getElementById("btnPdf").onclick=()=>{
  const { jsPDF } = window.jspdf;
  const doc=new jsPDF("p","mm","a4");
  doc.text("MDO Weekly Report",15,15);
  doc.text("NPI Summary",15,30);
  doc.text("Total Incentive: "+npiInc.innerText,15,40);
  doc.text("Total Revenue: "+npiRev.innerText,15,48);
  doc.save("MDO_Weekly_Report.pdf");
};
