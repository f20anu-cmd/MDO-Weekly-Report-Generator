(function(){
const { jsPDF } = window.jspdf;

function header(doc,p){
doc.setFontSize(12);
doc.text("Performance Report",15,14);
doc.setFontSize(9);
doc.text(`Page ${p}`,195,14,{align:"right"});
doc.line(15,18,195,18);
}

function ensure(doc,y,p){
if(y>260){
doc.addPage();
header(doc,++p);
y=28;
}
return [y,p];
}

window.generateA4Pdf = ({State,rs,typeLabel})=>{
const doc=new jsPDF();
let p=1,y=28;
header(doc,p);

function section(t){
[y,p]=ensure(doc,y,p);
doc.setFontSize(11);
doc.setFont(undefined,"bold");
doc.text(t,15,y);
doc.setFont(undefined,"normal");
y+=6;
}

section("1) MDO Information");
["mdoName","hq","region","territory","month","week"].forEach(k=>{
doc.text(`${k}: ${State[k]||""}`,15,y);y+=6;
});

section("2) NPI Performance");
doc.autoTable({startY:y,head:[["#","Product","Plan","Actual","Incentive"]],
body:State.npiRows.map((r,i)=>[i+1,r.product,r.plan,r.actual,rs(r.incentiveEarned)])});
y=doc.lastAutoTable.finalY+6;

section("3) Other Products");
doc.autoTable({startY:y,head:[["#","Product","Plan","Actual","Revenue"]],
body:State.otherRows.map((r,i)=>[i+1,r.product,r.plan,r.actual,rs(r.revenue)])});
y=doc.lastAutoTable.finalY+6;

section("4) Activities");
doc.autoTable({startY:y,head:[["#","Activity","Plan","Actual","NPI"]],
body:State.actRows.map((r,i)=>[i+1,typeLabel(r.typeObj),r.planNo,r.actualNo,r.npiNo])});
y=doc.lastAutoTable.finalY+6;

section("6) Next Week Plan");
doc.autoTable({startY:y,head:[["#","Product","Plan","Actual","Revenue","Incentive"]],
body:State.nwRows.map((r,i)=>[i+1,r.product,r.plan,r.actual,rs(r.revenue),rs(r.incentive)])});
y=doc.lastAutoTable.finalY+6;

section("8) Special Achievement");
doc.text(doc.splitTextToSize(State.spDesc||"-",180),15,y);

doc.save("Performance_Report.pdf");
};
})();
