(function(){
const {jsPDF}=window.jspdf;
window.generateA4Pdf=s=>{
const d=new jsPDF();let y=15;
d.text("Performance Report",10,y);y+=10;
["npi","other","act","nw","nwa"].forEach(k=>{d.text(`${k}: ${s[k].length}`,10,y);y+=8});
if(s.photos.length){
  d.addPage();y=15;d.text("Activity Photos",10,y);y+=5;
  s.photos.forEach(p=>{
    if(p.img){
      if(y>250){d.addPage();y=15}
      d.addImage(p.img,"JPEG",10,y,40,30);
      y+=35;
    }
  })
}
d.save("Performance_Report.pdf");
}
})();
