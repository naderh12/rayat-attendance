let rayatData=[];
document.getElementById("excelFile").addEventListener("change",loadExcel);
function loadExcel(e){
const file=e.target.files[0]; if(!file)return;
const reader=new FileReader();
reader.onload=function(ev){
const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
const ws=wb.Sheets[wb.SheetNames[0]];
rayatData=XLSX.utils.sheet_to_json(ws);
showInfo();
};
reader.readAsArrayBuffer(file);
}
function showInfo(){
if(!rayatData.length)return;
const f=rayatData[0];
document.getElementById('fileInfo').innerHTML=`<h3>تم تحميل الملف بنجاح ✅</h3><p>عدد الطلاب: ${rayatData.length}</p><p>CRN: ${f['CRN']||'-'}</p><p>Meeting ID: ${f['Meeting ID']||'-'}</p>`;
let html='<h3>أول 10 طلاب</h3><ol>';
rayatData.slice(0,10).forEach(s=>{html+=`<li>${s['Student ID']||''} - ${s['Student Name']||''}</li>`});
html+='</ol>';
document.getElementById('studentList').innerHTML=html;
}
