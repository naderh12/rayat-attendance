let rayatData = [];

document.getElementById("excelFile")
.addEventListener("change", e => {
const f = e.target.files[0];
const r = new FileReader();

r.onload = x => {
const wb = XLSX.read(new Uint8Array(x.target.result), { type: 'array' });
const ws = wb.Sheets[wb.SheetNames[0]];

rayatData = XLSX.utils.sheet_to_json(ws);

const a = rayatData[0] || {};

fileInfo.innerHTML = `عدد الطلاب: ${rayatData.length}<br>CRN: ${a['CRN'] || ''}<br>Meeting ID: ${a['Meeting ID'] || ''}`;

localStorage.setItem('rayatData', JSON.stringify(rayatData));
localStorage.setItem('currentCRN', a['CRN'] || '');
localStorage.setItem('meetingId', a['Meeting ID'] || '');
};

r.readAsArrayBuffer(f);
});

async function uploadStudentsToSupabase(){
const currentSession = localStorage.getItem('sessionId');

for(const student of rayatData){
await fetch(`${SUPABASE_URL}/rest/v1/allowed_students`,{
method:'POST',
headers:{
'apikey':SUPABASE_KEY,
'Authorization':`Bearer ${SUPABASE_KEY}`,
'Content-Type':'application/json'
},
body:JSON.stringify({
session_id:currentSession,
student_id:String(student['Student ID'])
})
});
}
}

document.getElementById('openBtn').onclick = async () => {
if (!rayatData.length) {
alert('ارفع الملف أولا');
return;
}

let sid = 'RAYAT-' + Date.now();
localStorage.setItem('sessionId', sid);

sessionInfo.innerHTML = '<h3>Session: ' + sid + '</h3>';
qrcode.innerHTML='';

new QRCode(document.getElementById('qrcode'),
location.origin + location.pathname.replace('index.html','') + 'attendance.html?session=' + sid);

await uploadStudentsToSupabase();
};

async function getAttendanceList(){
const currentSession = localStorage.getItem('sessionId');
const response = await fetch(`${SUPABASE_URL}/rest/v1/attendance_temp?session_id=eq.${currentSession}`,{
headers:{apikey:SUPABASE_KEY,Authorization:`Bearer ${SUPABASE_KEY}`}
});
return await response.json();
}

async function finishAttendance(){
const attendees = await getAttendanceList();
const attendanceIds = attendees.map(x => String(x.student_id));

const result = rayatData.map(student => {
const studentId = String(student['Student ID']);
student['Attendance Indicator'] = attendanceIds.includes(studentId) ? 'Present' : 'Absent';
return student;
});

localStorage.setItem('finalAttendance', JSON.stringify(result));
alert('تم تجهيز ملف الحضور بنجاح');
console.log(result);
}
