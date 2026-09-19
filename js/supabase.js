const SUPABASE_URL =
"https://paimizculszksstmcfsq.supabase.co";

const SUPABASE_KEY =
"ضع هنا مفتاح anon كاملاً";

async function addAttendance(
sessionId,
studentId,
studentName,
crn
){

const response = await fetch(
`${SUPABASE_URL}/rest/v1/attendance_temp`,
{
method:"POST",

headers:{
"apikey": SUPABASE_KEY,
"Authorization":
`Bearer ${SUPABASE_KEY}`,
"Content-Type":
"application/json",
"Prefer":
"return=minimal"
},

body:JSON.stringify({
session_id:sessionId,
student_id:studentId,
student_name:studentName,
crn:crn
})

}
);

return response.ok;

}
