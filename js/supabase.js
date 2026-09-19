const SUPABASE_URL =
"https://paimizculszksstmcfsq.supabase.co";

const SUPABASE_KEY =
"eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBhaW1pemN1bHN6a3NzdG1jZnNxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk4MDk4NTYsImV4cCI6MjEwNTM4NTg1Nn0.9MlGrXKL745jPWygSuOSWVh-eLX-sR565zHYDRuufK0";

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
