function exportAttendanceFile(){

const storedData =
localStorage.getItem(
'finalAttendance'
);

if(!storedData){

alert(
'لا يوجد ملف حضور جاهز للتصدير'
);

return;

}

const attendanceData =
JSON.parse(storedData);

attendanceData.forEach(student => {

const status =
student["Attendance Indicator"];

if(status === "Present"){

student["Actual Hours"] =
student["Expected Hours"];

student["Absent Hours"] =
"00:00";

}
else{

student["Actual Hours"] =
"00:00";

student["Absent Hours"] =
student["Expected Hours"];

}

});

const worksheet =
XLSX.utils.json_to_sheet(
attendanceData
);

const workbook =
XLSX.utils.book_new();

XLSX.utils.book_append_sheet(
workbook,
worksheet,
"Attendance"
);

const crn =
attendanceData[0]["CRN"] || "Attendance";

XLSX.writeFile(
workbook,
`Attendance_${crn}.xlsx`
);

}
