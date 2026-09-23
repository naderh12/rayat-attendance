let rayatData = [];


/* =========================================================
   عند فتح صفحة تحضير Rayat
   نبدأ الصفحة بدون عرض جلسة قديمة
   ========================================================= */

localStorage.removeItem('sessionId');
localStorage.removeItem('attendanceFinished');
localStorage.removeItem('finalAttendance');

if (typeof attendanceCount !== 'undefined') {
    attendanceCount.innerHTML = 'الحضور الحالي: 0';
}

if (typeof attendanceList !== 'undefined') {
    attendanceList.innerHTML = 'لا يوجد حضور حتى الآن';
}

if (typeof sessionInfo !== 'undefined') {
    sessionInfo.innerHTML = '';
}

if (typeof qrcode !== 'undefined') {
    qrcode.innerHTML = '';
}


/* =========================================================
   اختيار ملف Rayat
   ========================================================= */

document.getElementById("excelFile")
.addEventListener("change", e => {

    // مسح أي بيانات جلسة سابقة
    localStorage.removeItem('sessionId');
    localStorage.removeItem('attendanceFinished');
    localStorage.removeItem('finalAttendance');

    // تصفير العداد
    document.getElementById('attendanceCount').innerHTML =
        'الحضور الحالي: 0';

    // مسح قائمة الطلاب القديمة
    document.getElementById('attendanceList').innerHTML =
        'لا يوجد حضور حتى الآن';

    // مسح معلومات الجلسة القديمة
    if (typeof sessionInfo !== 'undefined') {
        sessionInfo.innerHTML = '';
    }

    // مسح QR القديم
    if (typeof qrcode !== 'undefined') {
        qrcode.innerHTML = '';
    }

    const f = e.target.files[0];

    if (!f) {
        return;
    }

    const r = new FileReader();

    r.onload = x => {

        const wb = XLSX.read(
            new Uint8Array(x.target.result),
            { type: 'array' }
        );

        const ws = wb.Sheets[wb.SheetNames[0]];

        rayatData = XLSX.utils.sheet_to_json(ws);

        const a = rayatData[0] || {};

        fileInfo.innerHTML =
            `عدد الطلاب: ${rayatData.length}<br>` +
            `CRN: ${a['CRN'] || ''}<br>` +
            `Meeting ID: ${a['Meeting ID'] || ''}`;

        localStorage.setItem(
            'rayatData',
            JSON.stringify(rayatData)
        );

        localStorage.setItem(
            'currentCRN',
            a['CRN'] || ''
        );

        localStorage.setItem(
            'meetingId',
            a['Meeting ID'] || ''
        );
    };

    r.readAsArrayBuffer(f);
});


/* =========================================================
   رفع الطلاب إلى Supabase
   ========================================================= */

async function uploadStudentsToSupabase() {

    const currentSession =
        localStorage.getItem('sessionId');

    if (!currentSession) {
        throw new Error('لا توجد جلسة حالية');
    }

    for (const student of rayatData) {

        const response = await fetch(
            `${SUPABASE_URL}/rest/v1/allowed_students`,
            {
                method: 'POST',

                headers: {
                    apikey: SUPABASE_KEY,
                    Authorization: `Bearer ${SUPABASE_KEY}`,
                    'Content-Type': 'application/json',
                    Prefer: 'return=minimal'
                },

                body: JSON.stringify({
                    session_id: currentSession,
                    student_id: String(student['Student ID'])
                })
            }
        );

        if (!response.ok) {

            const errorText =
                await response.text();

            throw new Error(
                `فشل رفع الطالب: ${errorText}`
            );
        }
    }
}


/* =========================================================
   فتح الحضور
   ========================================================= */

document.getElementById('openBtn').onclick = async () => {

    if (!rayatData.length) {

        alert('قم باختيار ملف Rayat أولاً');

        return;
    }

    try {

        const sid =
            'RAYAT-' + Date.now();

        // إنشاء جلسة جديدة
        localStorage.setItem(
            'sessionId',
            sid
        );

        localStorage.removeItem(
            'attendanceFinished'
        );

        localStorage.removeItem(
            'finalAttendance'
        );

        // عرض رقم الجلسة
        sessionInfo.innerHTML =
            '<h3>Session: ' + sid + '</h3>';

        // مسح QR القديم
        qrcode.innerHTML = '';

        // إنشاء QR جديد
        new QRCode(
            document.getElementById('qrcode'),

            location.origin +
            location.pathname.replace('index.html', '') +
            'attendance.html?session=' +
            sid
        );

        // تصفير العداد
        attendanceCount.innerHTML =
            'الحضور الحالي: 0';

        attendanceList.innerHTML =
            'لا يوجد حضور حتى الآن';

        // رفع الطلاب للجلسة الجديدة
        await uploadStudentsToSupabase();

    } catch (error) {

        console.error(error);

        alert(
            'حدث خطأ أثناء فتح الحضور: ' +
            error.message
        );
    }
};


/* =========================================================
   جلب قائمة الحضور
   ========================================================= */

async function getAttendanceList() {

    const currentSession =
        localStorage.getItem('sessionId');

    if (!currentSession) {

        throw new Error(
            'لا توجد جلسة حالية'
        );
    }

    const response = await fetch(

        `${SUPABASE_URL}/rest/v1/attendance_temp?session_id=eq.${encodeURIComponent(currentSession)}`,

        {
            headers: {
                apikey: SUPABASE_KEY,
                Authorization: `Bearer ${SUPABASE_KEY}`
            }
        }
    );

    if (!response.ok) {

        const errorText =
            await response.text();

        throw new Error(
            `فشل جلب الحضور: ${errorText}`
        );
    }

    return await response.json();
}


/* =========================================================
   إنهاء الحضور
   ========================================================= */

async function finishAttendance() {

    try {

        const currentSession =
            localStorage.getItem('sessionId');

        if (!currentSession) {

            alert(
                'لا توجد جلسة حضور حالية'
            );

            return;
        }

        const attendees =
            await getAttendanceList();

        const attendanceIds =
            attendees.map(
                x => String(x.student_id)
            );

        const result =
            rayatData.map(student => {

                const studentId =
                    String(student['Student ID']);

                if (
                    attendanceIds.includes(studentId)
                ) {

                    student['Attendance Indicator'] =
                        'Present';

                } else {

                    student['Attendance Indicator'] =
                        'Absent';
                }

                return student;
            });


        /* حفظ ملف الحضور النهائي */

        localStorage.setItem(
            'finalAttendance',
            JSON.stringify(result)
        );


        const savedFinalAttendance =
            localStorage.getItem(
                'finalAttendance'
            );

        if (!savedFinalAttendance) {

            throw new Error(
                'تعذر حفظ ملف الحضور النهائي'
            );
        }


        /* إغلاق تسجيل الطلاب */

        const deleteResponse =
            await fetch(

                `${SUPABASE_URL}/rest/v1/allowed_students?session_id=eq.${encodeURIComponent(currentSession)}`,

                {
                    method: 'DELETE',

                    headers: {
                        apikey: SUPABASE_KEY,
                        Authorization:
                            `Bearer ${SUPABASE_KEY}`
                    }
                }
            );


        if (!deleteResponse.ok) {

            const errorText =
                await deleteResponse.text();

            throw new Error(
                `تم تجهيز الملف لكن تعذر إغلاق التسجيل: ${errorText}`
            );
        }


        localStorage.setItem(
            'attendanceFinished',
            'true'
        );


        alert(
            'تم إنهاء الحضور وتجهيز ملف الحضور بنجاح'
        );


    } catch (error) {

        console.error(error);

        alert(
            'حدث خطأ أثناء إنهاء الحضور: ' +
            error.message
        );
    }
}
