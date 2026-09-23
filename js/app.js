let rayatData = [];

document.getElementById("excelFile")
.addEventListener("change", e => {

    // عند اختيار ملف Rayat جديد:
    // مسح بيانات الجلسة السابقة من المتصفح
    localStorage.removeItem('sessionId');
    localStorage.removeItem('attendanceFinished');
    localStorage.removeItem('finalAttendance');

    // تصفير العداد والقائمة القديمة على الصفحة
    document.getElementById('attendanceCount').innerHTML =
        'الحضور الحالي: 0';

    document.getElementById('attendanceList').innerHTML =
        'لا يوجد حضور حتى الآن';

    // مسح معلومات الجلسة القديمة ورمز QR القديم
    if (typeof sessionInfo !== 'undefined') {
        sessionInfo.innerHTML = '';
    }

    if (typeof qrcode !== 'undefined') {
        qrcode.innerHTML = '';
    }

    // قراءة الملف الجديد
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


async function uploadStudentsToSupabase() {

    const currentSession = localStorage.getItem('sessionId');

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

            const errorText = await response.text();

            throw new Error(
                `فشل رفع الطالب: ${errorText}`
            );
        }
    }
}


document.getElementById('openBtn').onclick = async () => {

    if (!rayatData.length) {

        alert('ارفع الملف أولاً');

        return;
    }

    try {

        const sid = 'RAYAT-' + Date.now();

        localStorage.setItem(
            'sessionId',
            sid
        );

        localStorage.removeItem(
            'attendanceFinished'
        );

        sessionInfo.innerHTML =
            '<h3>Session: ' + sid + '</h3>';

        qrcode.innerHTML = '';

        new QRCode(
            document.getElementById('qrcode'),

            location.origin +
            location.pathname.replace('index.html', '') +
            'attendance.html?session=' + sid
        );

        await uploadStudentsToSupabase();

    } catch (error) {

        console.error(error);

        alert(
            'حدث خطأ أثناء فتح الحضور: ' +
            error.message
        );
    }
};


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

        // حفظ ملف الحضور النهائي
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

        // إغلاق تسجيل الطلاب لهذه الجلسة
        const deleteResponse = await fetch(

            `${SUPABASE_URL}/rest/v1/allowed_students?session_id=eq.${encodeURIComponent(currentSession)}`,

            {
                method: 'DELETE',

                headers: {
                    apikey: SUPABASE_KEY,
                    Authorization: `Bearer ${SUPABASE_KEY}`
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
