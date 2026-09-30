let rayatData = [];

/* =========================================================
   دعم قوالب Rayat العربية والإنجليزية
   ========================================================= */

let rayatHeaderMap = {};
let rayatLanguage = 'en';

/*
    ترتيب الأعمدة في قالب Rayat:

    A  Term Code
    B  CRN
    C  Session Indicator
    D  Start Time
    E  Student ID
    F  Student Name
    G  Confidential Indicator
    H  Meeting Days
    I  Expected Hours
    J  Actual Hours
    K  Absent Hours
    L  Attendance Indicator
    M  Authorised Absence
    N  Comments
    O  Meeting ID

    النسخة العربية تستخدم نفس ترتيب الأعمدة،
    ولكن بعناوين عربية.
*/

const RAYAT_COLUMNS = [
    'Term Code',
    'CRN',
    'Session Indicator',
    'Start Time',
    'Student ID',
    'Student Name',
    'Confidential Indicator',
    'Meeting Days',
    'Expected Hours',
    'Actual Hours',
    'Absent Hours',
    'Attendance Indicator',
    'Authorised Absence',
    'Comments',
    'Meeting ID'
];


/* =========================================================
   تنظيف اسم رأس العمود للمقارنة
   ========================================================= */

function normalizeHeader(value) {

    return String(value || '')
        .replace(/\uFEFF/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}


/* =========================================================
   اكتشاف لغة القالب
   ========================================================= */

function detectRayatLanguage(headers) {

    const normalizedHeaders =
        headers.map(normalizeHeader);

    const isEnglish =
        normalizedHeaders.includes('student id') &&
        normalizedHeaders.includes('meeting id') &&
        normalizedHeaders.includes('crn');

    return isEnglish ? 'en' : 'ar';
}


/* =========================================================
   إنشاء خريطة بين الحقول الداخلية
   وأسماء الأعمدة الموجودة فعلياً في الملف
   ========================================================= */

function buildRayatHeaderMap(headers) {

    const map = {};

    const normalizedHeaders =
        headers.map(normalizeHeader);

    RAYAT_COLUMNS.forEach(
        (canonicalName, columnIndex) => {

            /*
                أولاً:
                إذا كان اسم العمود الإنجليزي موجوداً
                نستخدم موقعه الحقيقي.
            */

            let actualIndex =
                normalizedHeaders.indexOf(
                    normalizeHeader(canonicalName)
                );

            /*
                ثانياً:
                في القالب العربي تكون أسماء الأعمدة مختلفة،
                لكن ترتيب الأعمدة هو نفسه.

                لذلك نستخدم موقع العمود في قالب Rayat.
            */

            if (
                actualIndex === -1 &&
                headers[columnIndex] !== undefined
            ) {

                actualIndex = columnIndex;
            }

            if (actualIndex !== -1) {

                map[canonicalName] =
                    headers[actualIndex];
            }
        }
    );

    return map;
}


/* =========================================================
   قراءة حقل من الطالب
   سواء كان رأس العمود عربياً أو إنجليزياً
   ========================================================= */

function getRayatField(row, canonicalName) {

    if (!row) {
        return '';
    }

    const actualHeader =
        rayatHeaderMap[canonicalName];

    if (!actualHeader) {
        return '';
    }

    return row[actualHeader];
}


/* =========================================================
   كتابة قيمة في الحقل الأصلي
   مع الاحتفاظ بلغة رأس العمود
   ========================================================= */

function setRayatField(
    row,
    canonicalName,
    value
) {

    const actualHeader =
        rayatHeaderMap[canonicalName];

    if (!actualHeader) {

        throw new Error(
            'تعذر العثور على العمود: ' +
            canonicalName
        );
    }

    row[actualHeader] = value;
}


/* =========================================================
   التحقق من أن الملف هو قالب Rayat الصحيح
   ========================================================= */

function validateRayatTemplate(headers) {

    /*
        قالب Rayat المستخدم في المشروع
        يحتوي على الأعمدة من A إلى O.
    */

    if (!headers || headers.length < 15) {

        throw new Error(
            'الملف لا يحتوي على أعمدة قالب Rayat المطلوبة'
        );
    }

    const requiredFields = [
        'CRN',
        'Student ID',
        'Expected Hours',
        'Actual Hours',
        'Absent Hours',
        'Attendance Indicator',
        'Meeting ID'
    ];

    for (const field of requiredFields) {

        const actualHeader =
            rayatHeaderMap[field];

        if (
            actualHeader === undefined ||
            actualHeader === null ||
            String(actualHeader).trim() === ''
        ) {

            throw new Error(
                'لم يتم العثور على العمود المطلوب: ' +
                field
            );
        }
    }
}


/* =========================================================
   عند فتح صفحة تحضير Rayat
   نبدأ الصفحة بدون عرض جلسة قديمة
   ========================================================= */

localStorage.removeItem('sessionId');
localStorage.removeItem('attendanceFinished');
localStorage.removeItem('finalAttendance');
localStorage.removeItem('rayatHeaderMap');
localStorage.removeItem('rayatLanguage');

if (typeof attendanceCount !== 'undefined') {

    attendanceCount.innerHTML =
        'الحضور الحالي: 0';
}

if (typeof attendanceList !== 'undefined') {

    attendanceList.innerHTML =
        'لا يوجد حضور حتى الآن';
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

document
.getElementById("excelFile")
.addEventListener(
    "change",
    e => {

        /* ---------------------------------------------
           مسح أي بيانات جلسة سابقة
           --------------------------------------------- */

        localStorage.removeItem('sessionId');
        localStorage.removeItem('attendanceFinished');
        localStorage.removeItem('finalAttendance');
        localStorage.removeItem('rayatHeaderMap');
        localStorage.removeItem('rayatLanguage');

        rayatData = [];
        rayatHeaderMap = {};
        rayatLanguage = 'en';


        /* ---------------------------------------------
           تصفير العداد
           --------------------------------------------- */

        document
        .getElementById('attendanceCount')
        .innerHTML =
            'الحضور الحالي: 0';


        /* ---------------------------------------------
           مسح قائمة الطلاب القديمة
           --------------------------------------------- */

        document
        .getElementById('attendanceList')
        .innerHTML =
            'لا يوجد حضور حتى الآن';


        /* ---------------------------------------------
           مسح معلومات الجلسة القديمة
           --------------------------------------------- */

        if (
            typeof sessionInfo !==
            'undefined'
        ) {

            sessionInfo.innerHTML = '';
        }


        /* ---------------------------------------------
           مسح QR القديم
           --------------------------------------------- */

        if (
            typeof qrcode !==
            'undefined'
        ) {

            qrcode.innerHTML = '';
        }


        const f =
            e.target.files[0];

        if (!f) {

            return;
        }


        const r =
            new FileReader();


        r.onload =
            x => {

                try {

                    /* =====================================
                       قراءة ملف Excel
                       ===================================== */

                    const wb =
                        XLSX.read(
                            new Uint8Array(
                                x.target.result
                            ),
                            {
                                type: 'array'
                            }
                        );


                    const ws =
                        wb.Sheets[
                            wb.SheetNames[0]
                        ];


                    /* =====================================
                       قراءة صف العناوين أولاً
                       ===================================== */

                    const rawRows =
                        XLSX.utils.sheet_to_json(
                            ws,
                            {
                                header: 1,
                                defval: ''
                            }
                        );


                    if (
                        !rawRows.length ||
                        !rawRows[0]
                    ) {

                        throw new Error(
                            'الملف لا يحتوي على بيانات'
                        );
                    }


                    const headers =
                        rawRows[0].map(
                            header =>
                                String(
                                    header || ''
                                ).trim()
                        );


                    /* =====================================
                       تحديد لغة قالب Rayat
                       ===================================== */

                    rayatLanguage =
                        detectRayatLanguage(
                            headers
                        );


                    /* =====================================
                       إنشاء خريطة الأعمدة
                       ===================================== */

                    rayatHeaderMap =
                        buildRayatHeaderMap(
                            headers
                        );


                    /* =====================================
                       التحقق من القالب
                       ===================================== */

                    validateRayatTemplate(
                        headers
                    );


                    /* =====================================
                       قراءة بيانات الطلاب
                       مع الاحتفاظ بعناوين الملف الأصلية
                       ===================================== */

                    rayatData =
                        XLSX.utils.sheet_to_json(
                            ws,
                            {
                                defval: ''
                            }
                        );


                    if (!rayatData.length) {

                        throw new Error(
                            'لا يوجد طلاب في الملف'
                        );
                    }


                    const firstStudent =
                        rayatData[0] || {};


                    const currentCRN =
                        getRayatField(
                            firstStudent,
                            'CRN'
                        ) || '';


                    const currentMeetingId =
                        getRayatField(
                            firstStudent,
                            'Meeting ID'
                        ) || '';


                    /* =====================================
                       عرض معلومات الملف
                       ===================================== */

                    const languageName =
                        rayatLanguage === 'ar'
                            ? 'العربية'
                            : 'English';


                    fileInfo.innerHTML =
                        `عدد الطلاب: ${rayatData.length}<br>` +
                        `CRN: ${currentCRN}<br>` +
                        `Meeting ID: ${currentMeetingId}<br>` +
                        `لغة القالب: ${languageName}`;


                    /* =====================================
                       حفظ البيانات
                       ===================================== */

                    localStorage.setItem(
                        'rayatData',
                        JSON.stringify(
                            rayatData
                        )
                    );


                    localStorage.setItem(
                        'currentCRN',
                        String(
                            currentCRN
                        )
                    );


                    localStorage.setItem(
                        'meetingId',
                        String(
                            currentMeetingId
                        )
                    );


                    localStorage.setItem(
                        'rayatHeaderMap',
                        JSON.stringify(
                            rayatHeaderMap
                        )
                    );


                    localStorage.setItem(
                        'rayatLanguage',
                        rayatLanguage
                    );


                } catch (error) {

                    console.error(
                        'Rayat file error:',
                        error
                    );


                    rayatData = [];
                    rayatHeaderMap = {};


                    fileInfo.innerHTML =
                        'تعذر قراءة ملف Rayat';


                    alert(
                        'تعذر قراءة ملف Rayat: ' +
                        error.message
                    );
                }
            };


        r.onerror =
            () => {

                rayatData = [];
                rayatHeaderMap = {};

                alert(
                    'حدث خطأ أثناء قراءة الملف'
                );
            };


        r.readAsArrayBuffer(f);
    }
);


/* =========================================================
   رفع الطلاب إلى Supabase
   ========================================================= */

async function uploadStudentsToSupabase() {

    const currentSession =
        localStorage.getItem(
            'sessionId'
        );


    if (!currentSession) {

        throw new Error(
            'لا توجد جلسة حالية'
        );
    }


    for (
        const student
        of rayatData
    ) {

        const studentId =
            String(
                getRayatField(
                    student,
                    'Student ID'
                ) || ''
            ).trim();


        /*
            نتجاهل أي صف فارغ
            في نهاية ملف Excel.
        */

        if (!studentId) {

            continue;
        }


        const response =
            await fetch(

                `${SUPABASE_URL}/rest/v1/allowed_students`,

                {
                    method: 'POST',

                    headers: {

                        apikey:
                            SUPABASE_KEY,

                        Authorization:
                            `Bearer ${SUPABASE_KEY}`,

                        'Content-Type':
                            'application/json',

                        Prefer:
                            'return=minimal'
                    },

                    body:
                        JSON.stringify(
                            {
                                session_id:
                                    currentSession,

                                student_id:
                                    studentId
                            }
                        )
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

document
.getElementById('openBtn')
.onclick =
async () => {

    if (!rayatData.length) {

        alert(
            'قم باختيار ملف Rayat أولاً'
        );

        return;
    }


    try {

        const sid =
            'RAYAT-' +
            Date.now();


        /* ---------------------------------------------
           إنشاء جلسة جديدة
           --------------------------------------------- */

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


        /* ---------------------------------------------
           عرض رقم الجلسة
           --------------------------------------------- */

        sessionInfo.innerHTML =
            '<h3>Session: ' +
            sid +
            '</h3>';


        /* ---------------------------------------------
           مسح QR القديم
           --------------------------------------------- */

        qrcode.innerHTML = '';


        /* ---------------------------------------------
           إنشاء QR جديد
           --------------------------------------------- */

        new QRCode(

            document
            .getElementById(
                'qrcode'
            ),

            location.origin +

            location.pathname
            .replace(
                'index.html',
                ''
            ) +

            'attendance.html?session=' +

            sid
        );


        /* ---------------------------------------------
           تصفير العداد
           --------------------------------------------- */

        attendanceCount.innerHTML =
            'الحضور الحالي: 0';


        attendanceList.innerHTML =
            'لا يوجد حضور حتى الآن';


        /* ---------------------------------------------
           رفع الطلاب للجلسة الجديدة
           --------------------------------------------- */

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
        localStorage.getItem(
            'sessionId'
        );


    if (!currentSession) {

        throw new Error(
            'لا توجد جلسة حالية'
        );
    }


    const response =
        await fetch(

            `${SUPABASE_URL}/rest/v1/attendance_temp?session_id=eq.${encodeURIComponent(currentSession)}`,

            {
                headers: {

                    apikey:
                        SUPABASE_KEY,

                    Authorization:
                        `Bearer ${SUPABASE_KEY}`
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
            localStorage.getItem(
                'sessionId'
            );


        if (!currentSession) {

            alert(
                'لا توجد جلسة حضور حالية'
            );

            return;
        }


        /* =================================================
           جلب الحضور قبل الحذف
           ================================================= */

        const attendees =
            await getAttendanceList();


        const attendanceIds =
            attendees.map(
                x =>
                    String(
                        x.student_id
                    )
            );


        /* =================================================
           تجهيز الملف النهائي
           ================================================= */

        const result =
            rayatData.map(
                student => {

                    const studentId =
                        String(
                            getRayatField(
                                student,
                                'Student ID'
                            ) || ''
                        );


                    if (
                        attendanceIds.includes(
                            studentId
                        )
                    ) {

                        /*
                            نحتفظ بالقيمة الداخلية
                            Present كما في النظام الأصلي.

                            عند تعديل export.js
                            سنتعامل مع لغة ملف Rayat
                            عند إنشاء ملف التصدير النهائي.
                        */

                        setRayatField(
                            student,
                            'Attendance Indicator',
                            'Present'
                        );

                    } else {

                        setRayatField(
                            student,
                            'Attendance Indicator',
                            'Absent'
                        );
                    }


                    return student;
                }
            );


        /* =================================================
           حفظ ملف الحضور النهائي
           ================================================= */

        localStorage.setItem(
            'finalAttendance',
            JSON.stringify(
                result
            )
        );


        /*
            نحفظ خريطة الأعمدة أيضاً
            لكي يستخدمها export.js لاحقاً.
        */

        localStorage.setItem(
            'rayatHeaderMap',
            JSON.stringify(
                rayatHeaderMap
            )
        );


        localStorage.setItem(
            'rayatLanguage',
            rayatLanguage
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


        /* =================================================
           إغلاق تسجيل الطلاب
           حذف allowed_students
           ================================================= */

        const deleteAllowedResponse =
            await fetch(

                `${SUPABASE_URL}/rest/v1/allowed_students?session_id=eq.${encodeURIComponent(currentSession)}`,

                {
                    method:
                        'DELETE',

                    headers: {

                        apikey:
                            SUPABASE_KEY,

                        Authorization:
                            `Bearer ${SUPABASE_KEY}`
                    }
                }
            );


        if (
            !deleteAllowedResponse.ok
        ) {

            const errorText =
                await deleteAllowedResponse
                .text();


            throw new Error(
                `تم تجهيز الملف لكن تعذر إغلاق التسجيل: ${errorText}`
            );
        }


        /* =================================================
           تنظيف attendance_temp

           بعد تجهيز finalAttendance بنجاح
           نحذف سجلات الحضور المؤقتة للجلسة.
           ================================================= */

        const deleteAttendanceResponse =
            await fetch(

                `${SUPABASE_URL}/rest/v1/attendance_temp?session_id=eq.${encodeURIComponent(currentSession)}`,

                {
                    method:
                        'DELETE',

                    headers: {

                        apikey:
                            SUPABASE_KEY,

                        Authorization:
                            `Bearer ${SUPABASE_KEY}`
                    }
                }
            );


        if (
            !deleteAttendanceResponse.ok
        ) {

            const errorText =
                await deleteAttendanceResponse
                .text();


            throw new Error(
                `تم تجهيز الملف وإغلاق التسجيل، لكن تعذر تنظيف سجلات الحضور المؤقتة: ${errorText}`
            );
        }


        /* =================================================
           تسجيل أن الحضور انتهى
           ================================================= */

        localStorage.setItem(
            'attendanceFinished',
            'true'
        );


        alert(
            'تم إنهاء الحضور وتجهيز ملف الحضور وتنظيف السجلات المؤقتة بنجاح'
        );


    } catch (error) {

        console.error(error);


        alert(
            'حدث خطأ أثناء إنهاء الحضور: ' +
            error.message
        );
    }
}
