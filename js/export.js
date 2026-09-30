function exportAttendanceFile() {

    try {

        /* =====================================================
           قراءة ملف الحضور النهائي
           ===================================================== */

        const storedData =
            localStorage.getItem('finalAttendance');

        if (!storedData) {

            alert(
                'لا يوجد ملف حضور جاهز للتصدير.\n' +
                'يرجى الضغط على "إنهاء الحضور" أولاً.'
            );

            return;
        }


        const attendanceData =
            JSON.parse(storedData);


        if (
            !Array.isArray(attendanceData) ||
            attendanceData.length === 0
        ) {

            alert(
                'ملف الحضور النهائي فارغ أو غير صالح.'
            );

            return;
        }


        /* =====================================================
           قراءة لغة القالب وخريطة الأعمدة
           التي تم حفظها بواسطة app.js
           ===================================================== */

        const rayatLanguage =
            localStorage.getItem('rayatLanguage') || 'en';


        let rayatHeaderMap = {};

        try {

            rayatHeaderMap =
                JSON.parse(
                    localStorage.getItem('rayatHeaderMap') || '{}'
                );

        } catch (error) {

            rayatHeaderMap = {};
        }


        /* =====================================================
           أسماء الحقول الداخلية الثابتة
           ===================================================== */

        const canonicalHeaders = [

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


        /* =====================================================
           الحصول على اسم العمود الحقيقي في الملف

           إنجليزي:
           Student ID

           عربي:
           الاسم العربي الأصلي الموجود في قالب Rayat
           ===================================================== */

        function getActualHeader(canonicalName) {

            return (
                rayatHeaderMap[canonicalName] ||
                canonicalName
            );
        }


        /* =====================================================
           قراءة قيمة من السجل
           ===================================================== */

        function getField(student, canonicalName) {

            const actualHeader =
                getActualHeader(canonicalName);

            if (
                Object.prototype.hasOwnProperty.call(
                    student,
                    actualHeader
                )
            ) {

                return student[actualHeader];
            }


            /*
                احتياطياً للقوالب الإنجليزية القديمة
            */

            if (
                Object.prototype.hasOwnProperty.call(
                    student,
                    canonicalName
                )
            ) {

                return student[canonicalName];
            }


            return '';
        }


        /* =====================================================
           إنشاء رؤوس الأعمدة بنفس لغة الملف الأصلي
           ===================================================== */

        const headers =
            canonicalHeaders.map(
                canonicalName =>
                    getActualHeader(canonicalName)
            );


        /* =====================================================
           إنشاء الصفوف
           ===================================================== */

        const rows =
            attendanceData.map(student => {

                /* ---------------------------------------------
                   حالة الحضور الداخلية
                   --------------------------------------------- */

                const internalStatus =
                    String(
                        getField(
                            student,
                            'Attendance Indicator'
                        ) || ''
                    ).trim();


                const isPresent =
                    internalStatus === 'Present' ||
                    internalStatus === 'حاضر';


                /* ---------------------------------------------
                   حالة الحضور التي ستكتب في الملف النهائي
                   --------------------------------------------- */

                let exportedStatus = '';

                if (rayatLanguage === 'ar') {

                    exportedStatus =
                        isPresent
                            ? 'حاضر'
                            : 'غائب';

                } else {

                    exportedStatus =
                        isPresent
                            ? 'Present'
                            : 'Absent';
                }


                /* ---------------------------------------------
                   الساعات
                   --------------------------------------------- */

                const expectedHours =
                    getField(
                        student,
                        'Expected Hours'
                    ) ?? '';


                let actualHours = '';
                let absentHours = '';


                if (isPresent) {

                    actualHours =
                        expectedHours;

                    absentHours =
                        '00:00';

                } else {

                    actualHours =
                        '00:00';

                    absentHours =
                        expectedHours;
                }


                /* =================================================
                   Meeting Days

                   نحافظ على القيمة الرقمية الأصلية
                   حتى يتعرف عليها Banner كتاريخ Excel.
                   ================================================= */

                let meetingDays =
                    getField(
                        student,
                        'Meeting Days'
                    );


                if (
                    typeof meetingDays !== 'number' ||
                    !Number.isFinite(meetingDays)
                ) {

                    /*
                        في حالة وصول الرقم كنص
                        نحاول تحويله إلى رقم.
                    */

                    const numericMeetingDays =
                        Number(meetingDays);


                    if (
                        Number.isFinite(
                            numericMeetingDays
                        )
                    ) {

                        meetingDays =
                            numericMeetingDays;

                    } else {

                        meetingDays = '';
                    }
                }


                /* =================================================
                   ترتيب Rayat الأصلي A إلى O
                   ================================================= */

                return [

                    getField(
                        student,
                        'Term Code'
                    ) ?? '',

                    getField(
                        student,
                        'CRN'
                    ) ?? '',

                    getField(
                        student,
                        'Session Indicator'
                    ) ?? '',

                    getField(
                        student,
                        'Start Time'
                    ) ?? '',

                    getField(
                        student,
                        'Student ID'
                    ) ?? '',

                    getField(
                        student,
                        'Student Name'
                    ) ?? '',

                    getField(
                        student,
                        'Confidential Indicator'
                    ) ?? '',

                    meetingDays,

                    expectedHours,

                    actualHours,

                    absentHours,

                    exportedStatus,

                    getField(
                        student,
                        'Authorised Absence'
                    ) ?? '',

                    getField(
                        student,
                        'Comments'
                    ) ?? '',

                    getField(
                        student,
                        'Meeting ID'
                    ) ?? ''

                ];

            });


        /* =====================================================
           إنشاء ورقة Excel
           ===================================================== */

        const worksheet =
            XLSX.utils.aoa_to_sheet(
                [
                    headers,
                    ...rows
                ]
            );


        /* =====================================================
           إصلاح Meeting Days

           العمود H هو Meeting Days
           في القالب العربي والإنجليزي.

           نضع:
           - القيمة الرقمية الأصلية.
           - نوع الخلية رقم.
           - تنسيق dd/mm/yyyy.
           ===================================================== */

        for (
            let rowNumber = 2;
            rowNumber <= rows.length + 1;
            rowNumber++
        ) {

            const cellAddress =
                `H${rowNumber}`;


            const cell =
                worksheet[cellAddress];


            if (
                cell &&
                typeof cell.v === 'number' &&
                Number.isFinite(cell.v)
            ) {

                /* القيمة تبقى رقم Excel */

                cell.t = 'n';


                /* تنسيق التاريخ */

                cell.z =
                    'dd/mm/yyyy';


                /*
                   حذف النص المنسق القديم
                   حتى لا يحتفظ SheetJS
                   بقيمة عرض سابقة.
                */

                delete cell.w;


                /*
                   إزالة Style قديم
                   لمنع تحويل التاريخ إلى Text.
                */

                delete cell.s;
            }
        }


        /* =====================================================
           عرض مناسب للأعمدة
           لا يؤثر على بيانات الاستيراد
           ===================================================== */

        worksheet['!cols'] = [

            { wch: 14 },  // A
            { wch: 12 },  // B
            { wch: 18 },  // C
            { wch: 12 },  // D
            { wch: 16 },  // E
            { wch: 28 },  // F
            { wch: 22 },  // G
            { wch: 16 },  // H
            { wch: 18 },  // I
            { wch: 16 },  // J
            { wch: 16 },  // K
            { wch: 22 },  // L
            { wch: 22 },  // M
            { wch: 22 },  // N
            { wch: 16 }   // O

        ];


        /* =====================================================
           إنشاء Workbook
           ===================================================== */

        const workbook =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            'Exported Data'
        );


        /* =====================================================
           اسم الملف
           ===================================================== */

        let crn =
            getField(
                attendanceData[0],
                'CRN'
            );


        if (!crn) {

            crn =
                'Attendance';
        }


        crn =
            String(crn).replace(
                /[\\/:*?"<>|]/g,
                '_'
            );


        const fileName =
            `Attendance_${crn}.xlsx`;


        /* =====================================================
           تصدير Excel
           ===================================================== */

        XLSX.writeFile(
            workbook,
            fileName
        );


        /* =====================================================
           رسالة نجاح
           ===================================================== */

        alert(
            'تم إنشاء ملف رايات النهائي بنجاح'
        );


    } catch (error) {

        console.error(
            'Export Error:',
            error
        );


        alert(
            'حدث خطأ أثناء تحميل ملف رايات النهائي.\n\n' +
            error.message
        );
    }
}
