function exportAttendanceFile() {

    try {

        /* =====================================================
           التحقق من وجود ملف الحضور النهائي
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


        /* =====================================================
           قراءة البيانات
           ===================================================== */

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
           أعمدة Rayat الأصلية
           ===================================================== */

        const headers = [

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
           إنشاء الصفوف
           ===================================================== */

        const rows =
            attendanceData.map(student => {

                const status =
                    student['Attendance Indicator'];


                /* =================================================
                   حساب Actual Hours و Absent Hours
                   ================================================= */

                let actualHours = '';
                let absentHours = '';


                if (status === 'Present') {

                    actualHours =
                        student['Expected Hours'] ?? '';

                    absentHours =
                        '00:00';

                } else {

                    actualHours =
                        '00:00';

                    absentHours =
                        student['Expected Hours'] ?? '';
                }


                /* =================================================
                   Meeting Days

                   مهم جدًا:

                   نحتفظ برقم Excel الأصلي
                   ولكن نرسله كخلية رقمية مع
                   تنسيق تاريخ حقيقي.

                   مثال:
                   46285
                   سيظهر في Excel:
                   20/09/2026
                   ================================================= */

                let meetingDays =
                    student['Meeting Days'];


                if (
                    typeof meetingDays === 'number' &&
                    Number.isFinite(meetingDays)
                ) {

                    meetingDays = {

                        v: meetingDays,

                        t: 'n',

                        z: 'dd/mm/yyyy'

                    };

                } else {

                    meetingDays =
                        meetingDays ?? '';
                }


                /* =================================================
                   الصف النهائي
                   ================================================= */

                return [

                    student['Term Code'] ?? '',

                    student['CRN'] ?? '',

                    student['Session Indicator'] ?? '',

                    student['Start Time'] ?? '',

                    student['Student ID'] ?? '',

                    student['Student Name'] ?? '',

                    student['Confidential Indicator'] ?? '',

                    meetingDays,

                    student['Expected Hours'] ?? '',

                    actualHours,

                    absentHours,

                    status ?? '',

                    student['Authorised Absence'] ?? '',

                    student['Comments'] ?? '',

                    student['Meeting ID'] ?? ''

                ];

            });


        /* =====================================================
           إنشاء ورقة Excel

           Cell objects تستخدم كما هي،
           ولذلك Meeting Days سيحتفظ بتنسيق التاريخ.
           ===================================================== */

        const worksheet =
            XLSX.utils.aoa_to_sheet(
                [
                    headers,
                    ...rows
                ]
            );


        /* =====================================================
           التأكد مرة أخرى من تنسيق Meeting Days

           العمود H
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
                typeof cell.v === 'number'
            ) {

                cell.t = 'n';

                cell.z =
                    'dd/mm/yyyy';
            }
        }


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
           الحصول على CRN
           ===================================================== */

        let crn =
            attendanceData[0]['CRN'];


        if (!crn) {

            crn = 'Attendance';
        }


        /* =====================================================
           تنظيف اسم الملف
           ===================================================== */

        crn =
            String(crn).replace(
                /[\\/:*?"<>|]/g,
                '_'
            );


        const fileName =
            `Attendance_${crn}.xlsx`;


        /* =====================================================
           كتابة ملف Excel
           ===================================================== */

        XLSX.writeFile(
            workbook,
            fileName,
            {
                cellStyles: true
            }
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
