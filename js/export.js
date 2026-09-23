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
           ترتيب أعمدة Rayat الأصلي
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


                /*
                   مهم:

                   هنا نضع Meeting Days كرقم فقط.
                   وبعد إنشاء الورقة سنفرض تنسيق
                   التاريخ على خلايا H بشكل مباشر.
                */

                let meetingDays =
                    student['Meeting Days'];

                if (
                    typeof meetingDays !== 'number' ||
                    !Number.isFinite(meetingDays)
                ) {

                    meetingDays = '';

                }


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
           =====================================================

           العمود H هو Meeting Days.

           نضع:
           - القيمة الرقمية الأصلية
           - تنسيق تاريخ Excel
           - نحذف w حتى لا يحتفظ SheetJS
             بالتنسيق القديم.
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
                   مهم جدًا:
                   حذف النص المنسق القديم
                */

                delete cell.w;


                /*
                   إزالة أي Style قديم
                   حتى لا يعيد تنسيق @
                */

                delete cell.s;
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
           اسم الملف
           ===================================================== */

        let crn =
            attendanceData[0]['CRN'];


        if (!crn) {

            crn = 'Attendance';
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
           =====================================================

           ملاحظة مهمة:
           لا نستخدم cellStyles هنا.
           تنسيق z هو الذي يجب أن يحمل
           تنسيق التاريخ.
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
