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
           تجهيز الصفوف
           ===================================================== */

        const rows =
            attendanceData.map(student => {

                const status =
                    student['Attendance Indicator'];


                let actualHours = '';
                let absentHours = '';


                /* =================================================
                   الحضور
                   ================================================= */

                if (status === 'Present') {

                    actualHours =
                        student['Expected Hours'] ?? '';

                    absentHours =
                        '00:00';

                }


                /* =================================================
                   الغياب
                   ================================================= */

                else {

                    actualHours =
                        '00:00';

                    absentHours =
                        student['Expected Hours'] ?? '';
                }


                /* =================================================
                   Meeting Days
                   
                   مهم جدًا:
                   لا نحول القيمة إلى Date.
                   نحتفظ بالقيمة الأصلية مثل 46285.
                   ================================================= */

                const meetingDays =
                    student['Meeting Days'] ?? '';


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
            XLSX.utils.aoa_to_sheet([
                headers,
                ...rows
            ]);


        /* =====================================================
           تطبيق تنسيق Text (@)
           
           ملف Rayat المرجعي يستخدم @ على الأعمدة.
           ===================================================== */

        for (
            let rowNumber = 1;
            rowNumber <= rows.length + 1;
            rowNumber++
        ) {

            for (
                let columnNumber = 0;
                columnNumber < headers.length;
                columnNumber++
            ) {

                const cellAddress =
                    XLSX.utils.encode_cell({
                        r: rowNumber - 1,
                        c: columnNumber
                    });


                const cell =
                    worksheet[cellAddress];


                if (cell) {

                    cell.z = '@';
                }
            }
        }


        /* =====================================================
           التأكد من أن Meeting Days يبقى رقمًا
           
           مثل ملف Rayat الأصلي:
           H2 = 46278
           H3 = 46278
           ...
           
           مع تنسيق @
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


            if (cell) {

                const value =
                    attendanceData[rowNumber - 2]['Meeting Days'];


                if (
                    typeof value === 'number' &&
                    Number.isFinite(value)
                ) {

                    cell.v = value;

                    cell.t = 'n';

                    cell.z = '@';
                }
            }
        }


        /* =====================================================
           التأكد من Meeting ID
           
           أيضًا يبقى رقمًا مع تنسيق @
           مثل ملف Rayat الأصلي.
           ===================================================== */

        for (
            let rowNumber = 2;
            rowNumber <= rows.length + 1;
            rowNumber++
        ) {

            const cellAddress =
                `O${rowNumber}`;


            const cell =
                worksheet[cellAddress];


            if (cell) {

                const value =
                    attendanceData[rowNumber - 2]['Meeting ID'];


                if (
                    typeof value === 'number' &&
                    Number.isFinite(value)
                ) {

                    cell.v = value;

                    cell.t = 'n';

                    cell.z = '@';
                }
            }
        }


        /* =====================================================
           إنشاء ملف Excel
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
           إنشاء ملف Excel
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
