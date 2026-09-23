function excelSerialToDate(serial) {

    const numericValue = Number(serial);

    if (!Number.isFinite(numericValue)) {
        return serial;
    }

    return new Date(
        Date.UTC(1899, 11, 30) +
        numericValue * 24 * 60 * 60 * 1000
    );
}


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
           تجهيز بيانات الحضور
           ===================================================== */

        const rows =
            attendanceData.map(student => {

                const status =
                    student['Attendance Indicator'];


                let actualHours = '';
                let absentHours = '';


                if (status === 'Present') {

                    actualHours =
                        student['Expected Hours'];

                    absentHours =
                        '00:00';

                } else {

                    actualHours =
                        '00:00';

                    absentHours =
                        student['Expected Hours'];
                }


                /* =================================================
                   تحويل Meeting Days
                   من Excel Serial مثل 46285
                   إلى تاريخ فعلي
                   ================================================= */

                let meetingDays =
                    student['Meeting Days'];


                if (
                    typeof meetingDays === 'number' &&
                    Number.isFinite(meetingDays)
                ) {

                    meetingDays =
                        excelSerialToDate(
                            meetingDays
                        );
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
            XLSX.utils.aoa_to_sheet([
                headers,
                ...rows
            ]);


        /* =====================================================
           ضبط Meeting Days كتاريخ Excel حقيقي
           ===================================================== */

        for (
            let rowNumber = 2;
            rowNumber <= rows.length + 1;
            rowNumber++
        ) {

            const cell =
                worksheet[`H${rowNumber}`];


            if (
                cell &&
                cell.v instanceof Date
            ) {

                cell.t = 'd';

                cell.z =
                    'mm/dd/yyyy';
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
           تحميل الملف
           ===================================================== */

        XLSX.writeFile(
            workbook,
            fileName,
            {
                cellDates: true
            }
        );


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
