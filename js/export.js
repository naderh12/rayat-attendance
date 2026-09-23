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
           تجهيز ساعات الحضور والغياب
           ===================================================== */

        attendanceData.forEach(student => {

            const status =
                student["Attendance Indicator"];


            if (status === "Present") {

                student["Actual Hours"] =
                    student["Expected Hours"];

                student["Absent Hours"] =
                    "00:00";

            } else {

                student["Actual Hours"] =
                    "00:00";

                student["Absent Hours"] =
                    student["Expected Hours"];
            }
        });


        /* =====================================================
           إنشاء ورقة Excel
           ===================================================== */

        const worksheet =
            XLSX.utils.json_to_sheet(
                attendanceData
            );


        /* =====================================================
           إنشاء ملف Excel
           ===================================================== */

        const workbook =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            "Attendance"
        );


        /* =====================================================
           الحصول على CRN
           ===================================================== */

        let crn =
            attendanceData[0]["CRN"];


        if (!crn) {
            crn = "Attendance";
        }


        /* =====================================================
           تنظيف اسم الملف من الرموز غير المسموح بها
           ===================================================== */

        crn = String(crn).replace(
            /[\\/:*?"<>|]/g,
            "_"
        );


        const fileName =
            `Attendance_${crn}.xlsx`;


        /* =====================================================
           تحميل ملف Excel
           ===================================================== */

        XLSX.writeFile(
            workbook,
            fileName
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
