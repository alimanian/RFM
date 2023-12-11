
const userExcelFile = document.querySelector('#excel');
const afterSelectFile = document.querySelector('#afterSelectFile');
const btnProcess = document.querySelector('#btnProcess');
const recencySelect = document.querySelector('#recencySelect');
const frequencySelect = document.querySelector('#frequencySelect');
const monetarySelect = document.querySelector('#monetarySelect');
const valuableSelect = document.querySelector('#valuableSelect');
const sectionCount = document.querySelector('#sectionCount');

const recencyScoreText = "بخش‌بندی R";
const frequencyScoreText = "بخش‌بندی F";
const monetaryScoreText = "بخش‌بندی M";
const finalScoreText = "بخش‌بندی نهایی";

recencySelect.addEventListener('change', () => {
    valuableSelect.selectedIndex = recencySelect.selectedIndex;
});

userExcelFile.addEventListener('change', () => {
    readExcelFileForGetHeaders(userExcelFile.files[0]).then(headers => {
        fillSelectTags(headers);
    });

    afterSelectFile.style.display = 'flex';
});

btnProcess.addEventListener('click', () => {
    // type true = asc, false = desc
    sortExcelAndCreateNewExcel(userExcelFile.files[0],
        recencySelect.options[recencySelect.selectedIndex].text,
        frequencySelect.options[frequencySelect.selectedIndex].text,
        monetarySelect.options[monetarySelect.selectedIndex].text,
        valuableSelect.options[valuableSelect.selectedIndex].text).then(buffer => {

        // ایجاد لینک دانلود
        const url = URL.createObjectURL(new Blob([buffer]));
        document.getElementById("downloadExcel").href = url;
        document.getElementById("downloadExcel").style.display = "block";
    });
});

function sortExcelAndCreateNewExcel(file, recency, frequency, monetary, valuable) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {

            const data = e.target.result;

            // خواندن فایل
            const workbook = XLSX.read(data, {type: 'binary'});
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet);

            // مرتب سازی
            rows.sort((a, b) => {
                if (a[recency] < b[recency]) return -1;
                if (a[recency] > b[recency]) return 1;
                return 0;
            });

            // محاسبه تعداد کل سطرها
            const totalRows = rows.length;

            // محاسبه تعداد سطرهای هر گروه
            const rowsPerGroup = Math.floor(totalRows / sectionCount.value);

            // اضافه کردن ستون جدید
            for(let i= 0; i < totalRows; i++) {

                // محاسبه شماره گروه
                let groupNum = sectionCount.value - Math.floor((i + 1) / rowsPerGroup);
                if (groupNum < 1) groupNum = 1;
                // اضافه کردن شماره گروه به آخرین ستون
                Object.assign(rows[i], {[recencyScoreText]: groupNum});
            }

            rows.sort((a, b) => {
                if (b[frequency] < a[frequency]) return -1;
                if (b[frequency] > a[frequency]) return 1;
                return 0;
            });

            for(let i= 0; i < totalRows; i++) {
                // محاسبه شماره گروه
                let groupNum = sectionCount.value - Math.floor((i + 1) / rowsPerGroup);
                if (groupNum < 1) groupNum = 1;

                // اضافه کردن شماره گروه به آخرین ستون
                Object.assign(rows[i], {[frequencyScoreText]: groupNum});
            }

            rows.sort((a, b) => {
                if (b[monetary] < a[monetary]) return -1;
                if (b[monetary] > a[monetary]) return 1;
                return 0;
            });

            for(let i= 0; i < totalRows; i++) {
                // محاسبه شماره گروه
                let groupNum = sectionCount.value - Math.floor((i + 1) / rowsPerGroup);
                if (groupNum < 1) groupNum = 1;


                // اضافه کردن شماره گروه به آخرین ستون
                Object.assign(rows[i], {[monetaryScoreText]: groupNum});
            }

            let {firstColumnSort, secondColumnSort, thirdColumnSort} = {};

            switch (valuable) {
                case recency:
                    firstColumnSort = recencyScoreText;
                    secondColumnSort = frequencyScoreText;
                    thirdColumnSort = monetaryScoreText;
                    break;
                case frequency:
                    firstColumnSort = frequencyScoreText;
                    secondColumnSort = recencyScoreText;
                    thirdColumnSort = monetaryScoreText;
                    break;
                case monetary:
                    firstColumnSort = monetaryScoreText;
                    secondColumnSort = recencyScoreText;
                    thirdColumnSort = frequencyScoreText;
            }

            rows.sort((a, b) => {

                // اولویت ستون A
                if (a[firstColumnSort] > b[firstColumnSort]) return -1;
                if (a[firstColumnSort] < b[firstColumnSort]) return 1;

                // در صورت مساوی، اولویت ستون B
                if (a[secondColumnSort] > b[secondColumnSort]) return -1;
                if (a[secondColumnSort] < b[secondColumnSort]) return 1;

                // سپس ستون C
                if (a[thirdColumnSort] > b[thirdColumnSort]) return -1;
                if (a[thirdColumnSort] < b[thirdColumnSort]) return 1;

                return 0;
            });

            for(let i= 0; i < totalRows; i++) {

                // محاسبه شماره گروه
                let groupNum = sectionCount.value - Math.floor((i + 1) / rowsPerGroup);
                if (groupNum < 1) groupNum = 1;

                // اضافه کردن شماره گروه به آخرین ستون
                Object.assign(rows[i], {[finalScoreText]: groupNum});
            }

            // ایجاد ورک بوک جدید
            const newWorkbook = XLSX.utils.book_new();

            // تبدیل به ورکشیت و افزودن به ورک بوک جدید
            const newWorksheet = XLSX.utils.json_to_sheet(rows);
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sorted Data");

            // نوشتن ورک بوک جدید
            const buffer = XLSX.write(newWorkbook, {bookType: 'xlsx', type: 'buffer'});

            resolve(buffer);
        }

        reader.readAsArrayBuffer(file);
    });

}

function readExcelFileForGetHeaders(file) {

    return new Promise((resolve, reject) => {

        const reader = new FileReader();

        reader.onload = function (e) {

            let data = e.target.result;

            let workbook = XLSX.read(data, {type: "array"});
            let sheetName = workbook.SheetNames[0];
            let worksheet = workbook.Sheets[sheetName];

            let headers = [];
            let range = worksheet['!ref'];
            let colCount =
                range.split(":")[1].charCodeAt(0) - range.split(":")[0].charCodeAt(0);

            for (let i = 0; i <= colCount; i++) {

                let addr = XLSX.utils.encode_cell({c: i, r: 0});
                if (!worksheet[addr]) continue;
                headers.push(worksheet[addr].v);

            }

            resolve(headers);
        };

        reader.readAsArrayBuffer(file);

    });

}

function fillSelectTags(headers) {

    headers.forEach(function (h, i) {

        let col = XLSX.utils.encode_col(i);

        recencySelect.appendChild(
            new Option(h, col)
        );

        frequencySelect.appendChild(
            new Option(h, col)
        );

        monetarySelect.appendChild(
            new Option(h, col)
        );

        valuableSelect.appendChild(
            new Option(h, col)
        );
    });
}
