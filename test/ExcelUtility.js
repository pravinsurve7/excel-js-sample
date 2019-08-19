var Excel = require('exceljs');

var jsonData = {};
var readExcel = (excelPath, callback) => {
    var workbook = new Excel.Workbook();
    return workbook.xlsx.readFile(excelPath).then(() => {
        workbook.eachSheet((sheet) => {
            var arrayOfRows = [];
            console.log("Sheet name: " + sheet.name);

            var rowCount = sheet.actualRowCount
            console.log("Row count: " + (rowCount - 1));

            var columnCount = sheet.columnCount;
            console.log("Column count: " + columnCount);

            for (let i = 2; i <= rowCount; i++) {
                var rowObj = {};
                let row = sheet.getRow(i);
                for (let j = 1; j <= columnCount; j++) {
                    rowObj[sheet.getRow(1).getCell(j).value] = row.getCell(j).value;
                }
                arrayOfRows.push(rowObj);
                jsonData[sheet.name] = arrayOfRows;
            }
        });
        callback(null, jsonData);
    }).catch((err) => {
        callback(err);
    });
};

var writeExcel = (excelPath, jsonData, callback) => {
    var workbook = new Excel.Workbook();
    var worksheet = workbook.addWorksheet("test");
    // Add a couple of Rows by key-value, after the last current row, using the column keys
    // Add a row by contiguous Array (assign to columns A, B & C)
    worksheet.addRow([3, 'Sam', new Date()]);
    worksheet.addRow({
        id: 1,
        name: 'John Doe',
        dob: new Date(1970, 1, 1)
    });
    worksheet.addRow({
        id: 2,
        name: 'Jane Doe',
        dob: new Date(1965, 1, 7)
    });

    // Add an array of rows
    var rows = [
        [5, 'Bob', new Date()], // row by array
        {
            id: 6,
            name: 'Barbara',
            dob: new Date()
        }
    ];
    worksheet.addRows(rows);


    // for(worksheet.actualRowCount())
    var cell = worksheet.getCell('C3');

    // Modify/Add individual cell

    cell.value = new Date(1968, 5, 1);

    workbook.xlsx.writeFile('./report.xlsx');
    callback(null, "success");
};

module.exports = {
    readExcel,
    writeExcel
}