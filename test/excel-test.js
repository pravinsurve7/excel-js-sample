var excel = require("./ExcelUtility");

excel.readExcel("./test/data.xlsx", (err, data) => {
    if (err) {
        console.log(err);
    } else {
        console.log(data);
    }
});