var xl = require('./distribution');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('My First Sheet', {
    printOptions: {
        fitToWidth: 1
    }
});

wb.write('Excel.xlsx');
