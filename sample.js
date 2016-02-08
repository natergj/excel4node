var xl = require('./source/dev.js');

var wb = new xl.WorkBook();
wb.toString();
wb.write('Excel.xlsx');
