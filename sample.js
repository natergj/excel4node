var xl = require('./source/dev.js');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('My First Sheet', {
	pageSetup : {
		fitToWidth : 1
	}
});

wb.write('Excel.xlsx');
