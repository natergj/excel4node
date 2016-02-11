var xl = require('./source/dev.js');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('My First Sheet', {
	printOptions : {
		fitToWidth : 1
	}
});
ws.Cell(1,1).String('something');

wb.write('Excel.xlsx');
