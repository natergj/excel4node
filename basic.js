// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var wb = new xl.WorkBook();
wb.debug=false;

ws = wb.WorkSheet('First');

ws.Cell(2,1).Number(4);
ws.Cell(2,2).Number(6);
ws.Cell(2,3).Formula("SUM(A2:B2)");
ws.Cell(3,2).String('My String');
ws.Cell(4,3).Date(new Date(2015,0,10));

//console.log(JSON.stringify(ws.Cell(3,2),null,'\t'));
//console.log(ws.toXML());

wb.write('text.xlsx',process.exit);
