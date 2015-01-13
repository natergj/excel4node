// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var wb = new xl.WorkBook();
wb.debug=false;

var myStyle = wb.Style();
myStyle.Font.Family('Helvetica');
myStyle.Font.Bold();

ws = wb.WorkSheet('First');

ws.Row(1).Height(14);
ws.Cell(1,1).Style(myStyle);
ws.Cell(1,1).String('My String');


wb.write('basic.xlsx',process.exit);
