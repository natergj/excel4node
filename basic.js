// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var wbOpts = {
	colWidth: 25
}

var wb = new xl.WorkBook();
wb.debug=true;

var myStyle = wb.Style();
myStyle.Font.Family('Helvetica');
myStyle.Font.Bold();
myStyle.Font.Alignment.Horizontal('right');
myStyle.Font.WrapText();

ws = wb.WorkSheet('First');

ws.Row(1).Height(14);
ws.Column(1).Width(50);
ws.Cell(1,1).Style(myStyle);
ws.Cell(1,1).String('My String');
ws.Cell(2,2).String('my\nwrapped\nstring').Style(myStyle);

wb.write('basic.xlsx',process.exit);
