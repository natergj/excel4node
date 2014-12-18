// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var wb = new xl.WorkBook();
wb.debug=false;

ws = wb.WorkSheet('First');

ws.Row(1).Height(14);
ws.Cell(3,2);

console.log(ws);
process.exit();