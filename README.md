# excel4node
A full featured xlsx file generation library allowing for the creation of advanced Excel files.

### Basic Usage
```javascript
var xl = require('excel4node');
var wb = new xl.WorkBook();
var ws = wb.WorkSheet('Sheet 1');
var ws2 = wb.WorkSheet('Sheet 2');

var style = wb.Style({
	font: {
		color: '#FF0800',
		size: 12
	},
	numberFormat: '$#,##0.00; ($#,##0.00); -'
});

ws.Cell(1,1).Number(100).Style(style);
ws.Cell(1,2).Number(200).Style(style);
ws.Cell(1,3).Formula('A1 + B1').Style(style);
ws.Cell(2,1).String('string').Style(style);
ws.Cell(3,1).Bool(true).Style(style).Style({font: {size: 14}});

wb.write('Excel.xlsx');
```