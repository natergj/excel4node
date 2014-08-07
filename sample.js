var xl = require('excel4node'),
http = require('http');

function getAllMethods(object) {
    return Object.getOwnPropertyNames(object).filter(function(property) {
        return typeof object[property] == 'function';
    });
}



var wb = new xl.WorkBook();
wb.debug=false;

var myStyle = wb.Style();
myStyle.Font.Underline();
myStyle.Font.Bold();
myStyle.Font.Italics();
myStyle.Font.Size(16);
myStyle.Font.Family('Helvetica');
myStyle.Font.Color('FF0000');
myStyle.Number.Format("$#,##0.00;($#,##0.00);-");

var myStyle2 = wb.Style();
myStyle2.Font.Size(18);
myStyle2.Font.Color('ABABAB');
myStyle2.Fill.Pattern('solid');
myStyle2.Fill.Color('FF54FF');

var myStyle3 = wb.Style();
myStyle3.Border(
	{
		left:{
			style:'thin'
		},
		right:{
			style:'thin'
		},
		top:{
			style:'thin'
		},
		bottom:{
			style:'thin'
		}
	}
);
myStyle3.Font.Color('222222');
myStyle3.Number.Format("##%");


var ws = wb.WorkSheet('my worksheet');
var ws2 = wb.WorkSheet('my 2nd worksheet');
var ws3 = wb.WorkSheet('my 3rd worksheet');

var img = ws.Image('sampleFiles/image1.png');
var img2 = ws2.Image('sampleFiles/image2.jpg').Position(4,4,500000,500000);

ws.Row(1).Height(60);
ws.Column(1).Width(120);
ws.Cell(1,1).String('Cell A1').Style(myStyle2);
ws.Cell(1,2).String('Cell B1');
ws.Cell(1,3).String('newValue');
ws.Cell(1,4).String('newValue');
ws.Cell(1,5).String('newValue');
ws.Cell(2,5).String('2ndValue');
ws.Cell(2,1).Number(100).Style(myStyle);
ws2.Cell(1,4).String('cell data');
ws2.Cell(2,1).Number(5);
ws2.Cell(2,1).Style(myStyle);
ws2.Cell(2,2).Number(10).Style(myStyle);
ws2.Cell(2,3).Formula("A2-B2").Style(myStyle);
ws2.Cell(2,4).Formula("A2/B2").Style(myStyle3);

wb.write("Excel.xlsx");

/*
http.createServer(function(req, res){
	wb.write("My Excel.xlsx",res);
}).listen(3000);
*/