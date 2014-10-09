var xl = require('./lib/index.js'),
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
myStyle.Font.Color('FFFF0000');
myStyle.Number.Format("$#,##0.00;($#,##0.00);-");
myStyle.Border(
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

var myStyle2 = wb.Style();
myStyle2.Font.Size(16);
myStyle2.Font.Color('FFABABAB');
myStyle2.Fill.Pattern('solid');
myStyle2.Fill.Color('FF000000');
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
myStyle3.Font.Alignment.Vertical('top');
myStyle3.Font.Alignment.Horizontal('left');
myStyle3.Font.WrapText(true);
myStyle3.Font.Color('FF222222');

var myStyle4 = wb.Style();
myStyle4.Border(
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
myStyle4.Font.Alignment.Horizontal('right');
myStyle4.Font.Color('FF222222');

var ws = wb.WorkSheet('Invoice');
var ws2 = wb.WorkSheet('my 2nd worksheet');
var ws3 = wb.WorkSheet('my 3rd worksheet');

/*
	Code to generate sample invoice
*/

ws.Row(1).Height(140);
ws.Cell(1,1,2,6,true);
ws.Image('sampleFiles/image1.png').Position(1,1,0,0);
ws.Row(3).Height(50);
ws.Cell(3,1,3,6,true).Style(myStyle3).String('Harvard School of Engineering and Applied Sciences\r\n29 Oxford St\r\nCambridge MA 02138');
ws.Cell(4,1,15,6).Style(myStyle3);
ws.Cell(4,1,4,6).Style(myStyle2);
ws.Cell(4,1).String('Item');
ws.Cell(4,2).String('Quantity');
ws.Cell(4,3).String('Price/Unit');
ws.Cell(4,6).String('Subtotal');

var columnDefinitions = {
	item:1,
	quantity:2,
	cost:3,
	total:6
}

var invoiceItems = [
	{
		item: 'Soft Robot',
		quantity: 5,
		costPerUnit: 250.25
	},
	{
		item: 'Quantum Transistor',
		quantity:5,
		costPerUnit:500
	},
	{
		item:'Mountain Water Well',
		quantity:2,
		costPerUnit:50
	}
]

var curRow = 5;
invoiceItems.forEach(function(i){
	ws.Cell(curRow,columnDefinitions.item).String(i.item);
	ws.Cell(curRow,columnDefinitions.quantity).Number(i.quantity);
	ws.Cell(curRow,columnDefinitions.cost).Number(i.costPerUnit);
	ws.Cell(curRow,columnDefinitions.total).Formula(columnDefinitions.quantity.toExcelAlpha()+curRow+"*"+columnDefinitions.cost.toExcelAlpha()+curRow).Format.Number("$#,#00.00");
	curRow+=1;
});

ws.Cell(16,1,16,5,true).Style(myStyle4).String('Total');
ws.Cell(16,6).Style(myStyle4).Formula('SUM('+columnDefinitions.total.toExcelAlpha()+'5:'+columnDefinitions.total.toExcelAlpha()+'15)').Format.Number("$#,#00.00");


var img2 = ws2.Image('sampleFiles/image2.jpg').Position(4,4,500000,500000);
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
