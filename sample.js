
// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var http = require('http');

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
myStyle2.Fill.Color('#333333');
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
myStyle3.Font.Color('#222222');

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

var myStyle5 = wb.Style();
myStyle5.Font.Alignment.Vertical('top');
myStyle5.Font.Alignment.Horizontal('center');
myStyle5.Font.WrapText(true);
myStyle5.Fill.Pattern('solid');
myStyle5.Fill.Color('FF888888');

var ws = wb.WorkSheet('Sample Invoice');
var ws2 = wb.WorkSheet('Sample Budget');
var ws3 = wb.WorkSheet('Departmental Spending Report')
var seriesWS = wb.WorkSheet('Series with frozen Row');

/*
	Code to generate sample invoice
*/

ws.Row(1).Height(140);
ws.Cell(1,1,2,6,true);
ws.Image('sampleFiles/image1.png').Position(1,1,0,0);
ws.Row(3).Height(50);
ws.Cell(3,1,3,6,true);
ws.Row(17).Height(50);
ws.Cell(17,1,17,6,true).Style(myStyle5).String('Harvard School of Engineering and Applied Sciences\r\n29 Oxford St\r\nCambridge MA 02138\r\nhttp://www.seas.harvard.edu');
ws.Cell(4,1,15,6).Style(myStyle3);
ws.Cell(4,1,4,6).Style(myStyle2);
ws.Cell(4,1).String('Item');
ws.Cell(4,2).String('Quantity');
ws.Cell(4,3).String('Price/Unit');
ws.Cell(4,6).String('Subtotal');
ws.Cell(4,6).Format.Font.Family('Arial');
ws.Cell(4,6).Format.Font.Alignment.Horizontal('center');

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

/*
	Code to generate sample budget
*/

var myBudget = {
	Groceries:300,
	Cable:150,
	Telephone:80,
	Entertainment:200,
	Utilities:150
}
var income = {
	Employer:1500,
	FamilyTechSupport:0,
	ContractWork:500
}
var expenses = {
	Groceries:278,
	Cable:150,
	Telephone:80,
	Entertainment:350,
	Utilities:150
}
ws2.Cell(1,1).String('My Budget').Style(myStyle2);
ws2.Cell(3,1,3+Object.keys(income).length,2).Style(myStyle3);
ws2.Cell(3,1,3,2,true).String('Incomes');
ws2.Row(3).Height(30);
ws2.Cell(3,1).Format.Font.Size(18);
ws2.Cell(3,1).Format.Font.Color('FF888888');
ws2.Cell(3,1).Format.Fill.Pattern('solid');
ws2.Cell(3,1).Format.Fill.Color('FF000000');
var incomeStartRow = curRow = 4;
Object.keys(income).forEach(function(k){
	ws2.Cell(curRow,1).String(k);
	ws2.Cell(curRow,2).Number(income[k]);
	ws2.Cell(curRow,2).Format.Number("$#,##0.00");
	incomeEndRow = curRow;
	curRow+=1;
});
ws2.Cell(curRow,1).String('Total Income');
ws2.Cell(curRow,2).Formula('SUM(B'+incomeStartRow+':B'+incomeEndRow+')').Format.Number("$#,##0.00");

ws2.Cell(3,4,3+Object.keys(myBudget).length,5).Style(myStyle3);
ws2.Cell(3,4,3,5,true).String('Budget');
ws2.Cell(3,4,3,5).Format.Font.Size(18);
ws2.Cell(3,4).Format.Font.Color('FF888888');
ws2.Cell(3,4).Format.Fill.Pattern('solid');
ws2.Cell(3,4).Format.Fill.Color('FF000000');
var budgetStartRow = curRow = 4;
Object.keys(myBudget).forEach(function(k){
	ws2.Cell(curRow,4).String(k);
	ws2.Cell(curRow,5).Number(myBudget[k]);
	ws2.Cell(curRow,5).Format.Number("$#,##0.00");
	budgetEndRow = curRow;
	curRow+=1;
});
ws2.Cell(curRow,4).String('Total Budget');
ws2.Cell(curRow,5).Formula('SUM(B'+budgetStartRow+':B'+budgetEndRow+')').Format.Number("$#,##0.00");

ws2.Cell(3,7,3+Object.keys(expenses).length,8).Style(myStyle3);
ws2.Cell(3,7,3,8,true).String('Expenses');
ws2.Cell(3,7,3,8).Format.Font.Size(18);
ws2.Cell(3,7).Format.Font.Color('FF888888');
ws2.Cell(3,7).Format.Fill.Pattern('solid');
ws2.Cell(3,7).Format.Fill.Color('FF000000');
var expensesStartRow = curRow = 4;
Object.keys(expenses).forEach(function(k){
	ws2.Cell(curRow,7).String(k);
	ws2.Cell(curRow,8).Number(expenses[k]);
	ws2.Cell(curRow,8).Format.Number("$#,##0.00");
	expensesEndRow = curRow;
	curRow+=1;
});
ws2.Cell(curRow,7).String('Total Expenses');
ws2.Cell(curRow,8).Formula('SUM(B'+expensesStartRow+':B'+expensesEndRow+')').Format.Number("$#,##0.00");

ws2.Column(10).Width(16);
ws2.Cell(3,10,3+Object.keys(expenses).length,10).Style(myStyle3);
ws2.Cell(3,10).String('Differences').Format.Font.Size(18);
ws2.Cell(3,10).Format.Font.Color('FF888888');
ws2.Cell(3,10).Format.Fill.Pattern('solid');
ws2.Cell(3,10).Format.Fill.Color('FF000000');

for(var curRow=4;curRow < Object.keys(expenses).length + 5; curRow++){
	ws2.Cell(curRow,10).Formula("E"+curRow+"-H"+curRow).Format.Number("$#,##0.00");
};

/*
	Begin Departmental Spending report
*/
ws3.Settings.Outline.SummaryBelow();
ws3.Cell(1,1).String('cell A1');
ws3.Cell(1,2).String('cell B1');
ws3.Cell(2,1).String('cell A2');
ws3.Cell(2,2).String('cell B2');
ws3.Cell(3,1).String('cell A3');
ws3.Cell(3,2).String('cell B3');
ws3.Cell(4,1).String('cell A4');
ws3.Cell(4,2).String('cell B4');
ws3.Cell(5,1).String('cell A5');
ws3.Cell(5,2).String('cell B5');
ws3.Cell(6,1).String('cell A6');
ws3.Cell(6,2).String('cell B6');
ws3.Row(3).Group(1,false);
ws3.Row(6).Group(1,false);
ws3.Row(4).Group(2,true);
ws3.Row(5).Group(2,true);


for(var i = 1; i<=26; i++){
	seriesWS.Cell(i,1).Number(i);
	seriesWS.Cell(i,2).String(i.toExcelAlpha());
}
seriesWS.Row(5).Freeze(10);
seriesWS.Column(2).Freeze(5);
seriesWS.Row(2).Freeze(10);

wb.write("Excel.xlsx");

/*
http.createServer(function(req, res){
	wb.write("My Excel.xlsx",res);
}).listen(3000);
*/
