
// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var http = require('http');

function createWorkBook(){
	var options = {
		jszip:{
			compression:'DEFLATE'
		},
		allowInterrupt : false
	}
	var wb = new xl.WorkBook(options);
	wb.debug=process.argv[2]=='debug'?true:false;

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

	var rotateStyle = wb.Style();
	rotateStyle.Font.Alignment.Rotation('90');
	rotateStyle.Font.Alignment.Vertical('center');
	rotateStyle.Font.Alignment.Horizontal('center');

	var hyperLinkStyle = wb.Style();
	hyperLinkStyle.Font.Color('Blue');
	hyperLinkStyle.Font.Underline();

	var wsOpts = {
		margins:{
			left: .75,
			right: .75,
			top: 1.0,
			bottom: 1.0,
			footer: .5,
			header: .5
		},
		printOptions:{
			centerHorizontal: true,
			centerVertical: false
		},
		view:{
			zoom: 100
		},
		fitToPage:{
			fitToHeight:100,
			orientation:'landscape'
		}
	}
	var pWSOpts = {
		sheetProtection : {
			password : 'password',
			formatRows : true,
			selectUnlockedCells : true
		}
	}

	var gWsOpts = {
		outline : {
			summaryBelow : false
		}
	}
	var g2WsOpts = {
		outline : {
			summaryBelow : true
		}
	}

	var ws = wb.WorkSheet('Sample Invoice',wsOpts);
	var ws2 = wb.WorkSheet('Sample Budget',wsOpts);
	var ws3 = wb.WorkSheet('Auto Filter', {
			orientation:'landscape'
	});
	var seriesWS = wb.WorkSheet('Series with frozen Row',pWSOpts);
	var groupings = wb.WorkSheet('Groupings Summary Top',gWsOpts);
	var groupings2 = wb.WorkSheet('Groupings Summary Bottom',g2WsOpts);
	var stylings = wb.WorkSheet('Stylings');
	var validations = wb.WorkSheet('Validations');

	/*
		Code to generate sample invoice
	*/

	ws.Row(1).Height(140);
	ws.Cell(1,1,2,6,true);

	//Absolute position near D3
	ws.Image('sampleFiles/image1.png', ws.Image.ABSOLUTE).Position(218, 400).Size(255, 50);  //y-pixels, X-pixels
	
	//A3
	ws.Image('sampleFiles/image1.png', ws.Image.ONE_CELL).Position(3, 1, 10, 40).Size(255, 50); //row, column, offsetY, offsetX
	
	//A1-F2
	ws.Image('sampleFiles/image1.png', ws.Image.TWO_CELL).Position(1, 1, 2, 6, 2, 5); //begin-row, begin-column, end-row, end-column, offsetY, offsetX
	
	//D5
	ws.Image('sampleFiles/image1.png', ws.Image.TWO_CELL).Position(5, 4); //row, column (in one cell re-scale in cell)


	ws.Row(3).Height(50);
	ws.Cell(3,1,3,6,true);
	ws.Row(17).Height(60);
	ws.Cell(17,1,17,6,true).Style(myStyle5).String('Harvard School of Engineering and Applied Sciences\n29 Oxford St\nCambridge MA 02138\nhttp://www.seas.harvard.edu');
	ws.Cell(18,1,18,6,true).Link('http://www.seas.harvard.edu','The Harvard John A Paulson School of Engineering and Applied Sciences').Style(hyperLinkStyle);
	ws.Cell(4,1,15,6).Style(myStyle3);
	ws.Cell(4,1,4,6).Style(myStyle2);
	ws.Cell(4,1).String('Item');
	ws.Column(1).Width(20);
	ws.Cell(4,2).String('Quantity');
	ws.Column(2).Width(10);
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
		ws.Cell(curRow,columnDefinitions.quantity).Number(i.quantity).Format.Font.Alignment.Horizontal('right');;
		ws.Cell(curRow,columnDefinitions.cost).Number(i.costPerUnit).Format.Number("$#,##0.00");
		ws.Cell(curRow,columnDefinitions.total)
			.Formula(columnDefinitions.quantity.toExcelAlpha()+curRow+"*"+columnDefinitions.cost.toExcelAlpha()+curRow)
			.Format.Number("$#,#00.00")
			.Format.Font.Alignment.Horizontal('right');
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

	var deptSpending = {
		'Dept 1':{
			'food':100.00,
			'coffee':200.00,
			'paper':50.00,
			'supplies':75.00
		},
		'Dept 2':{
			'food':120.00,
			'coffee':100.00,
			'paper':10.00,
			'supplies':25.00
		},
		'Dept 3':{
			'food':400.00,
			'coffee':300.00,
			'paper':0.00,
			'supplies':5.00
		}
	}
	var filters = [
		{
			column:2,
			rules:[
				{
					val:'food'
				},
				{
					val:'coffee'
				}
			]
		}
	];

	ws3.Row(1).Filter(1,3,filters);

	ws3.Cell(1,1).String('Department');
	ws3.Cell(1,2).String('Item');
	ws3.Cell(1,3).String('Amount');
	curRow = 2;
	Object.keys(deptSpending).forEach(function(dept){
		Object.keys(deptSpending[dept]).forEach(function(desc){
			ws3.Cell(curRow,1).String(dept);
			ws3.Cell(curRow,2).String(desc);
			ws3.Cell(curRow,3).Number(deptSpending[dept][desc]);
			curRow+=1;
		});
	});

	var startRow = 15,
		boldStyle = wb.Style();
	boldStyle.Font.Bold();
	for (var r = 0; r < 5; r++) {
		for (var c = 1; c < 100; c++) {
			if (r === 0) {
				ws3.Cell(startRow + r, c).Style(boldStyle).String(c.toExcelAlpha());
			} else {
				ws3.Cell(startRow + r, c).String(c.toExcelAlpha() + r);
			}
		}
	}

	ws3.printTitles({
        rows: {
            begin: startRow,
            end: startRow
        },
        columns: {
            begin: 1,
            end: 3
        }
    });

	var now = new Date();
	for(var i = 1; i<=26; i++){
		seriesWS.Cell(i,1).Number(i);
		seriesWS.Cell(i,2).String(i.toExcelAlpha());
		now.setDate(now.getDate()+1);
		seriesWS.Cell(i,3).Date(now);
	}
	seriesWS.Row(5).Freeze(10);
	seriesWS.Column(2).Freeze(5);
	seriesWS.Row(2).Freeze(10);


	var groupingData = {
		'GrandParentA':{
			'ParentAA':{
				'ChildAAA':[
					'ToyAAA1',
					'ToyAAA2',
					'ToyAAA3'
				],
				'ChildAAB':[
					'ToyAAB1',
					'ToyAAB2',
					'ToyAAB3'
				]
			},
			'ParentAB':{
				'ChildABA':[
					'ToyABA1',
					'ToyABA2',
					'ToyABA3'
				],
				'ChildABB':[
					'ToyABB1',
					'ToyABB2',
					'ToyABB3'
				]
			}
		},
		'GrandParentB':{
			'ParentBA':{
				'ChildBAA':[
					'ToyBAA1',
					'ToyBAA2',
					'ToyBAA3'
				],
				'ChildBAB':[
					'ToyBAB1',
					'ToyBAB2',
					'ToyBAB3'
				]
			},
			'ParentBB':{
				'ChildBBA':[
					'ToyBBA1',
					'ToyBBA2',
					'ToyBBA3'
				],
				'ChildBBB':[
					'ToyBBB1',
					'ToyBBB2',
					'ToyBBB3'
				]
			}
		}
	}

	curRow = 1;
	Object.keys(groupingData).forEach(function(g,i){
		var isCollapsed = i==0?true:false;
		groupings.Cell(curRow,1).String(g);
		curRow+=1;
		Object.keys(groupingData[g]).forEach(function(p){
			groupings.Cell(curRow,2).String(p);
			groupings.Row(curRow).Group(1,isCollapsed);
			curRow+=1;
			Object.keys(groupingData[g][p]).forEach(function(c){
				groupings.Cell(curRow,3).String(c);
				groupings.Row(curRow).Group(2,isCollapsed);
				curRow+=1;
				groupingData[g][p][c].forEach(function(t){
					groupings.Cell(curRow,4).String(t);
					groupings.Row(curRow).Group(3,isCollapsed);
					curRow+=1;
				})
			})
		})
	});


	curRow = 1;
	Object.keys(groupingData).forEach(function(g,i){
		//var isCollapsed = i==0?true:false;
		var isCollapsed = true;
		Object.keys(groupingData[g]).forEach(function(p){
			Object.keys(groupingData[g][p]).forEach(function(c){
				groupingData[g][p][c].forEach(function(t){
					groupings2.Cell(curRow,4).String(t);
					groupings2.Row(curRow).Group(3,isCollapsed);
					curRow+=1;
				})
				groupings2.Cell(curRow,3).String(c);
				groupings2.Row(curRow).Group(2,isCollapsed);
				curRow+=1;
			})
			groupings2.Cell(curRow,2).String(p);
			groupings2.Row(curRow).Group(1,isCollapsed);
			curRow+=1;
		})
		groupings2.Cell(curRow,1).String(g);
		curRow+=1;
	});


	// Style sheet
	stylings.Column(1).Width(3);
	stylings.Cell(1,1,5,1,true).String('Group 1').Style(rotateStyle);
	stylings.Cell(6,1,10,1,true).String('Group 2').Style(rotateStyle);
	for(var i = 1; i <= 10; i++){
		stylings.Cell(i,2).String('Member '+i);
	}
	stylings.Row(2).Hide();
	stylings.Column(3).Hide();
	stylings.Cell(1,4).String('Row 2 and Column 3 are hiding');
	stylings.Cell(3,4).Bool('false');
	stylings.Cell(4,4).Bool(false);
	stylings.Cell(5,4,7,4).Bool('true');
	stylings.Cell(8,4).Bool(true);
	stylings.Cell(9,4,10,5,true).Bool(true);

	stylings.Cell(3,6).String('yes');
	stylings.Cell(4,6).String('no');

	stylings.Cell(3,7).String('one');
	stylings.Cell(4,7).String('two');

	stylings.setValidation({
		type: "list",
		allowBlank: 1,
		sqref: "F3:F10",
		formulas: [
			'=Validations!$A$1:$A$3'
		]
	});

	stylings.setValidation({
	    type: "list",
	    allowBlank: 1,
	    showInputMessage: 1,
	    showErrorMessage: 1,
	    sqref: "G2:G10",
	    formulas: [
	        'one,two'
	    ]
	});

	//Validations Sheet
	validations.Cell(1,1).String('yes');
	validations.Cell(2,1).String('no');
	validations.Cell(3,1).String('maybe');
	return wb;
}

// Synchronously write file
createWorkBook().write("Excel.xlsx");


// Asyncrhonously write file
/*
createWorkBook().write("Excel.xlsx", function(err){
	if(err){
		console.log(err);
	}else{
		console.log('File written');
	}
});
*/


// Write file to http response
/*
var dlCount=0;

http.createServer(function(req, res){
	createWorkBook().write('MyExcel'+dlCount+'.xlsx',res);
	dlCount++;
}).listen(3000,function(){
	console.log('Download your file at http://localhost:3000');
});
*/
