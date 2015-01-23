
// Check if sample is running from downloaded module or elsewhere.
try {
    var xl = require('excel4node');
} catch(e) {
    var xl = require('./lib/index.js');
}

var benchmarks = [noStyle, definedStyle, variableStyle];

runTests();

function runTests(){
	var curTest = benchmarks.shift();
	if(curTest){
		console.log('\n\n### Starting test');
		curTest();
	}else{
		process.exit();
	}
}

function noStyle(){
	var startTime = process.hrtime();
	var wb = new xl.WorkBook();
	var ws = wb.WorkSheet('Giant Sheet');

	for(var r = 1; r <= 5000; r++){
		for (var c = 1; c <= 20; c++){
			ws.Cell(r,c).String('String'+(Math.random()*100).toPrecision(2));
		}
	}

	var diff0 = process.hrtime(startTime);
	console.log('write started after %d nanoseconds', diff0[0] * 1e9 + diff0[1]);
	wb.write('noStyle.xlsx',function(){
		var diff1 = process.hrtime(startTime);
		console.log('no style benchmark took %d nanoseconds', diff1[0] * 1e9 + diff1[1]);
		console.log('Memory Usage: %s',JSON.stringify(process.memoryUsage()));
		runTests();
	});
}

function definedStyle(){
	var startTime = process.hrtime();
	var wb = new xl.WorkBook();
	var ws = wb.WorkSheet('Giant Sheet');

	var myStyle0 = wb.Style();
	myStyle0.Font.Family('Tahoma');
	myStyle0.Font.Color('gray-40');

	var myStyle1 = wb.Style();
	myStyle1.Fill.Color('light blue');
	myStyle1.Fill.Pattern('solid');

	var myStyle2 = wb.Style();
	myStyle2.Font.Underline();
	myStyle2.Font.Bold();

	var myStyle3 = wb.Style();
	myStyle3.Font.Color('blue');
	myStyle3.Fill.Color('red');
	myStyle3.Fill.Pattern('solid');

	var myStyle4 = wb.Style();
	myStyle4.Border({left:{style:'thin'},right:{style:'thin'},top:{style:'thin'},bottom:{style:'thin'}});

	for(var r = 1; r <= 5000; r++){
		for (var c = 1; c <= 20; c++){
			var styles = [myStyle0,myStyle1,myStyle2,myStyle3,myStyle4];
			var index = c%5;
			ws.Cell(r,c).String('String'+(Math.random()*100).toPrecision(2)).Style(styles[index]);
		}
	}

	var diff0 = process.hrtime(startTime);
	console.log('write started after %d nanoseconds', diff0[0] * 1e9 + diff0[1]);
	wb.write('noStyle.xlsx',function(){
		var diff1 = process.hrtime(startTime);
		console.log('defined style benchmark took %d nanoseconds', diff1[0] * 1e9 + diff1[1]);
		console.log('Memory Usage: %s',JSON.stringify(process.memoryUsage()));
		runTests();
	});
}

function variableStyle(){
	var startTime = process.hrtime();
	var wb = new xl.WorkBook();
	var ws = wb.WorkSheet('Giant Sheet');

	var myStyle = wb.Style();
	myStyle.Font.Family('Tahoma');
	myStyle.Font.Color('gray-40');

	for(var r = 1; r <= 5000; r++){
		for (var c = 1; c <= 20; c++){
			var colors = ['green','white','black','blue','red'];
			var index = c%5;
			ws.Cell(r,c).String('String'+(Math.random()*100).toPrecision(2));
			ws.Cell(r,c).Format.Font.Color(colors[index]);
		}
	}

	var diff0 = process.hrtime(startTime);
	console.log('write started after %d nanoseconds', diff0[0] * 1e9 + diff0[1]);
	wb.write('noStyle.xlsx',function(){
		var diff1 = process.hrtime(startTime);
		console.log('defined style benchmark took %d nanoseconds', diff1[0] * 1e9 + diff1[1]);
		console.log('Memory Usage: %s',JSON.stringify(process.memoryUsage()));
		runTests();
	});
}

function generateRId(){
    var text = "R";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < 16; i++ )
        text += possible.charAt(Math.floor(Math.random() * possible.length));

    return text;
}