var cell = function(row, col, wb){
	this["@r"] = col.toExcelAlpha()+row;
	var that = this;

	this.Formula = function(val){
		this.f = val;
		delete this.v;
		return that;
	}
	this.String = function(val){
		if(wb.workbook.sharedStrings.indexOf(val) == -1){
			wb.workbook.sharedStrings.push(val);
		}
		wb.workbook.strings.sst[0]['@count']=wb.workbook.strings.sst[0]['@count']+1;
		this['@t'] = 's'
		this.v = wb.workbook.sharedStrings.indexOf(val);
		return that;
	}
	this.Style = function(style){
		this["@s"] = style.data.IDs.xf;
		return that;
	}
	this.Number = function(val){
		this.v = val;
		return that;
	}
}

exports.Row = function(row1){
	if(row1 == undefined || row1 == 0){
		console.log('invalid value for row:'+row1+'\nsetting row to 1');
	}
	var row1 = parseInt(row1) > 0 ? row1:1;
	var that = this;

	if(this.sheet.sheetData.length < row1){
		for(var i = this.sheet.sheetData.length; i < row1; i++){
			this.sheet.sheetData.push({row:[{'@r':i+1}]});
		}
	}

	var thisRow = this.sheet.sheetData[row1-1].row;

	this.attrCount = function(){
		var count = 0;
		thisRow.forEach(function(v){
			if(typeof(v[Object.keys(v)[0]]) != 'object'){
				count++;
			}
		});
		return count;
	}

	this.cellCount = function(){
		var count = 0;
		thisRow.forEach(function(v){
			if(typeof(v[Object.keys(v)[0]]) == 'object'){
				count++;
			}
		});
		return count;
	}

	this.addCell = function(col){
		if(col == undefined || col == 0){
			console.log('invalid value for column:'+col+'\nsetting column to 1');
		}
		var col = parseInt(col) > 0 ? col:1;
		if(this.cellCount() < col){
			for(var i = this.cellCount(); i <= col; i++){
				thisCell = new cell(row1, i + 1, that.wb);
				thisRow.push({c:thisCell});
			}
		}
		var rowPos = this.attrCount() + col - 1;
		return thisRow[rowPos].c
	}

	this.Height = function(ht){
		var update = false;
		thisRow.forEach(function(v){
			if(Object.keys(v)[0] == 'ht'){
				v[Object.keys(v)[0]] = ht;
				update = true;
			}
		});
		if(!update){
			thisRow.push({'@ht':ht});
			thisRow.push({'@customHeight':1});
		}
	}

	return this;
}


exports.Cell = function(row1, col1){
	if(row1 == undefined || row1 == 0){
		console.log('invalid value for row:'+row1+'\nsetting row to 1');
	}
	if(col1 == undefined || col1 == 0){
		console.log('invalid value for column:'+col1+'\nsetting column to 1');
	}
	var row1 = parseInt(row1) > 0 ? row1:1;
	var col1 = parseInt(col1) > 0 ? col1:1;
	return this.Row(row1).addCell(col1);
}
