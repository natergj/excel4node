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

var cellRange = function(){
	var that = this;
	this.range = [];
	this.isMerged = false;
	this.rows = [];
	this.cols = [];
	this.cells = [];
	this.wb;
	this.Formula = cellFormula; 
	this.String = cellString;
	this.Style = cellStyle; 
	this.Number = cellNumber;
	this.getFirstCell = firstCellInRange;

	function cellFormula(val){
		if(!this.isMerged){
			this.cells.forEach(function(curCell){
				curCell.Formula(val);
			});
		}else{
			that.getFirstCell().Formula(val);
		}
		return that;
	}
	function cellString(val){
		if(!this.isMerged){
			this.cells.forEach(function(curCell){
				curCell.String(val);
			});
		}else{
			that.getFirstCell().String(val);
		}
		return that;
	}
	function cellStyle(style){
		if(!this.isMerged){
			this.cells.forEach(function(curCell){
				curCell.Style(style);
			});
		}else{
			if(style.data.xf['@applyBorder'] == 1){
				var curBorderId = style.data.xf['@borderId'];
				var masterBorder = {};
				that.wb.workbook.styles.styleSheet.borders[curBorderId].border.forEach(function(i){
					//convert the stored array version of the border definition to an object with ordinals as keys
					if(typeof i == 'object'){
						masterBorder[Object.keys(i)[0]] = {'style':i[Object.keys(i)[0]]['@style']};
					}
				});
				that.cells.forEach(function(bCell){
					var tmpStyle = that.wb.Style();
					tmpStyle.data = deepClone(style.data);
					var tmpBorder = {};
					that.rows.sort();
					that.cols.sort();
					var topRow = that.rows[0];
					var botRow = that.rows[that.rows.length - 1];
					var lCol = that.cols[0];
					var rCol = that.cols[that.cols.length - 1];
					var cellProps = bCell['@r'].toExcelRowCol();
					if(cellProps.row == topRow){
						tmpBorder.top = masterBorder.top;
					}
					if(cellProps.row == botRow){
						tmpBorder.bottom = masterBorder.bottom;
					}
					if(cellProps.col == lCol){
						tmpBorder.left = masterBorder.left;
					}
					if(cellProps.col == rCol){
						tmpBorder.right = masterBorder.right;
					}
					tmpStyle.Border(tmpBorder);
					bCell.Style(tmpStyle);
				});
			}else{
				that.getFirstCell().Style(style);
			}
		}
		return that;
	}
	function cellNumber(val){
		if(!this.isMerged){
			this.cells.forEach(function(curCell){
				curCell.Number(val);
			});
		}else{
			that.getFirstCell().Number(val);
		}
		return that;
	}
	function firstCellInRange(){
		this.cells.sort(function(a,b){
			if (a['@r'] < b['@r']) {
				return -1;
			}
			if (a['@r'] > b['@r']) {
				return 1;
			}
				return 0;
		});
		return this.cells[0];
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


exports.Cell = function(row1, col1, row2, col2, merged){
	var that = this;
	row2=row2?row2:row1;
	col2=col2?col2:col1;
	if(row1 == undefined || row1 == 0){
		console.log('invalid value for row:'+row1+'\nsetting row to 1');
	}
	if(col1 == undefined || col1 == 0){
		console.log('invalid value for column:'+col1+'\nsetting column to 1');
	}
	var row1 = parseInt(row1) > 0 ? row1:1;
	var col1 = parseInt(col1) > 0 ? col1:1;

	if(row2 && col2 && merged){
		if(!that.sheet['mergeCells']){
			that.sheet['mergeCells'] = [
				{
					'@count':0
				}
			];
		}
		var okToMerge = true;
		that.sheet['mergeCells'].forEach(function(c,i){
			if(c.mergeCell){
				var newRef = col1.toExcelAlpha()+row1+":"+col2.toExcelAlpha()+row2;
				var oldRef = c['mergeCell'][0]['@ref'];
				if(newRef == oldRef){
					//merged cell definition exists. no need to merge again.
					okToMerge = false;
				}
			}
		});
		if(okToMerge){
			merge(row1,col1,row2,col2);
		}
	}

	function merge(row1, col1, row2, col2){
		var mergeRangeValid = true;
		that.sheet['mergeCells'].forEach(function(c,i){
			if(c.mergeCell){
				var curCells = getAllCellsInExcelRange(c['mergeCell'][0]['@ref']);
				var newCells = getAllCellsInNumericRange(row1,col1,row2,col2);
				var intersection = arrayIntersectSafe(curCells,newCells);
				if(intersection.length > 0){
					mergeRangeValid = false;
					console.log([
									'Invalid Range for : '+col1.toExcelAlpha()+row1+":"+col2.toExcelAlpha()+row2,
									'Some cells in this range are already included in another merged cell range: '+c['mergeCell'][0]['@ref'],
									'The following are the intersection',
									intersection
								]);
				}
			}
		});

		if(mergeRangeValid){
			that.sheet['mergeCells'][0]['@count'] += parseInt(that.sheet['mergeCells'][0]['@count']) + 1;
			that.sheet['mergeCells'].push(			
				{
					mergeCell:[
						{
							'@ref':col1.toExcelAlpha()+row1+":"+col2.toExcelAlpha()+row2
						}
					]
				}
			);
		}
	}
	var response = new cellRange();
	response.range = getAllCellsInNumericRange(row1,col1,row2,col2);
	response.isMerged = merged?merged:false;
	response.wb = this.wb;
	for(i=row1; i<=row2; i++){
		if(response.rows.indexOf(i)<0){
			response.rows.push(i);
		}
		for(j=col1; j<=col2; j++){
			if(response.cols.indexOf(j)<0){
				response.cols.push(j);
			}
			response.cells.push(this.Row(i).addCell(j));
		}
	}
	//console.log(response);
	return response;
	//return this.Row(row1).addCell(col1);
}

function getAllCellsInNumericRange(row1, col1, row2, col2){
	var response = [];
	row2=row2?row2:row1;
	col2=col2?col2:col1;
	for(i=row1; i<=row2; i++){
		for(j=col1; j<=col2; j++){
			response.push(j.toExcelAlpha() + i);
		}
	}
	return response.sort();
}
function getAllCellsInExcelRange(range){
	var cells = range.split(':');
	var cell1props = cells[0].toExcelRowCol();
	var cell2props = cells[1].toExcelRowCol();
	return getAllCellsInNumericRange(cell1props.row,cell1props.col,cell2props.row,cell2props.col);
}
function arrayIntersectSafe(a, b){
  var ai=0, bi=0;
  var result = new Array();

  while( ai < a.length && bi < b.length )
  {
     if      (a[ai] < b[bi] ){ ai++; }
     else if (a[ai] > b[bi] ){ bi++; }
     else /* they're equal */
     {
       result.push(a[ai]);
       ai++;
       bi++;
     }
  }

  return result;
}

function deepClone( obj ) {
    var r,
        i = 0,
        len = obj.length;

    if ( typeof obj !== "object" ) { // string, number, boolean
        r = obj;
    }
    else if ( len ) { // Simple check for array
        r = [];
        for ( ; i < len; i++ ) {
            r.push( deepClone(obj[i]) );
        }
    } 
    else if ( obj.getTime ) { // Simple check for date
        r = new Date( +obj );
    }
    else if ( obj.nodeName ) { // Simple check for DOM node
        r = obj;
    }
    else { // Object
        r = {};
        for ( i in obj ) {
            r[i] = deepClone(obj[i]);
        }
    }

    return r;
}