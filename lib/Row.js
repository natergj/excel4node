var cell = function(row, col, wb){
	this["@r"] = col.toExcelAlpha()+row;
	var that = this;

	this.Formula = function(val){
		that.f = val;
		delete that.v;
		return that;
	}
	this.String = function(val){
		var val = val?val:'undefined';
		if(typeof val != 'string'){
			val = 'value passed was not a string';
		}
		if(wb.workbook.sharedStrings.indexOf(val) == -1){
			wb.workbook.sharedStrings.push(val);
		}
		wb.workbook.strings.sst[0]['@count']=wb.workbook.strings.sst[0]['@count']+1;
		that['@t'] = 's'
		that.v = wb.workbook.sharedStrings.indexOf(val);
		return that;
	}
	this.Style = function(style){
		that["@s"] = style.data.IDs.xf;
		return that;
	}
	this.Number = function(val){
		that.v = val;
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
	this.style;
	this.Style = cellStyle; 
	this.Number = cellNumber;
	this.Format = {
		cellRange:this,
		'Number':formatter.number,
		'Font':{
			cellRange:this,
			Family:formatter.font.family,
			Size:formatter.font.size,
			Bold:formatter.font.bold,
			Italics:formatter.font.italics,
			Underline:formatter.font.underline,
			Color:formatter.font.color,
			WrapText:formatter.font.wraptext,
			Alignment:{
				cellRange:that,
				Vertical: formatter.font.alignment.vertical,
				Horizontal: formatter.font.alignment.horizontal
			}
		},
		Fill:{
			cellRange:this,
			Color:formatter.fill.color,
			Pattern:formatter.fill.pattern
		}
	};
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
		if(typeof(this.style) == 'undefined'){
			this.style=style;
		}
		if(this.cells.length == 1){
			this.cells.forEach(function(curCell){
				curCell.Style(style);
			});
		}else{
			if(style.data.xf['@applyBorder'] == 1){
				var curBorderId = style.data.xf['@borderId'] + 1;
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
					that.rows.sort(function(a,b){return a - b;});
					that.cols.sort(function(a,b){return a - b;});
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
				that.cells.forEach(function(bCell){
					bCell.Style(style);
				});
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
	return this;
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
			thisRow.splice(1,0,{'@ht':ht});
			thisRow.splice(1,0,{'@customHeight':1});
		}
		return that;
	}

	this.Freeze = function(scrollTo){
		var sTo = scrollTo?scrollTo:row1;
		var sv = this.sheet.sheetViews[0].sheetView;
		var pane;
		var foundPane = false;
		sv.forEach(function(v,i){
			if(Object.keys(v).indexOf('pane') >= 0){
				pane = sv[i].pane;
				foundPane = true;
			}
		});
		if(!foundPane){
			var l = sv.push({
				pane:{
					'@activePane':'bottomLeft',
					'@state':'frozen',
					'@topLeftCell':'A'+sTo, 
					'@ySplit':row1-1
				}
			});
			pane = sv[l-1].pane;
		}else{
			var curTopLeft = pane['@topLeftCell'];
			var points = curTopLeft.toExcelRowCol();
			pane['@topLeftCell']=points.col.toExcelAlpha() + sTo;
			pane['@ySplit']=row1-1;
		}
	}

	this.Group = function(level,hidden){
		var hidden=hidden?1:0;
		var update = false;
		thisRow.forEach(function(v){
			if(Object.keys(v)[0] == 'ht'){
				v[Object.keys(v)[0]] = ht;
				update = true;
			}
		});
		if(!update){
			thisRow.splice(1,0,{'@outlineLevel':level});
			thisRow.splice(1,0,{'@hidden':hidden});
			var thisRowNum = 0;
			thisRow.forEach(function(p,i){
				if(Object.keys(p).indexOf('@r')>=0){
					thisRowNum=thisRow[i]['@r'];
				}
			});

			/*
			that.sheet.sheetData.forEach(function(r){
				var thisRowLevel = 0;
				thisRow.forEach(function(p,i){
					if(Object.keys(p).indexOf('@outlineLevel')>=0){
						thisRowLevel=thisRow[i]['@outlineLevel'];
					}
				});
				var thisRowHidden = 0;
				thisRow.forEach(function(p,i){
					if(Object.keys(p).indexOf('@hidden')>=0){
						thisRowHidden=thisRow[i]['@hidden'];
					}
				});
				switch(that.sheet.sheetPr.outlinePr['@summaryBelow']){
					case 1:
						var prevRow = [];
						if(thisRowNum > 1){
							prevRow = that.sheet.sheetData[thisRowNum-2]['row'];
						}

						var prevRowlevel = 0;
						prevRow.forEach(function(p,i){
							if(Object.keys(p).indexOf('@outlineLevel')>=0){
								prevRowlevel=prevRow[i]['@outlineLevel'];
							}
						});
						var prevRowHidden = 0;
						prevRow.forEach(function(p,i){
							if(Object.keys(p).indexOf('@hidden')>=0){
								prevRowHidden=prevRow[i]['@hidden'];
							}
						});
						if(thisRowLevel!=prevRowlevel && prevRowHidden==1){
							thisRow.splice(1,0,{'@collapsed':1});
						}
					break;
					default:
						var prevRow = [];
						if(thisRowNum > 1){
							prevRow = that.sheet.sheetData[thisRowNum-2]['row'];
						}

						var prevRowlevel = 0;
						prevRow.forEach(function(p,i){
							if(Object.keys(p).indexOf('@outlineLevel')>=0){
								prevRowlevel=prevRow[i]['@outlineLevel'];
							}
						});
						var prevRowHidden = 0;
						prevRow.forEach(function(p,i){
							if(Object.keys(p).indexOf('@hidden')>=0){
								prevRowHidden=prevRow[i]['@hidden'];
							}
						});
						if(thisRowLevel!=prevRowlevel && thisRowHidden==1){
							prevRow.splice(1,0,{'@collapsed':1});
						}
					break;
				}
			});
			*/
		}
		return that;
	}

	this.Filter = function(startCol, endCol){
		startCol=startCol?startCol:1;
		endCol=endCol?endCol:this.cellCount();
		this.sheet['autoFilter']={
			'@ref':startCol.toExcelAlpha()+row1+':'+endCol.toExcelAlpha()+row1
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

var formatter = {
	number:function(fmt){
		var that = this;
		this.cellRange.cells.forEach(function(bCell){
			if(bCell['@s']){
				var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
			}else{
				var curStyle = that.cellRange.wb.Style();
			}
			var tmpStyle = curStyle.Clone();
			tmpStyle.Number.Format(fmt);
			bCell.Style(tmpStyle);
		});
		return that;
	},
	font:{
		family:function(val){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.Family(val);
				bCell.Style(tmpStyle);
			});
			return this;
		},
		size:function(val){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.Size(val);
				bCell.Style(tmpStyle);
			});
			return this;
		},
		bold:function(){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.Bold();
				bCell.Style(tmpStyle);
			});
			return this;
		},
		italics:function(){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.Italics();
				bCell.Style(tmpStyle);
			});
			return this;
		},
		underline:function(){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.Underline();
				bCell.Style(tmpStyle);
			});
			return this;
		},
		color:function(val){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.Color(val);
				bCell.Style(tmpStyle);
			});
			return this;
		},
		wraptext:function(val){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Font.WrapText(val);
				bCell.Style(tmpStyle);
			});
			return this;
		},
		alignment:{
			vertical:function(val){
				var that = this;
				this.cellRange.cells.forEach(function(bCell){
					if(bCell['@s']){
						var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
					}else{
						var curStyle = that.cellRange.wb.Style();
					}
					var tmpStyle = curStyle.Clone();
					tmpStyle.Font.Alignment.Vertical(val);
					bCell.Style(tmpStyle);
				});
				return this;
			},
			horizontal:function(val){
				var that = this;
				this.cellRange.cells.forEach(function(bCell){
					if(bCell['@s']){
						var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
					}else{
						var curStyle = that.cellRange.wb.Style();
					}
					var tmpStyle = curStyle.Clone();
					tmpStyle.Font.Alignment.Horizontal(val);
					bCell.Style(tmpStyle);
				});
				return this;
			}
		}

	},
	fill:{
		color:function(val){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Fill.Color(val);
				bCell.Style(tmpStyle);
			});
			return that;
		},
		pattern:function(val){
			var that = this;
			this.cellRange.cells.forEach(function(bCell){
				if(bCell['@s']){
					var curStyle = that.cellRange.wb.styles['@'+bCell['@s']];
				}else{
					var curStyle = that.cellRange.wb.Style();
				}
				var tmpStyle = curStyle.Clone();
				tmpStyle.Fill.Pattern(val);
				bCell.Style(tmpStyle);
			});
			return that;
		}
	}
}

function deepClone( obj ) {
    var r,
        i = 0,
        len = obj.length;

    if ( typeof obj !== "object" ) { // string, number, boolean
        r = obj;
    }
    else if ( len  || len === 0 ) { // Simple check for array
        r = [];
        for (i=0 ; i < len; i++ ) {
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