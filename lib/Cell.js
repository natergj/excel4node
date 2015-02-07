var Style = require(__dirname+'/Style.js'),
	_ = require('underscore');

exports.Cell = function(row1, col1, row2, col2, isMerged){

	var thisWS = this;
	var theseCells = {
		cells:[],
		excelRefs:[]
	};

	/******************************
		Cell Range Methods
	 ******************************/
	theseCells.String = string;
	theseCells.Number = number;
	theseCells.Format = format();
	theseCells.Style = style;
	theseCells.Date = date;
	theseCells.Formula = formula;
	theseCells.Merge = mergeCells;
	theseCells.Format = format();
	theseCells.getCell = getCell;

	row2=row2?row2:row1;
	col2=col2?col2:col1;

	/******************************
		Add all cells in range to Cell definition
	 ******************************/
	for(var r = row1; r <= row2; r++){
		var thisRow = thisWS.Row(r);
		for(var c = col1; c<= col2; c++){
			var thisCol = thisWS.Column(parseInt(c));
			if(!thisRow.cells[c]){
				thisRow.cells[c] = new cell(thisWS);
			}
			var thisCell = thisRow.cells[c];
			thisCell.attributes['r'] = c.toExcelAlpha() + r;
			thisRow.attributes['spans'] = '1:'+thisRow.cellCount();
			theseCells.cells.push(thisCell);
		}
	}
	theseCells.excelRefs = getAllCellsInNumericRange(row1,col1,row2,col2);
	theseCells.cells.sort(excelCellsSort);

	if(isMerged){
		theseCells.Merge();
	}
	/******************************
		Cell Range Method Definitions
	 ******************************/
	function string(val){
		if(typeof(val) != 'string'){
			console.log('Value sent to String function of cells %s was not a string, it has type of %s',JSON.stringify(theseCells.excelRefs),typeof(val));
			val = '';
		}
		val=val.toString();

		theseCells.cells.forEach(function(c,i){
			c.String(thisWS.wb.getStringIndex(val));
		});
		return theseCells;
	}
	function format(){
		var methods = {
			'Number':formatter().number,
			'Date':formatter().date,
			'Font':{
				Family:formatter().font.family,
				Size:formatter().font.size,
				Bold:formatter().font.bold,
				Italics:formatter().font.italics,
				Underline:formatter().font.underline,
				Color:formatter().font.color,
				WrapText:formatter().font.wraptext,
				Alignment:{
					Vertical: formatter().font.alignment.vertical,
					Horizontal: formatter().font.alignment.horizontal
				}
			},
			Fill:{
				Color:formatter().fill.color,
				Pattern:formatter().fill.pattern
			}
		};
		return methods;
	}
	function style(sty){

		// If style has a border, split excel cells into rows
		if(sty.xf.applyBorder > 0){
			var cellRows = [];
			var curRow = [];
			var curCol = "";
			theseCells.excelRefs.forEach(function(cr,i){
				var thisCol = cr.replace(/[0-9]/g,'');
				if(thisCol!=curCol){
					if(curRow.length > 0){
						cellRows.push(curRow);
						curRow=[];
					}
					curCol=thisCol;
				}
				curRow.push(cr);
				if(i == theseCells.excelRefs.length-1 && curRow.length > 0){
					cellRows.push(curRow);
				}
			});

			var borderEdges = {}
			borderEdges.left = cellRows[0][0].replace(/[0-9]/g,'');
			borderEdges.right = cellRows[cellRows.length-1][0].replace(/[0-9]/g,'');
			borderEdges.top = cellRows[0][0].replace(/[a-zA-Z]/g,'');
			borderEdges.bottom = cellRows[0][cellRows[0].length-1].replace(/[a-zA-Z]/g,'');
		}

		theseCells.cells.forEach(function(c,i){
			if(theseCells.excelRefs.length == 1 || sty.xf.applyBorder == 0){
				c.Style(sty.xf.xfId);
			}else{
				var curBorderId = sty.xf.borderId;
				var masterBorder = JSON.parse(JSON.stringify(c.ws.wb.styleData.borders[curBorderId]));

				var thisBorder = {};
				var cellRef = c.getAttribute('r');
				var cellCol = cellRef.replace(/[0-9]/g,'');
				var cellRow = cellRef.replace(/[a-zA-Z]/g,'');

				if(cellRow == borderEdges.top){
					thisBorder.top = masterBorder.top;
				}
				if(cellRow == borderEdges.bottom){
					thisBorder.bottom = masterBorder.bottom;
				}
				if(cellCol == borderEdges.left){
					thisBorder.left = masterBorder.left;
				}
				if(cellCol == borderEdges.right){
					thisBorder.right = masterBorder.right;
				}

				if(c.getAttribute('s')!=undefined){
					var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
				}else{
					var curStyle = Style.getStyleById(c.ws.wb,sty.xf.xfId);
				}

				var newBorder = new Style.border(c.ws.wb,thisBorder);
				var curXF = JSON.parse(JSON.stringify(curStyle));
				curXF.applyBorder=1;
				curXF.borderId=newBorder.borderId;
				var newXF = new Style.cellXfs(c.ws.wb,curXF);
				c.setAttribute('s',newXF.xfId);
			}
		});
		return theseCells;
	}
	function date(val){
		if(!val || !val.toISOString || val.toISOString() != new Date(val).toISOString()){
			val = new Date();
			console.log('Value sent to Date function of cells %s was not a date, it has type of %s',JSON.stringify(theseCells.excelRefs),typeof(val));
		}
		theseCells.cells.forEach(function(c,i){
			var styleInfo = c.getStyleInfo();
			if(styleInfo.applyNumberFormat==1){
				if(styleInfo.numFmt.formatCode.substr(0,7) != '[$-409]'){
					console.log('Number format was already set for cell %s. It will be overridden with date format',thisCell.getAttribute('r'));
					c.toCellRange().Format.Date();
				}
			}else{
				c.toCellRange().Format.Date();
			}
			c.Date(val);
		});
		return theseCells;
	}
	function formula(val){
		if(typeof(val) != 'string'){
			console.log('Value sent to Formula function of cells %s was not a string, it has type of %s',JSON.stringify(theseCells.excelRefs),typeof(val));
			val = '';
		}
		theseCells.cells.forEach(function(c,i){
			c.Formula(val);
		});

		return theseCells;
	}
	function number(val){
		if(val == undefined || parseFloat(val) != val){
			console.log('Value sent to Number function of cells %s was not a number, it has type of %s%s',
				JSON.stringify(theseCells.excelRefs),
				typeof(val),
				typeof(val)=='string'?' and value of "'+val+'"':''
			);
			val = '';
		}
		val=parseFloat(val);

		theseCells.cells.forEach(function(c,i){
			c.Number(val);
			if(c.getAttribute('t')){
				c.deleteAttribute('t');
			}
		});
		return theseCells;
	}
	function mergeCells(){
		if(!thisWS.mergeCells){
			thisWS.mergeCells=[];
		};
		var cellRange = this.excelRefs[0]+':'+this.excelRefs[this.excelRefs.length-1];
		var rangeCells = this.excelRefs;
		var okToMerge = true;
		thisWS.mergeCells.forEach(function(cr){
			// Check to see if currently merged cells contain cells in new merge request
			var curCells = getAllCellsInExcelRange(cr);
			var intersection = arrayIntersectSafe(rangeCells,curCells);
			if(intersection.length > 0){
				okToMerge = false;
				console.log([
					'Invalid Range for : '+col1.toExcelAlpha()+row1+":"+col2.toExcelAlpha()+row2,
					'Some cells in this range are already included in another merged cell range: '+c['mergeCell'][0]['@ref'],
					'The following are the intersection',
					intersection
				]);
			}
		});
		if(okToMerge){
			thisWS.mergeCells.push(cellRange);
		}
	}
	function formatter(){

		var methods = {
			number:function(fmt){
				theseCells.cells.forEach(function(c){
					setNumberFormat(c, 'formatCode', fmt);
				});
				return theseCells;
			},
			date:function(fmt){
				theseCells.cells.forEach(function(c){
					setDateFormat(c, 'formatCode', fmt);
				});
				return theseCells;
			},
			font:{
				family:function(val){
					theseCells.cells.forEach(function(c){
						setCellFontAttribute(c,'name',val);
					});
					return theseCells;
				},
				size:function(val){
					theseCells.cells.forEach(function(c){
						setCellFontAttribute(c,'size',val);
					});
					return theseCells;
				},
				bold:function(){
					theseCells.cells.forEach(function(c){
						setCellFontAttribute(c,'bold',true);
					});
					return theseCells;
				},
				italics:function(){
					theseCells.cells.forEach(function(c){
						setCellFontAttribute(c,'italics',true);
					});
					return theseCells;
				},
				underline:function(){
					theseCells.cells.forEach(function(c){
						setCellFontAttribute(c,'underline',true);
					});
					return theseCells;
				},
				color:function(val){
					theseCells.cells.forEach(function(c){
						setCellFontAttribute(c,'color',Style.cleanColor(val));
					});
					return theseCells;
				},
				wraptext:function(val){
					theseCells.cells.forEach(function(c){
						setAlignmentAttribute(c,'wrapText',1);
					});
					return theseCells;
				},
				alignment:{
					vertical:function(val){
						theseCells.cells.forEach(function(c){
							setAlignmentAttribute(c,'vertical',val);
						});
						return theseCells;
					},
					horizontal:function(val){
						theseCells.cells.forEach(function(c){
							setAlignmentAttribute(c,'horizontal',val);
						});
						return theseCells;
					}
				}
			},
			fill:{
				color:function(val){
					theseCells.cells.forEach(function(c){
						setCellFill(c,'fgColor',Style.cleanColor(val));
					});
					return theseCells;
				},
				pattern:function(val){
					theseCells.cells.forEach(function(c){
						setCellFill(c,'patternType',val);
					});
					return theseCells;
				}
			}
		}

		function setAlignmentAttribute(c,attr,val){
			if(c.getAttribute('s')!=undefined){
				var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
			}else{
				var curStyle = Style.getStyleById(c.ws.wb,0);
			}

			var curXF = JSON.parse(JSON.stringify(curStyle));
			if(!curXF.alignment){
				curXF.alignment = {};
			}
			curXF.applyAlignment=1;
			curXF.alignment[attr]=val;

			var newXF = new Style.cellXfs(c.ws.wb,curXF);
			c.setAttribute('s',newXF.xfId);
		}

		function setNumberFormat(c, attr, val){
			if(c.getAttribute('s')!=undefined){
				var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
			}else{
				var curStyle = Style.getStyleById(c.ws.wb,0);
			}

			if(curStyle.numFmtId != 164 && curStyle.numFmtId != 14){
				var curNumFmt = JSON.parse(JSON.stringify(c.ws.wb.styleData.numFmts[curStyle.numFmtId - 165]));
			}else{
				var curNumFmt = {};
			}

			curNumFmt[attr] = val;
			var thisFmt = new Style.numFmt(c.ws.wb,curNumFmt);
			var curXF = JSON.parse(JSON.stringify(curStyle));
			curXF.applyNumberFormat=1;
			curXF.numFmtId=thisFmt.numFmtId;
			var newXF = new Style.cellXfs(c.ws.wb,curXF);
			c.setAttribute('s',newXF.xfId);
		}

		function setDateFormat(c, attr, val){
			if(c.getAttribute('s')!=undefined){
				var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
			}else{
				var curStyle = Style.getStyleById(c.ws.wb,0);
			}

			if(curStyle.numFmtId != 164 && curStyle.numFmtId != 14 && c.ws.wb.styleData.numFmts[curStyle.numFmtId - 165].formatCode.substr(0,7) == '[$-409]'){
				var curNumFmt = JSON.parse(JSON.stringify(c.ws.wb.styleData.numFmts[curStyle.numFmtId - 165]));
			}else{
				var curNumFmt = {};
			}

			var curXF = JSON.parse(JSON.stringify(curStyle));
			curXF.applyNumberFormat=1;
			if(val){
				curNumFmt[attr] = val;
				var thisFmt = new Style.numFmt(c.ws.wb,curNumFmt);
				curXF.numFmtId=thisFmt.numFmtId;
			}else{
				curXF.numFmtId=14;
			}
			var newXF = new Style.cellXfs(c.ws.wb,curXF);
			c.setAttribute('s',newXF.xfId);
		}

		function setCellFill(c, attr, val){
			if(c.getAttribute('s')!=undefined){
				var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
			}else{
				var curStyle = Style.getStyleById(c.ws.wb,0);
			}

			var curFill = JSON.parse(JSON.stringify(c.ws.wb.styleData.fills[curStyle.fillId]));

			curFill[attr] = val;
			var thisFill = new Style.fill(c.ws.wb,curFill);
			var curXF = JSON.parse(JSON.stringify(curStyle));
			curXF.applyFill=1;
			curXF.fillId=thisFill.fillId;
			var newXF = new Style.cellXfs(c.ws.wb,curXF);
			c.setAttribute('s',newXF.xfId);;
		}

		function setCellFontAttribute(c,attr,val){
			if(c.getAttribute('s')!=undefined){
				var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
			}else{
				var curStyle = Style.getStyleById(c.ws.wb,0);
			}
			var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
			curFont[attr] = val;

			var thisFont = new Style.font(c.ws.wb,curFont);
			var curXF = JSON.parse(JSON.stringify(curStyle));
			curXF.applyFont=1;
			curXF.fontId=thisFont.fontId;
			var newXF = new Style.cellXfs(c.ws.wb,curXF);
			c.setAttribute('s',newXF.xfId);
		}

		return methods;
	}
	function getCell(ref){
		return theseCells.cells[theseCells.excelRefs.indexOf(ref)];
	}
	return theseCells;
}

var cell = function(ws){
	var thisCell = this;
	thisCell.ws = ws;

	/******************************
		Cell Definition
	 ******************************/
	thisCell.attributes = {};
	thisCell.children = {};

	return thisCell;
}

/******************************
	Cell Methods
 ******************************/
cell.prototype.setAttribute = setAttribute;
cell.prototype.getAttribute = getAttribute;
cell.prototype.deleteAttribute = deleteAttribute;
cell.prototype.getStyleInfo = getStyleInfo;
cell.prototype.toCellRange = getCellRange;
cell.prototype.addChild = addChild;
cell.prototype.String = string;
cell.prototype.Number = number;
cell.prototype.Date = date;
cell.prototype.Formula = formula;
cell.prototype.Style = styler;	


/******************************
	Cell Method Definitions
 ******************************/
function addChild(key,val){

	this.children[key]=val;
}

function setAttribute(attr,val){

	this.attributes[attr] = val;
}

function getAttribute(attr){

	return this.attributes[attr];
}

function deleteAttribute(attr){

	return delete this.attributes[attr];
}

function getStyleInfo(){

	var styleData = {};
	if(this.getAttribute('s')){
		//UNDO var wbStyleData = this.ws.wb.styleData;
		var wbStyleData = ws.wb.styleData;
		var xf = wbStyleData.cellXfs[this.Attribute('s')];
		styleData.xf = xf;
		styleData.applyAlignment=xf.applyAlignment;
		styleData.applyBorder=xf.applyBorder;
		styleData.applyNumberFormat=xf.applyNumberFormat;
		styleData.applyFill=xf.applyFill;
		styleData.applyFont=xf.applyFont;
		if(xf.applyAlignment!=0){
			styleData.alignment=xf.alignment;
		}
		if(xf.applyBorder != 0){
			styleData.border=wbStyleData.borders[xf.borderId];
		}
		if(xf.applyNumberFormat != 0){
			styleData.numFmt=wbStyleData.numFmts[xf.numFmtId - 165];
		}
		if(xf.applyFill != 0){
			styleData.fill=wbStyleData.fills[xf.fillId];
		}
		if(xf.applyFont != 0){
			styleData.font=wbStyleData.fonts[xf.fontId];
		}
	}
	return styleData;
}

function getCellRange(){

	//Since all formatting is done on cell ranges, convert cell to range of single cell
	var rc = this.getAttribute('r').toExcelRowCol();
	//UNDO return this.ws.Cell(rc.row,rc.col);
	return this.ws.Cell(rc.row,rc.col);
}

function string(strIndex){

	this.setAttribute('t','s');
	this.addChild('v',strIndex);
}

function number(val){

	this.addChild('v',val);
}

function date(val){

	val = new Date(val);
	var ts = val.getExcelTS();
	this.addChild('v',ts);
}

function formula(val){

	this.addChild('f',val);
}

function styler(style){

	this.setAttribute('s',style);
}

function getAllCellsInNumericRange(row1, col1, row2, col2){

	var response = [];
	row2=row2?row2:row1;
	col2=col2?col2:col1;
	for(var i=row1; i<=row2; i++){
		for(var j=col1; j<=col2; j++){
			response.push(j.toExcelAlpha() + i);
		}
	}
	return response.sort(excelRefSort);
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

function excelRefSort(a, b){
  if(a.replace(/[0-9]/g,'') === b.replace(/[0-9]/g,'')){
    return a.replace(/[a-zA-Z]/g,'') - b.replace(/[a-zA-Z]/g,'');
  }

  return compareCharCodes(a, b);
}

function excelCellsSort(a, b){
  var ar = a.attributes.r;
  var br = b.attributes.r;
  if(ar.replace(/[0-9]/g,'') === br.replace(/[0-9]/g,'')){
    return ar.replace(/[a-zA-Z]/g,'') - br.replace(/[a-zA-Z]/g,'');
  }

  return compareCharCodes(ar, br);
}

function compareCharCodes(a, b){
  var alphaOne = a.replace(/[0-9]/g,'').toUpperCase();
  var alphaTwo = b.replace(/[0-9]/g,'').toUpperCase();
  var numOne = '';
  var numTwo = '';

  for( i = 0; i < alphaOne.length; i++){
    numOne += alphaOne.charCodeAt(i);
  }
  for( i = 0; i < alphaTwo.length; i++){
    numTwo += alphaTwo.charCodeAt(i);
  }

  return Number(numOne) - Number(numTwo);
}
