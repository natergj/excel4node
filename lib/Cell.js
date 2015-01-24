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
			theseCells.excelRefs.sort();
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

			// Sort cells by column number so that A11 is not less than A2
			cellRows.forEach(function(r){
				r.sort(function(a,b){return a.replace(/[a-zA-Z]/g,'') - b.replace(/[a-zA-Z]/g,'')});
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
					var curStyle = Style.getStyleById(c.ws.wb,0);
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
		if(!val){
			val = new Date();
			console.log('Value sent to Date function of cells %s was not a date, it has type of %s',JSON.stringify(theseCells.excelRefs),typeof(val));
		}
		theseCells.cells.forEach(function(c,i){
			c.ws = thisWS;
			c.wb = thisWS.wb;
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
		if(!val || parseFloat(val) != val){
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

			var newXF = JSON.parse(JSON.stringify(curStyle));
			c.setAttribute('s',newXF.xfId);
		}

		function setNumberFormat(c, attr, val){
			if(c.getAttribute('s')!=undefined){
				var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
			}else{
				var curStyle = Style.getStyleById(c.ws.wb,0);
			}
			var curNumFmt = JSON.parse(JSON.stringify(c.ws.wb.styleData.numFmts[curStyle.numFmtId - 164]));

			curNumFmt[attr] = val;
			var thisFmt = new Style.numFmt(c.ws.wb,curNumFmt);
			var curXF = JSON.parse(JSON.stringify(curStyle));
			curXF.applyNumberFormat=1;
			curXF.numFmtId=thisFmt.numFmtId;
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


	/******************************
		Cell Methods
	 ******************************/
	thisCell.setAttribute = setAttribute;
	thisCell.getAttribute = getAttribute;
	thisCell.addChild = addChild;
	thisCell.String = string;
	thisCell.Number = number;
	thisCell.Date = date;
	thisCell.Formula = formula;
	thisCell.Style = styler;	


	/******************************
		Cell Method Definitions
	 ******************************/
	function addChild(key,val){
		thisCell.children[key]=val;
	}
	function setAttribute(attr,val){
		thisCell.attributes[attr] = val;
	}
	function getAttribute(attr){
		return thisCell.attributes[attr];
	}
	function string(strIndex){
		thisCell.setAttribute('t','s');
		thisCell.addChild('v',strIndex);
	}
	function number(val){
		thisCell.addChild('v',val);
	}
	function date(val){
		var ts = val.getExcelTS();
		thisCell.addChild('v',ts);
	}
	function formula(val){
		thisCell.addChild('f',val);
	}
	function styler(style){
		thisCell.setAttribute('s',style);
	}
	return thisCell;
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
