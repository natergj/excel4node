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
		theseCells.cells.forEach(function(c,i){
			c.Style(sty.xf.xfId);
		});
		return theseCells;
	}
	function date(val){
		theseCells.cells.forEach(function(c,i){
			c.ws = thisWS;
			c.wb = thisWS.wb;
			c.Date(val);
		});
		return theseCells;
	}
	function formula(val){
		theseCells.cells.forEach(function(c,i){
			c.Formula(val);
		});
		return theseCells;
	}
	function number(val){
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
				});
			},
			font:{
				family:function(val){
					theseCells.cells.forEach(function(c){
						if(c.getAttribute('s')!=undefined){
							var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
						}else{
							var curStyle = Style.getStyleById(c.ws.wb,0);
						}
						var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
						console.log(curFont);
						curFont.name = val;

						var thisFont = new Style.font(c.ws.wb,curFont);
						var curXF = curStyle;
						curXF.applyFont=1;
						curXF.fontId=thisFont.fontId;
						var newXF = new Style.cellXfs(c.ws.wb,curXF);
						c.setAttribute('s',newXF.xfId);
					});
				},
				size:function(val){
					theseCells.cells.forEach(function(c){
						if(c.getAttribute('s')!=undefined){
							var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
						}else{
							var curStyle = Style.getStyleById(c.ws.wb,0);
						}
						var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
						curFont.size = val;
						
						var thisFont = new Style.font(c.ws.wb,curFont);
						var curXF = curStyle;
						curXF.applyFont=1;
						curXF.fontId=thisFont.fontId;
						var newXF = new Style.cellXfs(c.ws.wb,curXF);
						c.setAttribute('s',newXF.xfId);
					});
				},
				bold:function(){
					theseCells.cells.forEach(function(c){
						if(c.getAttribute('s')!=undefined){
							var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
						}else{
							var curStyle = Style.getStyleById(c.ws.wb,0);
						}
						var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
						curFont.bold = true;
						
						var thisFont = new Style.font(c.ws.wb,curFont);
						var curXF = curStyle;
						curXF.applyFont=1;
						curXF.fontId=thisFont.fontId;
						var newXF = new Style.cellXfs(c.ws.wb,curXF);
						c.setAttribute('s',newXF.xfId);
					});
				},
				italics:function(){
					theseCells.cells.forEach(function(c){
						if(c.getAttribute('s')!=undefined){
							var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
						}else{
							var curStyle = Style.getStyleById(c.ws.wb,0);
						}
						var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
						curFont.italics = true;
						
						var thisFont = new Style.font(c.ws.wb,curFont);
						var curXF = curStyle;
						curXF.applyFont=1;
						curXF.fontId=thisFont.fontId;
						var newXF = new Style.cellXfs(c.ws.wb,curXF);
						c.setAttribute('s',newXF.xfId);
					});
				},
				underline:function(){
					theseCells.cells.forEach(function(c){
						if(c.getAttribute('s')!=undefined){
							var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
						}else{
							var curStyle = Style.getStyleById(c.ws.wb,0);
						}
						var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
						curFont.underline = true;
						
						var thisFont = new Style.font(c.ws.wb,curFont);
						var curXF = curStyle;
						curXF.applyFont=1;
						curXF.fontId=thisFont.fontId;
						var newXF = new Style.cellXfs(c.ws.wb,curXF);
						c.setAttribute('s',newXF.xfId);
					});
				},
				color:function(val){
					theseCells.cells.forEach(function(c){
						if(c.getAttribute('s')!=undefined){
							var curStyle = Style.getStyleById(c.ws.wb,c.getAttribute('s'));
						}else{
							var curStyle = Style.getStyleById(c.ws.wb,0);
						}
						var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
						curFont.color = Style.cleanColor(val);
						
						var thisFont = new Style.font(c.ws.wb,curFont);
						var curXF = curStyle;
						curXF.applyFont=1;
						curXF.fontId=thisFont.fontId;
						var newXF = new Style.cellXfs(c.ws.wb,curXF);
						c.setAttribute('s',newXF.xfId);
					});
				},
				wraptext:function(val){
					theseCells.cells.forEach(function(c){
					});
				},
				alignment:{
					vertical:function(val){
						theseCells.cells.forEach(function(c){
						});
					},
					horizontal:function(val){
						theseCells.cells.forEach(function(c){
						});
					}
				}
			},
			fill:{
				color:function(val){
					theseCells.cells.forEach(function(c){
					});
				},
				pattern:function(val){
					theseCells.cells.forEach(function(c){
					});
				}
			}
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
