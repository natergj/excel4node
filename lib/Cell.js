exports.Cell = function(row1, col1, row2, col2, isMerged){

	var thisWS = this;
	var theseCells = {
		cells:[]
	};
	theseCells.String = string;
	theseCells.Number = number;
	theseCells.Format = format();
	theseCells.Style = style;
	theseCells.Date = date;
	theseCells.Formula = formula;
	theseCells.Format = format();

	row2=row2?row2:row1;
	col2=col2?col2:col1;

	for(var r = row1; r <= row2; r++){
		var thisRow = thisWS.Row(r);
		for(var c = col1; c<= col2; c++){
			if(!thisRow.cells[c]){
				thisRow.cells[c] = new cell();
			}
			var thisCell = thisRow.cells[c];
			thisCell.attributes['r'] = c.toExcelAlpha() + r;
			thisRow.attributes['spans'] = '1:'+thisRow.cellCount();
			theseCells.cells.push(thisCell);
		}
	}

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

	function style(val){
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

	return theseCells;
}

var cell = function(){
	thisCell = this;
	thisCell.attributes = {};
	thisCell.children = {};
	thisCell.setAttribute = setAttribute;
	thisCell.addChild = addChild;
	thisCell.String = string;
	thisCell.Number = number;
	thisCell.Date = date;
	thisCell.Formula = formula;
	thisCell.Format = format();
	thisCell.Style = styler;	

	function addChild(key,val){
		thisCell.children[key]=val;
	}
	function setAttribute(attr,val){
		thisCell.attributes[attr] = val;
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

	return this;
}

function formatter(){
	var methods = {
		number:function(fmt){
		},
		font:{
			family:function(val){
			},
			size:function(val){
			},
			bold:function(){
			},
			italics:function(){
			},
			underline:function(){
			},
			color:function(val){
			},
			wraptext:function(val){
			},
			alignment:{
				vertical:function(val){
				},
				horizontal:function(val){
				}
			}
		},
		fill:{
			color:function(val){
			},
			pattern:function(val){
			}
		}
	}
	return methods;
}