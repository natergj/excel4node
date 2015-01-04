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
			thisRow.cells[c].attributes['r'] = c.toExcelAlpha() + r;
			thisRow.attributes['spans'] = '1:'+thisRow.cellCount();
			theseCells.cells.push(thisRow.cells[c]);
		}
	}
	//console.log(this);
	//console.log(theseCells);

	function string(val){
		theseCells.cells.forEach(function(c,i){
			c.String(thisWS.wb.getStringIndex(val));
			//console.log(i);
			//console.log(c);
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
		return theseCells;
	}
	function formula(val){
		return theseCells;
	}
	function number(val){
		return theseCells;
	}

	return theseCells;
}

var cell = function(){
	thisCell = this;
	thisCell.attributes = {};
	thisCell.children = {};
	thisCell.String = string;
	thisCell.Number = number;
	thisCell.Date = date;
	thisCell.Formula = formula;
	thisCell.Format = format();
	thisCell.Style = styler;

	function string(strIndex){
		thisCell.children = {
			'v':strIndex
		}
		thisCell.attributes.t='s';
	}
	function number(val){

	}
	function date(val){

	}
	function formula(val){

	}
	function styler(style){

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