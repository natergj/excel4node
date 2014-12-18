exports.Cell = function(row1, col1, row2, col2, isMerged){
	var theseCells = [];
	var thisWS = this;
	row2=row2?row2:row1;
	col2=col2?col2:col1;

	for(var r = row1; r <= row2; r++){
		var thisRow = thisWS.Row(r);
		for(var c = col1; c<= col2; c++){
			if(!thisRow.cells[c]){
				for(var i = thisRow.cells.length; i <= c; i++){
					thisRow.cells.splice(i,0,new cell());
				}
			}
			theseCells.push(thisRow.cells[c]);
		}
	}
	console.log(this);
	console.log(theseCells);
}

var cell = function(){
	this.String = string;
	this.Number = number;
	this.Date = date;
	this.Formula = formula;
	this.Format = format;
	this.Style = styler;

	function string(val){

	}
	function number(val){

	}
	function date(val){

	}
	function formula(val){

	}
	function styler(style){

	}

	var format = {
		'Number':formatter.number,
		'Font':{
			Family:formatter.font.family,
			Size:formatter.font.size,
			Bold:formatter.font.bold,
			Italics:formatter.font.italics,
			Underline:formatter.font.underline,
			Color:formatter.font.color,
			WrapText:formatter.font.wraptext,
			Alignment:{
				Vertical: formatter.font.alignment.vertical,
				Horizontal: formatter.font.alignment.horizontal
			}
		},
		Fill:{
			Color:formatter.fill.color,
			Pattern:formatter.fill.pattern
		}
	};

	return this;
}

var formatter = {
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