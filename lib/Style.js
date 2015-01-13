var _ = require('underscore');

var style = function(wb,opts){
	this.wb = wb;
	this.xf = new cellXfs(wb);

	this.Font = fontFunctions(this);
	this.Border = setBorder;
	this.Number = numberFunctions();
	this.Fill = fillFunctions();

	function numberFunctions(){
		this.Format = setFormat;

		function setFormat(fmt){

		}

		return this;
	}

	function fillFunctions(){
		this.Color = setFillColor;
		this.Pattern = setFillPattern;

		function setFillColor(color){

		}
		function setFillPattern(pattern){

		}

		return this;
	}

	function fontFunctions(curStyle){
		this.Options = setFontOptions;
		this.Family = setFontFamily;
		this.Bold = setFontBold;
		this.Italics = setFontItalics;
		this.Underline = setFontUnderline;
		this.Size = setFontSize;
		this.Color = setFontColor;
		this.WrapText = setFontWrap;
		this.Alignment = {
			Vertical:setFontAlignmentVertical,
			Horizontal:setFontAlignmentHorizontal
		}

		var curFont = JSON.parse(JSON.stringify(wb.styleData.fonts[curStyle.xf.fontId]));

		function setFontOptions(opts){
			Object.keys(opts).forEach(function(o){
				curFont[o]=opts[o];
			});
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontFamily(val){
			curFont.name = val;
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontBold(){
			curFont.bold = true;
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontItalics(){
			curFont.italics = true;
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontUnderline(){
			curFont.underline = true;
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontSize(val){
			curFont.sz = val;
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontColor(val){
			curFont.color = cleanColor(val);
			var thisFont = new font(curStyle.wb,curFont);
			curStyle.xf = new cellXfs(curStyle.wb, {applyFont:1,fontId:thisFont.fontId});
		}
		function setFontWrap(){

		}
		function setFontAlignmentVertical(){

		}
		function setFontAlignmentHorizontal(){

		}
		return this;
	}

	function setBorder(){

	}

	return this;
}

var cellXfs = function(wb,opts){
	opts=opts?opts:{};
	this.applyNumberFormat=opts.applyNumberFormat?opts.applyNumberFormat:0;
	this.applyFill=opts.applyFill?opts.applyFill:0;
	this.applyFont=opts.applyFont?opts.applyFont:0;
	this.borderId=opts.borderId?opts.borderId:0;
	this.fillId=opts.fillId?opts.fillId:0;
	this.fontId=opts.fontId?opts.fontId:0;
	this.numFmtId=opts.numFmtId?opts.numFmtId:164;
	this.generateXMLObj=genXMLObj;

	if(wb.styleData.cellXfs.length == 0){
		this.xfId = 0;
		wb.styleData.cellXfs.push(this);
	}else{
		var isMatched = false;
		var curXfId = 0;
		var xf2 = JSON.parse(JSON.stringify(this));

		while( isMatched == false && curXfId < wb.styleData.cellXfs.length ){
			var xf1 = JSON.parse(JSON.stringify(wb.styleData.cellXfs[curXfId]));

			xf1.xfId=null;
			xf2.xfId=null;
			isMatched = _.isEqual(xf1,xf2);
			if(isMatched){
				this.xfId = curXfId;
			}else{
				curXfId+=1;
			}
		}
		if(!isMatched){
			this.xfId = wb.styleData.cellXfs.length;
			wb.styleData.cellXfs.push(this);
		}
	}

	function genXMLObj(){
		var data = {xf:{
			'@applyNumberFormat':this.applyNumberFormat,
			'@applyFill':this.applyFill,
			'@applyFont':this.applyFont,
			'@borderId':this.borderId,
			'@fillId':this.fillId,
			'@fontId':this.fontId,
			'@numFmtId':this.numFmtId
		}};
		return data;
	}

	return this;
}

var font = function(wb,opts){
	opts=opts?opts:{};
	this.bold=opts.bold?opts.bold:false;
	this.italics=opts.italics?opts.italics:false;
	this.underline=opts.underline?opts.underline:false;
	this.sz=opts.size?opts.size:12;
	this.color=opts.color?cleanColor(opts.color):'FF000000';
	this.name=opts.name?opts.name:'Calibri';
	//this.scheme=opts.scheme?opts.scheme:'minor';
	this.generateXMLObj=genXMLObj;

	if(wb.styleData.fonts.length == 0){
		this.fontId = 0;
		wb.styleData.fonts.push(this);
	}else{
		var isMatched = false;
		var curFontId = 0;
		var font2 = JSON.parse(JSON.stringify(this));

		while( isMatched == false && curFontId < wb.styleData.fonts.length ){
			var font1 = JSON.parse(JSON.stringify(wb.styleData.fonts[curFontId]));

			font1.fontId=null;
			font2.fontId=null;
			isMatched = _.isEqual(font1,font2);
			if(isMatched){
				this.fontId = curFontId;
			}else{
				curFontId+=1;
			}
		}
		if(!isMatched){
			this.fontId = wb.styleData.fonts.length;
			wb.styleData.fonts.push(this);
		}
	}

	function genXMLObj(){
		var data = {
						font:[
							{
								sz:{
									'@val':this.sz
								}
							},
							{
								color:{
									'@rgb':this.color
								}
							},
							{
								name:{
									'@val':this.name
								}
							}
						]
					};
		if(this.underline){
			data.font.splice(0,0,'u');
		};
		if(this.italics){
			data.font.splice(0,0,'i');
		};
		if(this.bold){
			data.font.splice(0,0,'b');
		};
		return data;
	}

	return this;
}

var numFmt = function(wb, opts){
}

var fill = function(wb,opts){
	opts=opts?opts:{};
	this.patternType=opts.patternType?opts.patterType:'solid';
	this.fgColor=opts.fgColor?cleanColor(opts.fgColor):'FFFFFFFF';
	this.bgColor=opts.bgColor?cleanColor(opts.bgColor):'FFFFFFFF';
	this.generateXMLObj=genXMLObj;

	if(wb.styleData.fills.length == 0){
		this.fillId=0;
		wb.styleData.fills.push(this);
	}else{
		var isMatched = false;
		var curFillId = 0;
		var fill2 = JSON.parse(JSON.stringify(this));

		while( isMatched == false && curFillId < wb.styleData.fills.length ){
			var fill1 = JSON.parse(JSON.stringify(wb.styleData.fills[curFillId]));

			fill1.fillId=null;
			fill2.fillId=null;
			isMatched = _.isEqual(fill1,fill2);
			if(isMatched){
				this.fillId = curFillId;
			}else{
				curFillId+=1;
			}
		}
		if(!isMatched){
			this.fillId = wb.styleData.fills.length;
			wb.styleData.fills.push(this);
		}
	}

	function genXMLObj(){
		var data = {
			fill:{
				patternFill:[
					{
						'@patternType':this.patternType
					},
					{
						'fgColor':
							{
								'@rgb':this.fgColor
							}
					},
					{
						'bgColor':
							{
								'@rgb':this.bgColor
							}
					}
				]
			}
		};
		return data;
	}
}

var border = function(wb,opts){

	opts=opts?opts:{};
	this.left=opts.left?opts.left:null;
	this.right=opts.right?opts.right:null;
	this.top=opts.top?opts.top:null;
	this.bottom=opts.bottom?opts.bottom:null;
	this.diagonal=opts.diagonal?opts.diagonal:null;
	this.generateXMLObj=genXMLObj;

	if(wb.styleData.borders.length == 0){
		this.borderId=0;
		wb.styleData.borders.push(this);
	}else{
		var isMatched = false;
		var curborderId = 0;
		var border2 = JSON.parse(JSON.stringify(this));

		while( isMatched == false && curborderId < wb.styleData.borders.length ){
			var border1 = JSON.parse(JSON.stringify(wb.styleData.borders[curborderId]));

			border1.borderId=null;
			border2.borderId=null;
			isMatched = _.isEqual(border1,border2);
			if(isMatched){
				this.borderId = curborderId;
			}else{
				curborderId+=1;
			}
		}
		if(!isMatched){
			this.borderId = wb.styleData.borders.length;
			wb.styleData.borders.push(this);
		}
	}

	function genXMLObj(){
		var ordinals = ['left','right','top','bottom','diagonal'];
		var data = {
			border:[]
		};
		ordinals.forEach(function(o){
			if(this[o]!=null){
				data.border.push({
					o:[
						{
							'@style':this[o].style
						},
						{
							'color':{
								'@rgb':cleanColor(this[o]).color
							}
						}
					]
				});
			}else{
				data.border.push(o);
			}
		});
		return data;
	}
}

function cleanColor(val){
	// check for RGB, RGBA or Excel Color Names and return RGBA
	return val;
}


exports.Style = function(opts){
	if(this.styleData.fonts.length == 0){
		var defaultFont = new font(this);
	}
	if(this.styleData.fills.length == 0){
		var defaultFill = new fill(this);
	}
	if(this.styleData.borders.length == 0){
		var defaultBorder = new border(this);
	}
	var newStyle = new style(this,opts);
	this.styleData.cellXfs[newStyle.xf.xfId] = newStyle.xf;

	return newStyle;
}