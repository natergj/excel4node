exports.Style = function(){
	var that = this;
	this.data = {
		numFmt:[],
		font:[
			{
				sz:{
					'@val':12
				}
			},
			{
				color:{
					'@theme':1
				}
			},
			{
				name:{
					'@val':'Calibri'
				}
			},
			{
				scheme:{
					'@val':'minor'
				}
			}
		],
		border:[],
		xf:{
			'@applyNumberFormat':0,
			'@applyFill':0,
			'@applyFont':0,
			'@borderId':0,
			'@fillId':0,
			'@fontId':0,
			'@numFmtId':164,
			'@xfId':0
		},
		IDs:{
			numFmt:0,
			font:0,
			fill:0,
			border:0,
			xf:0
		},
		patternFill:{
			patternType:'none',
			fgColor:{
				rgb:'00000000'
			},
			bgColor:{
				indexed:64
			}
		}
	}

	this.setFont = function(){
		that.data.xf['@applyFont']=1;
		if(that.data.IDs.font == 0){
			var curCount = that.wb.workbook.styles.styleSheet.fonts[0]['@count'];
			that.wb.workbook.styles.styleSheet.fonts[0]['@count'] = curCount+1;
			that.wb.workbook.styles.styleSheet.fonts.push({font:that.data.font});
			that.data.IDs.font = curCount+1;
			that.data.xf['@fontId']=curCount;
		}else{
			that.wb.workbook.styles.styleSheet.fonts[that.data.IDs.font] = {font:that.data.font};
		}
	}
	this.setXF = function(){
		if(that.data.IDs.xf == 0){
			var curCount = that.wb.workbook.styles.styleSheet.cellXfs[0]['@count'];
			that.wb.workbook.styles.styleSheet.cellXfs[0]['@count'] = curCount+1;
			that.wb.workbook.styles.styleSheet.cellXfs.push({xf:that.data.xf});
			that.data.IDs.xf = curCount;
		}else{
			that.wb.workbook.styles.styleSheet.cellXfs[that.data.IDs.xf + 1] = {xf:that.data.xf};
		}
	}
	this.setNumFmt = function(){
		that.data.xf['@applyNumberFormat']=1;
		if(that.data.IDs.numFmt == 0){
			var curCount = that.wb.workbook.styles.styleSheet.numFmts[0]['@count'];
			that.wb.workbook.styles.styleSheet.numFmts[0]['@count'] = curCount+1;
			that.wb.workbook.styles.styleSheet.numFmts.push({numFmt:that.data.numFmt});
			that.data.IDs.numFmt = curCount;
		}else{
			that.wb.workbook.styles.styleSheet.numFmts[that.data.IDs.numFmt + 1] = {numFmt:that.data.numFmt};
		}
	}
	this.setFill = function(){
		that.data.xf['@applyFill']=1;
		that.data.fill = {patternFill:obj2XMLobj(that.data.patternFill)};
		if(that.data.IDs.fill == 0){
			var curCount = that.wb.workbook.styles.styleSheet.fills[0]['@count'];
			that.wb.workbook.styles.styleSheet.fills[0]['@count'] = curCount+1;
			that.wb.workbook.styles.styleSheet.fills.push({fill:that.data.fill});
			that.data.IDs.fill = curCount+1;
			that.data.xf['@fillId']=curCount;
		}else{
			that.wb.workbook.styles.styleSheet.fills[that.data.IDs.fill] = {fill:that.data.fill};
		}
	}

	this.update = function(){
		that.setFont();
		that.setXF();
	}

	this.Font = {
		Family:function(name){
			var update = false;
			that.data.font.forEach(function(v){
				if(v['name'] != undefined){
					v['name'] = {'@val':name};
					update = true;
				}
			});
			if(!update){
				that.data.font.push({
					name: {
						'@val':name
					}
				});
			}
			that.update();
		},
		Bold:function(){
			that.data.font.unshift("b");
			that.update();
		},
		Italics:function(){
			var insertPos = 0;
			if(that.data.font.indexOf("b")>=0){
				insertPos++;
			}
			that.data.font.splice(insertPos,0,"i");
			that.update();
		},
		Underline:function(){
			var insertPos = 0;
			if(that.data.font.indexOf("b")>=0){
				insertPos++;
			}
			if(that.data.font.indexOf("i")>=0){
				insertPos++;
			}
			that.data.font.splice(insertPos,0,"u");
			that.update();
		},
		Size:function(size){
			var update = false;
			that.data.font.forEach(function(v){
				if(v['sz'] != undefined){
					v['sz'] = {'@val':size};
					update = true;
				}
			});
			if(!update){
				that.data.font.push({
					sz: {
						'@val':size
					}
				});
			}
			that.update();
		},
		Color:function(rgb){
			var update = false;
			that.data.font.forEach(function(v){
				if(v['color'] != undefined){
					v['color'] = {'@rgb':rgb};
					update = true;
				}
			});
			if(!update){
				that.data.font.push({
					color: {
						'@rgb':rgb
					}
				});
			}
			that.update();
		},
		WrapText:function(wrap){
			if(that.data.xf.alignment == undefined){
				that.data.xf.alignment = {};
			}
			if(wrap){
				that.data.xf.alignment['@wrapText'] = 1;
			}else{
				that.data.xf.alignment['@wrapText'] = 0;
			}
			that.update();
		},
		Alignment:{
			Vertical:function(align){
				if(that.data.xf.alignment == undefined){
					that.data.xf.alignment = {};
				}
				that.data.xf.alignment['@vertical'] = align
				that.update();
			},
			Horizontal:function(align){
				if(that.data.xf.alignment == undefined){
					that.data.xf.alignment = {};
				}
				that.data.xf.alignment['@horizontal'] = align
				that.update();
			}
		}
	}

	this.Number = {
		Format:function(fmt){
			var curCount = that.wb.workbook.styles.styleSheet.numFmts[0]['@count'];
			var fmtCode = curCount + 165;
			that.data.numFmt.push({'@formatCode':fmt,'@numFmtId':fmtCode});
			that.data.xf['@numFmtId']=fmtCode;
			that.setNumFmt();
			that.update();
		}

	}

	this.Fill = {
		Color:function(rgb){
			that.data.patternFill.fgColor = {
				'rgb':rgb
			}
			that.setFill();
			that.update();
		},
		Pattern:function(pattern){
			that.data.patternFill.patternType = pattern;
			that.setFill();
			that.update();
		}
	}


	this.Border = function(obj){
		/*
			obj should be in form 
			style is required, color is optional and defaults to black if not specified.
			not all ordinals are required. No board will be on side that is not specified.
			{
				left:{
					style:'style',
					color:'rgb'
				},
				right:{
					style:'style',
					color:'rgb'
				},
				top:{
					style:'style',
					color:'rgb'
				},
				bottom:{
					style:'style',
					color:'rgb'
				},
				diagonal:{
					style:'style',
					color:'rgb'
				}
			}
		*/
		var validStyles = [
			"hair",
			"dotted",
			"dashDotDot",
			"dashDot",
			"dashed",
			"thin",
			"mediumDashDotDot",
			"slantDashDot",
			"mediumDashDot",
			"mediumDashed",
			"medium",
			"thick",
			"double"
		];

		var thisBorder = {
			border:[
				'left',
				'right',
				'top',
				'bottom',
				'diagonal'
			]
		};

		/*
			takes input object and replaces ordinal border definitions
		*/
		Object.keys(obj).forEach(function(ord){
			var k = thisBorder.border.indexOf(ord);
			if(k>=0){
				/*
					will replace ordinal string in border array with object in format useable by xmlbuilder
					{
						ordinal:{
							style:'style',
							color:'color'
						}						
					}
				*/
				thisBorder.border[k] = {};
				thisBorder.border[k][ord]={}
				if(obj[ord]['style'] && validStyles.indexOf(obj[ord].style) >= 0){
					thisBorder.border[k][ord]['@style']=obj[ord].style;
				}
				if(obj[ord]['color']){
					thisBorder.border[k][ord]['color'] = {
						'@rgb':obj[ord]['color']
					}
				}
			}
		});

		that.data.IDs.border = that.wb.workbook.styles.styleSheet.borders.push(thisBorder) - 2;

		that.data.xf['@applyBorder'] = 1;
		that.data.xf['@borderId'] = that.data.IDs.border;

		return that;
	}

	return this;
}

function obj2XMLobj(o){
	var arr = [];
	Object.keys(o).forEach(function(k){
		if(typeof(o[k]) == 'object'){
			var tmpObj = {};
			tmpObj[k]=obj2XMLobj(o[k]);
			arr.push(tmpObj);
		}else{
			var thisKey = '@'+k;
			var tmpObj = {};
			tmpObj[thisKey] = o[k];
			arr.push(tmpObj);
		}
	});
	return arr;
}

