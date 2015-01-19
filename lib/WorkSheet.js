var col = require('./Column.js'),
row = require('./Row.js'),
cell = require('./Cell.js'),
image = require('./Images.js'),
xml = require('xmlbuilder');

exports.WorkSheet = function(name){

	var thisWS = this;
	this.name = name;
	this.hasGroupings = false;
	this.toXML = generateXML;
	this.sheet={};
	this.sheet['@mc:Ignorable']="x14ac";
	this.sheet['@xmlns']='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	this.sheet['@xmlns:mc']="http://schemas.openxmlformats.org/markup-compatibility/2006";
	this.sheet['@xmlns:r']='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	this.sheet['@xmlns:x14ac']="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
	thisWS.sheet.sheetPr = {
		outlinePr:{
			'@summaryBelow':1
		}
	}
	
	this.sheet.sheetViews = [
		{
			sheetView:[	
				{
					//'@tabSelected':1,
					'@workbookViewId':0
				}
			]
		}
	];
	this.sheet.sheetFormatPr = {
		'@baseColWidth':10,
		'@defaultRowHeight':15,
		'@x14ac:dyDescent':0
	};
	this.sheet.sheetData = [];
	this.cols = {};
	this.rows = {};


	this.Cell = cell.Cell;
	this.Row = row.Row;
	this.Column = col.Column;
	this.Image = image.Image;
	this.Settings = {
		Outline: {
			SummaryBelow: settings().outlineSummaryBelow
		}
	}

	function settings(){
		this.outlineSummaryBelow = function(val){
			val = val?1:0;
			thisWS.sheet.sheetPr = {
				outlinePr:{
					'@summaryBelow':val
				}
			}
		}
		return this;
	}

	function generateXML(){
		var sheetData = thisWS.sheet.sheetData;
		var xmlOutVars = {};
		var xmlDebugVars = { pretty: true, indent: '  ',newline: '\n' };

		/*
			Process Column Definitions
		*/
		if(Object.keys(thisWS.cols).length > 0){
			thisWS.sheet.cols = [];
			Object.keys(thisWS.cols).forEach(function(i){
				var c = thisWS.cols[i];
				var thisCol = {col:[]};
				thisCol.col.push({'@customWidth':c.customWidth});
				thisCol.col.push({'@min':c.min});
				thisCol.col.push({'@max':c.max});
				thisCol.col.push({'@width':c.width});
				if(!thisWS.sheet.cols){
					thisWS.sheet.cols=[];
				}
				thisWS.sheet.cols.push(thisCol);
			});
		}

		/*
			Process groupings and add collapsed attributes to rows where applicable
		*/
		if(thisWS.hasGroupings){
			var lastRowNum = Object.keys(thisWS.rows).sort()[Object.keys(thisWS.rows).length - 1];

			var outlineLevels = {
				curHighestLevel:0,
				0:{
					startRow:0,
					endRow:0,
					isHidden:0
				}
			};
			var summaryBelow = parseInt(thisWS.sheet.sheetPr.outlinePr['@summaryBelow']) == 0 ? false : true;
			if(summaryBelow && thisWS.rows[lastRowNum].attributes.hidden){
				thisWS.Row(parseInt(lastRowNum)+1);
			}
			Object.keys(thisWS.rows).forEach(function(rNum,i){
				var rID = parseInt(rNum);
				var curRow = thisWS.rows[rNum];
				var thisLevel = curRow.attributes.outlineLevel?curRow.attributes.outlineLevel:0;
				var isHidden = curRow.attributes.hidden?curRow.attributes.hidden:0;
				var rowNum = curRow.attributes.r;

				outlineLevels[0].endRow=i;
				outlineLevels[0].isHidden=isHidden;

				if(typeof(outlineLevels[thisLevel]) == 'undefined'){
					outlineLevels[thisLevel] = {
						startRow:rID,
						endRow:rID,
						isHidden:isHidden
					}
				}

				if(thisLevel <= outlineLevels.curHighestLevel){
					outlineLevels[thisLevel].endRow = rID;
				}

				if(thisLevel != outlineLevels.curHighestLevel || rID == lastRowNum){
					if(summaryBelow && thisLevel < outlineLevels.curHighestLevel){
						if(rID == lastRowNum){
							thisLevel=1;
						}
						for(oLi = outlineLevels.curHighestLevel; oLi > thisLevel; oLi--){
							if(outlineLevels[oLi]){
								var rowToCollapse = outlineLevels[oLi].endRow + 1;
								var lastRow = thisWS.Row(rowToCollapse);
								lastRow.setAttribute('collapsed',outlineLevels[oLi].isHidden);
								delete outlineLevels[oLi];
							}
						}
					}else if(!summaryBelow && thisLevel != outlineLevels.curHighestLevel){
						if(outlineLevels[thisLevel]){
							if(thisLevel>outlineLevels.curHighestLevel){
								var rowToCollapse = outlineLevels[outlineLevels.curHighestLevel].startRow - 1;
							}else{
								var rowToCollapse = outlineLevels[thisLevel].startRow - 1;
							}
							var lastRow = thisWS.Row(rowToCollapse);
							lastRow.setAttribute('collapsed',outlineLevels[thisLevel].isHidden);
							outlineLevels[thisLevel].startRow = rowNum;
						}
					}
				}
				if(thisLevel!=outlineLevels.curHighestLevel){
					outlineLevels.curHighestLevel=thisLevel;
				}
			});
		}

		Object.keys(thisWS.rows).forEach(function(r,i){
			var thisRow = {row:[]};
			Object.keys(thisWS.rows[r].attributes).forEach(function(a,i){
				var attr = '@'+a;
				var obj = {};
				obj[attr] = thisWS.rows[r].attributes[a];
				thisRow.row.push(obj);
			});
			Object.keys(thisWS.rows[r].cells).forEach(function(c,i){
				var thisCellIndex = thisRow.row.push({'c':{}});
				var thisCell = thisRow.row[thisCellIndex - 1]['c'];
				Object.keys(thisWS.rows[r].cells[c].attributes).forEach(function(a,i){
					thisCell['@'+a] = thisWS.rows[r].cells[c].attributes[a];
				});
				Object.keys(thisWS.rows[r].cells[c].children).forEach(function(v,i){
					thisCell[v] = thisWS.rows[r].cells[c].children[v];
				});
			})
			sheetData.push(thisRow)
		})

		/*
			Excel complains if specific attributes on not in the correct order in the XML doc. 
		*/
		var excelOrder = [
			'sheetPr',
			'sheetViews',
			'sheetFormatPr',
			'cols',
			'sheetData'
		];
		var orderedDef = [];
		Object.keys(thisWS.sheet).forEach(function(k){
			if(k.charAt(0)==='@'){
				var def = {};
				def[k] = thisWS.sheet[k];
				orderedDef.push(def);
			}
		});
		excelOrder.forEach(function(k){
			if(thisWS.sheet[k]){
				var def = {};
				def[k] = thisWS.sheet[k];
				orderedDef.push(def);
			}
		});
		var wsXML = xml.create({'worksheet':orderedDef});
		var xmlStr = wsXML.end(xmlDebugVars);
		console.log(xmlStr);
		return xmlStr;
	}

	return this;
};
