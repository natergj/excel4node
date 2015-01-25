var col = require('./Column.js'),
row = require('./Row.js'),
cell = require('./Cell.js'),
image = require('./Images.js'),
xml = require('xmlbuilder');

exports.WorkSheet = function(name, opts){
	var opts = opts?opts:{};

	var thisWS = {};
	thisWS.name = name;
	thisWS.hasGroupings = false;
	thisWS.margins = {
		bottom:1.0,
		footer:.5,
		header:.5,
		left:.75,
		right:.75,
		top:1.0
	}
	thisWS.toXML = generateXML;
	thisWS.getCell = getCell;
	thisWS.sheet={};
	thisWS.sheet['@mc:Ignorable']="x14ac";
	thisWS.sheet['@xmlns']='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	thisWS.sheet['@xmlns:mc']="http://schemas.openxmlformats.org/markup-compatibility/2006";
	thisWS.sheet['@xmlns:r']='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	thisWS.sheet['@xmlns:x14ac']="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
	thisWS.sheet.sheetPr = {
		outlinePr:{
			'@summaryBelow':1
		}
	}
	
	thisWS.sheet.sheetViews = [
		{
			sheetView:[	
				{
					//'@tabSelected':1,
					'@workbookViewId':0
				}
			]
		}
	];
	thisWS.sheet.sheetFormatPr = {
		'@baseColWidth':10,
		'@defaultRowHeight':15,
		'@x14ac:dyDescent':0
	};
	thisWS.sheet.sheetData = [];
	thisWS.sheet.pageMargins = [];
	thisWS.cols = {};
	thisWS.rows = {};


	thisWS.Cell = cell.Cell;
	thisWS.Row = row.Row;
	thisWS.Column = col.Column;
	thisWS.Image = image.Image;
	thisWS.Settings = {
		Outline: {
			SummaryBelow: settings().outlineSummaryBelow
		}
	}

	function setWSOpts(){
		if(opts.margins){
			thisWS.margins.bottom = opts.margins.bottom?opts.margins.bottom:1.0;
			thisWS.margins.footer = opts.margins.footer?opts.margins.footer:.5;
			thisWS.margins.header = opts.margins.header?opts.margins.header:.5;
			thisWS.margins.left = opts.margins.left?opts.margins.left:.75;
			thisWS.margins.right = opts.margins.right?opts.margins.right:.75;
			thisWS.margins.top = opts.margins.top?opts.margins.top:1.0;
		}
		Object.keys(thisWS.margins).forEach(function(k){
			var margin = {};
			margin['@'+k] = thisWS.margins[k];
			thisWS.sheet.pageMargins.push(margin);
		});
		return thisWS;
	}

	function getCell(a,b){
		var props = {};
		if(typeof(a) == 'string'){
			props = a.toExcelRowCol();
		}else if(typeof(a) == 'number' && typeof(b) == 'number'){
			props.row = a;
			props.col = b;
		}else{
			return undefined;
		}
		return thisWS.rows[props.row].cells[props.col];
	}

	function settings(){
		var theseSettings = {};
		theseSettings.outlineSummaryBelow = function(val){
			val = val?1:0;
			thisWS.sheet.sheetPr = {
				outlinePr:{
					'@summaryBelow':val
				}
			}
		}
		return theseSettings;
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

		/*
			Process Rows of data
		*/
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
			Process and merged cells
		*/
		if(thisWS.mergeCells && thisWS.mergeCells.length>=0){
			thisWS.sheet.mergeCells = [];
			thisWS.sheet.mergeCells.push({'@count':thisWS.mergeCells.length});
			thisWS.mergeCells.forEach(function(cr){
				thisWS.sheet.mergeCells.push({'mergeCell':{'@ref':cr}});
			});
		};

		/*
			Excel complains if specific attributes on not in the correct order in the XML doc. 
		*/
		var excelOrder = [
			'sheetPr',
			'sheetViews',
			'sheetFormatPr',
			'cols',
			'sheetData',
			'autoFilter',
			'mergeCells',
			'pageMargins',
			'drawing'
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
		var xmlStr = wsXML.end(xmlOutVars);
		return xmlStr;
	}

	return setWSOpts();
};
