var Column = require('./Column.js'),
Row = require('./Row.js'),
Cell = require('./Cell.js'),
Image = require('./Images.js'),
xml = require('xmlbuilder');

var WorkSheet = function(wb){
	this.wb = wb;
	this.opts = {};
	this.name = '';
	this.hasGroupings = false;
	this.margins = {
		bottom:1.0,
		footer:.5,
		header:.5,
		left:.75,
		right:.75,
		top:1.0
	};
	this.printOptions = {
		centerHorizontal: false,
		centerVertical: true
	}
	this.sheetView = {
		workbookViewId:0,
		zoomScale:100,
		zoomScaleNormal:100,
		zoomScalePageLayoutView:100
	}
	this.sheet={};
	this.sheet['@mc:Ignorable']="x14ac";
	this.sheet['@xmlns']='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	this.sheet['@xmlns:mc']="http://schemas.openxmlformats.org/markup-compatibility/2006";
	this.sheet['@xmlns:r']='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	this.sheet['@xmlns:x14ac']="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
	this.sheet.sheetPr = {
		outlinePr:{
			'@summaryBelow':1
		}
	}
	this.sheet.sheetViews = [
		{
			sheetView:[]
		}
	];
	this.sheet.sheetFormatPr = {
		'@baseColWidth':10,
		'@defaultRowHeight':15,
		'@x14ac:dyDescent':0
	};
	this.sheet.sheetData = [];
	this.sheet.pageMargins = [];
	this.sheet.printOptions = [];
	this.cols = {};
	this.rows = {};

	this.Settings = {
		Outline:{
			SummaryBelow : settings(this).outlineSummaryBelow
		}
	}
};

WorkSheet.prototype.setName = function(name){
	this.name = name;
}
WorkSheet.prototype.Cell = Cell.Cell;
WorkSheet.prototype.Row = Row.Row;
WorkSheet.prototype.Column = Column.Column;
WorkSheet.prototype.Image = Image.Image;
WorkSheet.prototype.getCell = getCell;
WorkSheet.prototype.setWSOpts = setWSOpts;
WorkSheet.prototype.toXML = toXML;

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

function settings(thisWS){
	var theseSettings = {};
	theseSettings.outlineSummaryBelow = function(val){
		console.log("####################################################################################");
		console.log("# WorkSheet.Settings is deprecated and will be removed in version 1.0.0            #");
		console.log("# Create WorkBooks with opts paramater instead.                                    #");
		console.log("####################################################################################");
		val = val?1:0;
		thisWS.sheet.sheetPr = {
			outlinePr:{
				'@summaryBelow':val
			}
		}
	}
	return theseSettings;
}

function setWSOpts(opts){
	var opts = opts?opts:{};
	var thisWS = this;
	// Set Margins
	if(opts.margins){
		this.margins.bottom = opts.margins.bottom?opts.margins.bottom:1.0;
		this.margins.footer = opts.margins.footer?opts.margins.footer:.5;
		this.margins.header = opts.margins.header?opts.margins.header:.5;
		this.margins.left = opts.margins.left?opts.margins.left:.75;
		this.margins.right = opts.margins.right?opts.margins.right:.75;
		this.margins.top = opts.margins.top?opts.margins.top:1.0;
	}
	Object.keys(this.margins).forEach(function(k){
		var margin = {};
		margin['@'+k] = thisWS.margins[k];
		thisWS.sheet.pageMargins.push(margin);
	});

	// Set Print Options
	if(opts.printOptions){
		this.printOptions.centerHorizontal = opts.printOptions.centerHorizontal?opts.printOptions.centerHorizontal:false;
		this.printOptions.centerVertical = opts.printOptions.centerVertical?opts.printOptions.centerVertical:false;
	}
	this.sheet.printOptions.push({'@horizontalCentered':this.printOptions.centerHorizontal?1:0});
	this.sheet.printOptions.push({'@verticalCentered':this.printOptions.centerVertical?1:0});

	// Set Page View options
	var thisView = this.sheet.sheetViews[0].sheetView;
	if(opts.view){
		if(parseInt(opts.view.zoom) != opts.view.zoom){
			console.log("invalid value for zoom. value must be an integer. value was %s",opts.view.zoom);
			opts.view.zoom = 100;
		}
		this.sheetView.zoomScale = opts.view.zoom?opts.view.zoom:100;
		this.sheetView.zoomScaleNormal = opts.view.zoom?opts.view.zoom:100;
		this.sheetView.zoomScalePageLayoutView = opts.view.zoom?opts.view.zoom:100;
	}

	// Set Outline Options
	if(opts.outline){
		thisWS.sheet.sheetPr = {
			outlinePr:{
				'@summaryBelow':opts.outline.summaryBelow==false?0:1
			}
		}
	}


	thisView.push({'@workbookViewId':this.sheetView.workbookViewId?this.sheetView.workbookViewId:0});
	thisView.push({'@zoomScale':this.sheetView.zoomScale?this.sheetView.zoomScale:100});
	thisView.push({'@zoomScaleNormal':this.sheetView.zoomScaleNormal?this.sheetView.zoomScaleNormal:100});
	thisView.push({'@zoomScalePageLayoutView':this.sheetView.zoomScalePageLayoutView?this.sheetView.zoomScalePageLayoutView:100});
}

function toXML(){
	var thisWS = this;
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
				startRow:1,
				endRow:1,
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
			var isHidden = curRow.attributes.hidden==1?curRow.attributes.hidden:0;
			var rowNum = curRow.attributes.r;

			outlineLevels[0].endRow=rID;
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
				outlineLevels[thisLevel].isHidden = isHidden;
			}

			if(thisLevel != outlineLevels.curHighestLevel || rID == lastRowNum){
				if(summaryBelow && thisLevel != outlineLevels.curHighestLevel){

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
							var rowToCollapse = outlineLevels[outlineLevels.curHighestLevel].startRow;
						}else{
							var rowToCollapse = outlineLevels[thisLevel].startRow;
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
		'printOptions',
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


module.exports = WorkSheet;