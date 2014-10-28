var col = require('./Column.js'),
row = require('./Row.js'),
image = require('./Images.js'),
xmlbuilder = require('xmlbuilder');

exports.WorkSheet = function(name){

	var that = this;
	this.name = name;
	this.sheet={};
	this.sheet['@mc:Ignorable']="x14ac";
	this.sheet['@xmlns']='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	this.sheet['@xmlns:mc']="http://schemas.openxmlformats.org/markup-compatibility/2006";
	this.sheet['@xmlns:r']='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	this.sheet['@xmlns:x14ac']="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
	that.sheet.sheetPr = {
		outlinePr:{
			'@summaryBelow':1
		}
	}
	
	this.sheet.sheetViews = [
		{
			sheetView:[	
				{
					'@tabSelected':1,
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
	this.sheet.cols = [
		{
			col:{
				'@customWidth':1,
				'@max':1,
				'@min':1,
				'@width':20
			}
		}
	];
	this.sheet.sheetData = [
		{
			row:[
				{
					'@r':1
				}
			]
		}
	];

	this.Cell = row.Cell;
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
			that.sheet.sheetPr = {
				outlinePr:{
					'@summaryBelow':val
				}
			}
		}
		return this;
	}

	return this;
};
