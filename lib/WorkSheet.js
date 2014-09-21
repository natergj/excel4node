var col = require('./Column.js'),
row = require('./Row.js'),
image = require('./Images.js'),
xmlbuilder = require('xmlbuilder');

exports.WorkSheet = function(name){

	this.name = name;
	this.sheet={};
	this.sheet['@mc:Ignorable']="x14ac";
	this.sheet['@xmlns']='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	this.sheet['@xmlns:mc']="http://schemas.openxmlformats.org/markup-compatibility/2006";
	this.sheet['@xmlns:r']='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	this.sheet['@xmlns:x14ac']="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

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

	return this;
};
