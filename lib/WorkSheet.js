var col = require('./Column.js'),
row = require('./Row.js'),
xmlbuilder = require('xmlbuilder');

exports.WorkSheet = function(name){

	this.name = name;
	this.sheet={};
	this.sheet['@xmlns']='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	this.sheet['@xmlns:r']='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	this.sheet.cols = [{
		col:{
			'@customWidth':1,
			'@max':1,
			'@min':1,
			'@width':20
		}
	}];
	this.sheet.sheetData = [];

	this.Cell = row.Cell;
	this.Row = row.Row;
	this.Column = col.Column;

	return this;
};
