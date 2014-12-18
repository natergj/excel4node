exports.Row = function(rowNum){
	var rowIndex = rowNum - 1;
	var thisWS = this.sheet;
	if(!thisWS.rows[rowIndex]){
		for(var i = thisWS.rows.length; i <= rowIndex; i++){
			thisWS.rows.splice(i,0,new row());
		}
	}
	var thisRow = thisWS.rows[rowIndex];

	console.log(thisRow);
	return thisRow;
}

var row = function(){
	this.cells = [];
	this.cellCount = countCells;
	this.attributes = getAttributes;
	this.addCell = addCell;
	this.Height = height;
	this.Freeze = freeze;
	this.Group = group;
	this.Filter = filter;

	return this;

	function height(height){
		console.log(height);
	}
	function countCells(){

	}
	function getAttributes(){

	}
	function addCell(){

	}
	function freeze(){

	}
	function group(level,isHidden){

	}
	function filter(){

	}
}

