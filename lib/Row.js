exports.Row = function(rowNum){
	var thisWS = this;
	if(!thisWS.rows[rowNum]){
		thisWS.rows[rowNum] = new row();
	}
	var thisRow = thisWS.rows[rowNum];
	thisRow.ws = thisWS;
	thisRow.setAttribute('r',rowNum);

	//console.log(thisRow);
	return thisRow;
}

var row = function(){
	var thisRow = this;
	thisRow.cells = {};
	thisRow.attributes = {};
	thisRow.setAttribute = setAttribute;
	thisRow.cellCount = countCells;
	thisRow.addCell = addCell;
	thisRow.Height = height;
	thisRow.Freeze = freeze;
	thisRow.Group = group;
	thisRow.Filter = filter;

	return thisRow;

	function height(height){
		//console.log(height);
	}
	function countCells(){

	}
	function setAttribute(attr,val){
		thisRow.attributes[attr] = val;
	}
	function getAttributes(){
	}
	function addCell(){

	}
	function freeze(){

	}
	function group(level,isHidden){
		thisRow.ws.hasGroupings = true;
		var hidden=isHidden?1:0;
		thisRow.setAttribute('outlineLevel',level);
		thisRow.setAttribute('hidden',hidden);
		return thisRow;
	}
	function filter(){

	}
}

