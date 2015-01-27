exports.Row = function(rowNum){
	var thisWS = this;
	if(!thisWS.rows[rowNum]){
		thisWS.rows[rowNum] = new row();
	}
	var thisRow = thisWS.rows[rowNum];
	thisRow.ws = thisWS;
	thisRow.setAttribute('r',rowNum);
	thisRow.setAttribute('spans','1:'+thisRow.cellCount());

	//console.log(thisRow);
	return thisRow;
}

var row = function(){
	var thisRow = this;
	thisRow.cells = {};
	thisRow.attributes = {};
	return thisRow;
}

/******************************
	Row Methods
 ******************************/
row.prototype.Height = height;
row.prototype.setAttribute = setAttribute;
row.prototype.cellCount = countCells;
row.prototype.Freeze = freeze;
row.prototype.Group = group;
row.prototype.Filter = filter;

/******************************
	Row Method Definitions
 ******************************/
function countCells(){
	var cellCount = parseInt(Object.keys(this.cells).sort()[Object.keys(this.cells).length - 1]);
	if(isNaN(cellCount)){
		cellCount = 1;
	}
	return cellCount;
}
function setAttribute(attr,val){
	this.attributes[attr] = val;
}
function freeze(scrollTo){
	var rID = this.attributes.r;
	var sTo = scrollTo?scrollTo:rID;
	var sv = this.ws.sheet.sheetViews[0].sheetView;
	var pane;
	var foundPane = false;
	sv.forEach(function(v,i){
		if(Object.keys(v).indexOf('pane') >= 0){
			pane = sv[i].pane;
			foundPane = true;
		}
	});
	if(!foundPane){
		var l = sv.push({
			pane:{
				'@activePane':'bottomLeft',
				'@state':'frozen',
				'@topLeftCell':'A'+sTo, 
				'@ySplit':rID-1
			}
		});
		pane = sv[l-1].pane;
	}else{
		var curTopLeft = pane['@topLeftCell'];
		var points = curTopLeft.toExcelRowCol();
		pane['@topLeftCell']=points.col.toExcelAlpha() + sTo;
		pane['@ySplit']=rID-1;
	}
}
function group(level,isHidden){
	this.ws.hasGroupings = true;
	var hidden=isHidden?1:0;
	this.setAttribute('outlineLevel',level);
	this.setAttribute('hidden',hidden);
	return this;
}
function filter(startCol,endCol){
	var rID = this.attributes.r;
	startCol=startCol?startCol:1;
	endCol=endCol?endCol:this.cellCount();
	this.ws.sheet['autoFilter']={
		'@ref':startCol.toExcelAlpha()+rID+':'+endCol.toExcelAlpha()+rID
	}
}
function height(height){
	this.setAttribute('customHeight',1);
	this.setAttribute('ht',height);
}