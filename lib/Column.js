var Column = function(colNum){
	var thisWS = this;
	if(!thisWS.cols){
		thisWS.cols={};
	}
	if(!thisWS.cols[colNum]){
		var newCol = new column();
		newCol.setAttribute('max',colNum);
		newCol.setAttribute('min',colNum);
		newCol.setAttribute('customWidth',0);
		newCol.setAttribute('width',thisWS.wb.defaults.colWidth);
		thisWS.cols[colNum] = newCol;
	}

	var thisCol = thisWS.cols[colNum];
	thisCol.ws = thisWS;

	return thisCol;
}


var column = function(){
	return this;
}

/******************************
	Column Methods
 ******************************/
column.prototype.setAttribute = setAttribute;
column.prototype.Hide = hideColumn;
column.prototype.Group = setGroup;
column.prototype.Width = setWidth;
column.prototype.Freeze = freezeColumn;

/******************************
	Column Method Definitions
 ******************************/
function setWidth(w){
	this.setAttribute('width',w);
	this.setAttribute('customWidth',1);

	return this;
}

function hideColumn(){

	this.setAttribute('hidden',1);

	return this;
}

function setAttribute(attr,val){

	this[attr] = val;

}

function setGroup(level,isHidden){
	this.ws.hasGroupings = true;
	var hidden=isHidden?1:0;
	this.setAttribute('outlineLevel',level);
	this.setAttribute('hidden',hidden);
	return this;
}

function freezeColumn(scrollTo){
	var sTo = scrollTo?scrollTo:this.min;
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
				'@activePane':'topRight',
				'@state':'frozen',
				'@topLeftCell':sTo.toExcelAlpha() + '1', 
				'@xSplit':this.min-1
			}
		});
		pane = sv[l-1].pane;
	}else{
		var curTopLeft = pane['@topLeftCell'];
		var points = curTopLeft.toExcelRowCol();
		pane['@activePane']='bottomRight';
		pane['@topLeftCell']=sTo.toExcelAlpha() + points.row;
		pane['@xSplit']=this.min-1;
	}

	return this;
}


exports.Column = Column;

