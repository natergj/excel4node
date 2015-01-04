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
	thisRow.setAttribute = setAttribute;
	thisRow.cellCount = countCells;
	thisRow.Height = height;
	thisRow.Freeze = freeze;
	thisRow.Group = group;
	thisRow.Filter = filter;

	function height(height){
		thisRow.setAttribute('customHeight',1);
		thisRow.setAttribute('ht',height);
	}
	function countCells(){
		var cellCount = parseInt(Object.keys(thisRow.cells).sort()[Object.keys(thisRow.cells).length - 1]);
		if(isNaN(cellCount)){
			cellCount = 1;
		}
		return cellCount;
	}
	function setAttribute(attr,val){
		thisRow.attributes[attr] = val;
	}
	function freeze(scrollTo){
		var rID = thisRow.attributes.r;
		var sTo = scrollTo?scrollTo:rID;
		var sv = thisRow.ws.sheet.sheetViews[0].sheetView;
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
		thisRow.ws.hasGroupings = true;
		var hidden=isHidden?1:0;
		thisRow.setAttribute('outlineLevel',level);
		thisRow.setAttribute('hidden',hidden);
		return thisRow;
	}
	function filter(startCol,endCol){
		var rID = thisRow.attributes.r;
		startCol=startCol?startCol:1;
		endCol=endCol?endCol:thisRow.cellCount();
		thisRow.ws.sheet['autoFilter']={
			'@ref':startCol.toExcelAlpha()+rID+':'+endCol.toExcelAlpha()+rID
		}
	}

	return thisRow;
}

