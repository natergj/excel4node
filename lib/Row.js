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
row.prototype.Hide = hideRow;

/******************************
	Row Method Definitions
 ******************************/
function countCells(){
	var cellCount = parseInt(Object.keys(this.cells).sort(alphaNumSort)[Object.keys(this.cells).length - 1]);
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
		pane['@activePane']='bottomRight';
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

function filter(startCol,endCol,filters){
	var rID = this.attributes.r;
	filters = filters instanceof Array ? filters : [];

	if(typeof(startCol) != 'number' || typeof(endCol) != 'number'){
		startCol = 1;
		endCol = this.cellCount();
	}
	if(startCol instanceof Array){ //no start and end col specified
		filters = startCol;
	}
	
	startCol=startCol?startCol:1;
	endCol=endCol?endCol:this.cellCount()+startCol;

	var thisWS = this.ws.sheet;
	thisWS['autoFilter']=[{
		'@ref':startCol.toExcelAlpha()+rID+':'+endCol.toExcelAlpha()+rID
	}];

	/* Filter Object Definition
		{
			column:Int,
			matchAll:Optional Boolean
			rules:[
				{
					val:String,
					operator:Optional String,
				}
			]
		}
	*/
	filters.forEach(function(f){
		if(typeof(startCol) == 'number' && f.rules instanceof Array){
			var thisFilter = {
				'filterColumn' : [
					{
						'@colId':f.column - 1
					},
					{
						'customFilters':[]
					}
				]
			}
			if(f.matchAll == true){
				thisFilter.filterColumn[1].customFilters.push({'@and':'1'});
			}
			f.rules.forEach(function(r){
				var thisRule = {
					'customFilter':{
						'@val':r.val
					}
				};
				if(r.operator){
					thisRule['customFilter']['@operator'] = r.operator;
				}

				thisFilter.filterColumn[1].customFilters.push(thisRule);
			});
			thisWS['autoFilter'].push(thisFilter);
		}
	});
}

function height(height){
	this.setAttribute('customHeight',1);
	this.setAttribute('ht',height);
}

function hideRow(){

	this.setAttribute('hidden',1);
}

function alphaNumSort(a,b){
	if(parseInt(a) == a && parseInt(b) == b){
		var numA = parseInt(a);
		var numB = parseInt(b);
		return a - b;
	}else{
		if(a < b){
			return -1;
		}else if (b < a){
			return 1;
		}else{
			return 0;
		}
	}
}