var Column = function(col){
	var theseCols = {};
	var thisWS = this;
	if(!thisWS.cols){
		thisWS.cols={};
	}
	var thisCol = thisWS.cols[col]?thisWS.cols[col]:new column(col,thisWS);

	theseCols.Width = function(w){
		thisCol.setWidth(w);
	}

	theseCols.Freeze = function(scrollTo){
		var sTo = scrollTo?scrollTo:col;
		var sv = thisWS.sheet.sheetViews[0].sheetView;
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
					'@xSplit':col-1
				}
			});
			pane = sv[l-1].pane;
		}else{
			var curTopLeft = pane['@topLeftCell'];
			var points = curTopLeft.toExcelRowCol();
			pane['@topLeftCell']=sTo.toExcelAlpha() + points.row;
			pane['@xSplit']=col-1;
		}
	}

	return theseCols;
}

function column(num,ws,opts){
	var opts=opts?opts:{};
	this.max=num;
	this.min=num;
	this.width=opts.width?opts.width:ws.wb.defaults.colWidth;
	this.customWidth=this.width?1:0;

	ws.cols[num]=this;
	return this;
}

column.prototype.setWidth = setWidth;


function setWidth(w){
	this.width = w;
	this.customWidth = 1;
}

exports.Column = Column;