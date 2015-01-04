exports.Column = function(col){

	this.Width = function(w){
		var colObj = {
			col:{
				'@customWidth':1,
				'@max':col,
				'@min':col,
				'@width':w
			}
		}

		this.cols.push(colObj);
	}

	this.Freeze = function(scrollTo){
		var sTo = scrollTo?scrollTo:col;
		var sv = this.sheet.sheetViews[0].sheetView;
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

	return this;
}