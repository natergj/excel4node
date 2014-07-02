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

		this.sheet.cols.push(colObj);
	}

	return this;
}