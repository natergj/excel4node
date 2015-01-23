var wb = require('./WorkBook.js'),
style = require('./Style.js');

exports.WorkBook = wb.WorkBook;
exports.Style = style.Style;

Number.prototype.toExcelAlpha = function(isCaps){
	//converts column number to text equivalent for Excel
	var isCaps = isCaps == undefined?true:isCaps;

    var d = (this - 1) / 26;
    d = Math.floor(d);
    if (d > 0){
        r = (isCaps ? 65 : 97) + (d - 1);
    }

    var num = 0;
    if (this % 26 > 0){
        num = (this - (26 * d)) % 26;
    }else{
        num = 26;
    }

    var c = (isCaps ? 65 : 97) + (num - 1);
    if (d > 0){
        return String.fromCharCode(r) + String.fromCharCode(c);
    }else{
        return String.fromCharCode(c);
    }  
}

String.prototype.toExcelRowCol = function(){
	var re1 = /\d/;
	var re2 = /\D/;
	var alpha = this.split(re1).filter(function(el){return el!=''});
	var numeric = this.split(re2).filter(function(el){return el!=''});
	if(alpha.length > 1 || numeric.length > 1){
		console.log(this+' does not appear to be a valid excel cell represenation.');
		return {row:0,col:0};
	}else{
		var thisCol = 0;
		var colParts = alpha[0].split('');
		colParts.forEach(function(a,i){
			if(i == 0 && colParts.length > 1){
				thisCol += (a.toUpperCase().charCodeAt() - 64) * 26; 
			}else{
				thisCol += a.toUpperCase().charCodeAt() - 64;
			}
		});
		return {row:parseInt(numeric[0]),col:thisCol}
	}
}

Date.prototype.getExcelTS = function(){
	var epoch = new Date(1899,11,31);
	var dt = this.setDate(this.getDate()+1);
	var ts = (dt-epoch)/(1000*60*60*24);
	return ts;
}