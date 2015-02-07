var wb = require('./WorkBook.js'),
style = require('./Style.js');

exports.WorkBook = wb.WorkBook;
exports.Style = style.Style;

Number.prototype.toExcelAlpha = function(isCaps){
	//converts column number to text equivalent for Excel
	var chars = [];
	var n = this;
	var e = 0;
	for(var i = getColsInStringWithLength(n); i > 0; i--){
		var startCharCode = i>1?64:65;
		if(getDigitsInCharCount(e) > 0){
			var charNum = parseInt(26*(n-1)/Math.pow(26,i)) % getDigitsInCharCount(e) % 26;
		}else{
			var charNum = parseInt(26*(n-1)/Math.pow(26,i))
					- ( getDigitsInCharCount(e) 
						* parseInt(26*(n-1)/Math.pow(26,i+1)) );
		}
		
		chars.push(String.fromCharCode(charNum + startCharCode));
		e++;
	}
	return chars.join('');
}

String.prototype.toExcelRowCol = function(){
	var re1 = /\d/;
	var re2 = /\D/;
	var alpha = this.split(re1).filter(function(el){return el!=''});
	var numeric = this.split(re2).filter(function(el){return el!=''});

	var colNum = 0;
	var letters = alpha[0].split('');
	letters.reverse();
	letters.forEach(function(k,i){
		colNum += Math.pow(26,i) * (k.toUpperCase().charCodeAt(0) - 64)
	});

	return {
		row:parseInt(numeric[0]),
		col:parseInt(colNum)
	}
}

Date.prototype.getExcelTS = function(){
	var epoch = new Date(1899,11,31);
	var dt = this.setDate(this.getDate()+1);
	var ts = (dt-epoch)/(1000*60*60*24);
	return ts;
}


function getColsInStringWithLength(n){
	var chars = [];
	var charCount = 0;

	while(n > getDigitsInCharCount(charCount) || charCount == 0){
		charCount++;
	}
	return charCount;
}

function getDigitsInCharCount(c){
	var num = 0;
	for(var i = 1; i <= c; i++){
		num+=Math.pow(26,i);
	}
	return num;
}
