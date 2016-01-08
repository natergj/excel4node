module.exports = {
    generateRId: generateRId,
    getHashOfPassword: getHashOfPassword,
    aSyncForEach: aSyncForEach
};

function generateRId() {
    var text = 'R';
    var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (var i = 0; i < 16; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
}

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
function getHashOfPassword(str) {
    var curHash = '0000';
    for (var i = str.length - 1; i >= 0; i--) {
        curHash = _getHashForChar(str[i], curHash);
    }
    var curHashBin = parseInt(curHash, 16).toString(2);
    var charCountBin = parseInt(str.length, 10).toString(2);
    var saltBin = parseInt('CE4B', 16).toString(2);

    var firstXOR = _bitXOR(curHashBin, charCountBin);
    var finalHashBin = _bitXOR(firstXOR, saltBin);
    var finalHash = String('0000' + parseInt(finalHashBin, 2).toString(16).toUpperCase()).slice(-4);

    return finalHash;
}


/*
 * Helper Functions
 */

function _rotateBinary(bin) {
    return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
}

function _getHashForChar(char, hash) {    
    hash = hash ? hash : '0000';
    var charCode = char.charCodeAt(0);
    var hashBin = parseInt(hash, 16).toString(2);
    var charBin = parseInt(charCode, 10).toString(2);
    hashBin = String('000000000000000' + hashBin).substr(-15);
    charBin = String('000000000000000' + charBin).substr(-15);
    var nextHash = _bitXOR(hashBin, charBin);
    nextHash = _rotateBinary(nextHash);
    nextHash = parseInt(nextHash, 2).toString(16);

    return nextHash;
}

function _bitXOR(a, b) {
    var maxLength = a.length > b.length ? a.length : b.length;

    var padString = '';
    for (var i = 0; i < maxLength; i++) {
        padString += '0';
    }

    a = String(padString + a).substr(-maxLength);
    b = String(padString + b).substr(-maxLength);

    var response = '';
    for(var i = 0; i < a.length; i++) {
        response += a[i] === b[i] ? 0 : 1;
    }
    return response;
}

function aSyncForEach(fn) {
    this.forEach(function (v, i) {
        setTimeout(fn(v, i), 0);
    });
}
