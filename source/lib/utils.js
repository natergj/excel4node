let _bitXOR = (a, b) => {
    let maxLength = a.length > b.length ? a.length : b.length;

    let padString = '';
    for (let i = 0; i < maxLength; i++) {
        padString += '0';
    }

    a = String(padString + a).substr(-maxLength);
    b = String(padString + b).substr(-maxLength);

    let response = '';
    for(let i = 0; i < a.length; i++) {
        response += a[i] === b[i] ? 0 : 1;
    }
    return response;
};

let generateRId = () => {
    let text = 'R';
    let possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (let i = 0; i < 16; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
};

let _rotateBinary = (bin) => {
    return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
};

let _getHashForChar = (char, hash) => {    
    hash = hash ? hash : '0000';
    let charCode = char.charCodeAt(0);
    let hashBin = parseInt(hash, 16).toString(2);
    let charBin = parseInt(charCode, 10).toString(2);
    hashBin = String('000000000000000' + hashBin).substr(-15);
    charBin = String('000000000000000' + charBin).substr(-15);
    let nextHash = _bitXOR(hashBin, charBin);
    nextHash = _rotateBinary(nextHash);
    nextHash = parseInt(nextHash, 2).toString(16);

    return nextHash;
};

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
let getHashOfPassword = (str) => {
    let curHash = '0000';
    for (let i = str.length - 1; i >= 0; i--) {
        curHash = _getHashForChar(str[i], curHash);
    }
    let curHashBin = parseInt(curHash, 16).toString(2);
    let charCountBin = parseInt(str.length, 10).toString(2);
    let saltBin = parseInt('CE4B', 16).toString(2);

    let firstXOR = _bitXOR(curHashBin, charCountBin);
    let finalHashBin = _bitXOR(firstXOR, saltBin);
    let finalHash = String('0000' + parseInt(finalHashBin, 2).toString(16).toUpperCase()).slice(-4);

    return finalHash;
};

/*
 * Helper Functions
 */

module.exports = {
    generateRId: generateRId,
    getHashOfPassword: getHashOfPassword
};