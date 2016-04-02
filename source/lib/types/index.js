let fs = require('fs');
let path = require('path');
let dirItems = fs.readdirSync(__dirname);

dirItems.forEach((i) => {
    if (i !== 'index.js' && i.substr(i.length - 3, 3) === '.js') {
        module.exports[i.substr(0, i.length - 3)] = require(path.resolve(__dirname, i));
    }
});