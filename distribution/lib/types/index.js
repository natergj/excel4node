'use strict';

var fs = require('fs');
var path = require('path');
var dirItems = fs.readdirSync(__dirname);

dirItems.forEach(function (i) {
    if (i !== 'index.js' && i.substr(i.length - 3, 3) === '.js') {
        module.exports[i.substr(0, i.length - 3)] = require(path.resolve(__dirname, i));
    }
});
//# sourceMappingURL=index.js.map