const fs = require('fs');

class MediaCollection {
    constructor() {
        this.items = [];
    }

    add(filePath) {
        fs.accessSync(filePath, fs.R_OK);
        this.items.push(filePath);
        return this.items.length;
    }

    get isEmpty() {
        if (this.items.length === 0) {
            return true;
        } else {
            return false;
        }
    }
}

module.exports = MediaCollection;