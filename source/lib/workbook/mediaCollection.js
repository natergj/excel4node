const fs = require('fs');

class MediaCollection {
    constructor() {
        this.items = [];
    }

    add(item) {
        if (typeof item === 'string') {
            fs.accessSync(item, fs.R_OK);
        }

        this.items.push(item);
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