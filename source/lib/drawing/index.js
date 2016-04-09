let Drawing = require('./drawing.js');
let Picture = require('./picture.js');

class DrawingCollection {
    constructor() {
        this.drawings = [];
    }

    get length() {
        return this.drawings.length;
    }

    add(opts) {
        switch (opts.type) {
        case 'picture':
            let newPic = new Picture(opts);
            this.drawings.push(newPic);
            return newPic;

        default:
            throw new TypeError('this option is not yet supported');
        }
    }

    get isEmpty() {
        if (this.drawings.length === 0) {
            return true;
        } else {
            return false;
        }
    }
}

module.exports = { DrawingCollection, Drawing, Picture };
