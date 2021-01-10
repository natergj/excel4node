let Drawing = require('./drawing.js');
let Picture = require('./picture.js');
let headerFooterPicture = require('./headerFooterPicture.js');
let Chart = require('./chart.js')

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
        case 'headerFooterPicture':
            let xPic = new headerFooterPicture(opts);
            this.drawings.push(xPic );
            return xPic;
        case 'chart':
            let cPic = new Chart(opts);
            this.drawings.push(cPic);
            return cPic
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

module.exports = { DrawingCollection, Drawing, Picture, headerFooterPicture };
