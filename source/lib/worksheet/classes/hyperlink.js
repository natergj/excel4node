
class Hyperlink { //§18.3.1.47 hyperlink (Hyperlink)
    constructor(opts) {
        opts = opts ? opts : {};
        
        if (opts.ref === undefined) {
            throw new TypeError('ref is a required option when creating a hyperlink');
        }
        this.ref = opts.ref;

        if (opts.display !== undefined) {
            this.display = opts.display;
        } else {
            this.display = opts.location;
        }
        if (opts.location !== undefined) {
            this.location = opts.location;
        }
        if (opts.tooltip !== undefined) {
            this.tooltip = opts.tooltip;
        } else {
            this.tooltip = opts.location;
        }
        this.id;
    }

    get rId() {
        return 'rId' + this.id;
    }

    addToXMLEle(ele) {
        let thisEle = ele.ele('hyperlink');
        thisEle.att('ref', this.ref);
        thisEle.att('r:id', this.rId);
        if (this.display !== undefined) {
            thisEle.att('display', this.display);
        }
        if (this.location !== undefined) {
            thisEle.att('address', this.location);
        }
        if (this.tooltip !== undefined) {
            thisEle.att('tooltip', this.tooltip);
        }
        thisEle.up();  
    }
}

class HyperlinkCollection { //§18.3.1.48 hyperlinks (Hyperlinks)
    constructor() {
        this.links = [];
    }

    get length() {
        return this.links.length;
    }

    add(opts) {
        let thisLink = new Hyperlink(opts);
        thisLink.id = this.links.length + 1;
        this.links.push(thisLink);
        return thisLink;
    }

    addToXMLele(ele) {
        if (this.length > 0) {
            let linksEle = ele.ele('hyperlinks');
            this.links.forEach((l) => {
                l.addToXMLEle(linksEle);
            });
            linksEle.up();
        }
    }
}

module.exports = { HyperlinkCollection, Hyperlink };