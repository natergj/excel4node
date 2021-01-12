'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Hyperlink = function () {
    //§18.3.1.47 hyperlink (Hyperlink)
    function Hyperlink(opts) {
        _classCallCheck(this, Hyperlink);

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

    _createClass(Hyperlink, [{
        key: 'addToXMLEle',
        value: function addToXMLEle(ele) {
            var thisEle = ele.ele('hyperlink');
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
    }, {
        key: 'rId',
        get: function get() {
            return 'rId' + this.id;
        }
    }]);

    return Hyperlink;
}();

var HyperlinkCollection = function () {
    //§18.3.1.48 hyperlinks (Hyperlinks)
    function HyperlinkCollection() {
        _classCallCheck(this, HyperlinkCollection);

        this.links = [];
    }

    _createClass(HyperlinkCollection, [{
        key: 'add',
        value: function add(opts) {
            var thisLink = new Hyperlink(opts);
            thisLink.id = this.links.length + 1;
            this.links.push(thisLink);
            return thisLink;
        }
    }, {
        key: 'addToXMLele',
        value: function addToXMLele(ele) {
            if (this.length > 0) {
                var linksEle = ele.ele('hyperlinks');
                this.links.forEach(function (l) {
                    l.addToXMLEle(linksEle);
                });
                linksEle.up();
            }
        }
    }, {
        key: 'length',
        get: function get() {
            return this.links.length;
        }
    }]);

    return HyperlinkCollection;
}();

module.exports = { HyperlinkCollection: HyperlinkCollection, Hyperlink: Hyperlink };
//# sourceMappingURL=hyperlink.js.map