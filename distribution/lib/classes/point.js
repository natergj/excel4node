"use strict";

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Point =
/** 
 * An XY coordinate point on the Worksheet with 0.0 being top left corner
 * @class Point
 * @property {Number} x X coordinate of Point
 * @property {Number} y Y coordinate of Point
 * @returns {Point} Excel Point
 */
function Point(x, y) {
    _classCallCheck(this, Point);

    this.x = x;
    this.y = y;
};

module.exports = Point;
//# sourceMappingURL=point.js.map