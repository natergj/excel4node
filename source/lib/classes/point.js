class Point {    
    /** 
     * An XY coordinate point on the Worksheet with 0.0 being top left corner
     * @class Point
     * @property {Number} x X coordinate of Point
     * @property {Number} y Y coordinate of Point
     * @returns {Point} Excel Point
     */
    constructor(x, y) {    
        this.x = x;
        this.y = y;
    }
}

module.exports = Point;