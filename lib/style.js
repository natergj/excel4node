var lodash = require('lodash');

var style = function (wb, opts) {
    var curStyle = {};

    curStyle.wb = wb;
    curStyle.xf = new exports.cellXfs(wb, opts);

    curStyle.Font = fontFunctions(this);
    curStyle.Border = setBorder;
    curStyle.Number = numberFunctions(this);
    curStyle.Fill = fillFunctions(this);
    curStyle.Clone = cloneStyle;
    curStyle.getFont = getFont;

    function getFont() {
        return wb.styleData.fonts[curStyle.xf.fontId];
    }

    function numberFunctions() {
        var methods = {};
        methods.Format = setFormat;

        function setFormat(fmt) {
            if (!wb.styleData.numFmts[curStyle.xf.numFmtId - 164]) {
                var curFmt = new exports.numFmt(curStyle.wb);
            } else {
                var curFmt = JSON.parse(JSON.stringify(wb.styleData.numFmts[curStyle.xf.numFmtId - 164]));
            }
            curFmt.formatCode = fmt;
            var thisFmt = new exports.numFmt(curStyle.wb, curFmt);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyNumberFormat = 1;
            curXF.numFmtId = thisFmt.numFmtId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }

        return methods;
    }

    function fillFunctions() {
        var methods = {};
        methods.Color = setFillColor;
        methods.Pattern = setFillPattern;

        var curFill = JSON.parse(JSON.stringify(wb.styleData.fills[curStyle.xf.fillId]));

        function setFillColor(color) {
            curFill.fgColor = exports.cleanColor(color);
            var thisFill = new exports.fill(curStyle.wb, curFill);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFill = 1;
            curXF.fillId = thisFill.fillId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFillPattern(pattern) {
            curFill.patternType = pattern;
            var thisFill = new exports.fill(curStyle.wb, curFill);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFill = 1;
            curXF.fillId = thisFill.fillId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }

        return methods;
    }

    function fontFunctions() {
        var methods = {};
        methods.Options = setFontOptions;
        methods.Family = setFontFamily;
        methods.Bold = setFontBold;
        methods.Italics = setFontItalics;
        methods.Underline = setFontUnderline;
        methods.Size = setFontSize;
        methods.Color = setFontColor;
        methods.WrapText = setTextWrap;
        methods.Alignment = {
            Vertical: setFontAlignmentVertical,
            Horizontal: setFontAlignmentHorizontal,
            Rotation: setTextRotation
        };

        var curFont = JSON.parse(JSON.stringify(wb.styleData.fonts[curStyle.xf.fontId]));

        function setFontOptions(opts) {
            Object.keys(opts).forEach(function (o) {
                curFont[o] = opts[o];
            });
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontFamily(val) {
            curFont.name = val;
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontBold() {
            curFont.bold = true;
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontItalics() {
            curFont.italics = true;
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontUnderline() {
            curFont.underline = true;
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontSize(val) {
            curFont.sz = val;
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontColor(val) {
            curFont.color = exports.cleanColor(val);
            var thisFont = new exports.font(curStyle.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setTextWrap() {
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            if (!curXF.alignment) {
                curXF.alignment = {};
            }
            curXF.applyAlignment = 1;
            curXF.alignment.wrapText = 1;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontAlignmentVertical(val) {
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            if (!curXF.alignment) {
                curXF.alignment = {};
            }
            curXF.applyAlignment = 1;
            curXF.alignment.vertical = val;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setFontAlignmentHorizontal(val) {
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            if (!curXF.alignment) {
                curXF.alignment = {};
            }
            curXF.applyAlignment = 1;
            curXF.alignment.horizontal = val;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }
        function setTextRotation(val) {
            var curXF = JSON.parse(JSON.stringify(curStyle.xf));
            if (!curXF.alignment) {
                curXF.alignment = {};
            }
            curXF.applyAlignment = 1;
            curXF.alignment.textRotation = val;
            curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
            return curStyle;
        }

        return methods;
    }

    function setBorder(opts) {
        /*
            opts should be object in form 
            style is required, color is optional and defaults to black if not specified.
            not all ordinals are required. No board will be on side that is not specified.
            {
                left:{
                    style:'style',
                    color:'rgb'
                },
                right:{
                    style:'style',
                    color:'rgb'
                },
                top:{
                    style:'style',
                    color:'rgb'
                },
                bottom:{
                    style:'style',
                    color:'rgb'
                },
                diagonal:{
                    style:'style',
                    color:'rgb'
                }
            }
        */

        var curBorder = new exports.border(curStyle.wb, opts);
        var curXF = JSON.parse(JSON.stringify(curStyle.xf));
        curXF.applyBorder = 1;
        curXF.borderId = curBorder.borderId;
        curStyle.xf = new exports.cellXfs(curStyle.wb, curXF);
        return curStyle;
    }

    function cloneStyle() {
        var oldStyle = this;
        var newStyle = this.wb.Style();

        var doNotCloneKeys = ['xfId', 'generateXMLObj'];
        Object.keys(newStyle.xf).forEach(function (k) {
            if (doNotCloneKeys.indexOf(k) < 0) {
                newStyle.xf[k] = oldStyle.xf[k];
            }
        });
        return newStyle;
    }

    return curStyle;
};

exports.border = function (wb, opts) {

    var curBorder = this;

    opts = opts ? opts : {};
    curBorder.left = opts.left ? opts.left : null;
    curBorder.right = opts.right ? opts.right : null;
    curBorder.top = opts.top ? opts.top : null;
    curBorder.bottom = opts.bottom ? opts.bottom : null;
    curBorder.diagonal = opts.diagonal ? opts.diagonal : null;
    curBorder.generateXMLObj = genXMLObj;
    var validStyles = [
        'hair',
        'dotted',
        'dashDotDot',
        'dashDot',
        'dashed',
        'thin',
        'mediumDashDotDot',
        'slantDashDot',
        'mediumDashDot',
        'mediumDashed',
        'medium',
        'thick',
        'double'
    ];
    var ordinals = ['left', 'right', 'top', 'bottom', 'diagonal'];

    ordinals.forEach(function (o) {
        if (this && this[o]) {
            if (validStyles.indexOf(this[o].style) < 0) {
                console.log('Invalid or missing option %s specified for border style. replacing with "thin"', this[o].style);
                console.log('Valid Options: %s' + validStyles.join(','));
            }
        }
    });

    if (wb.styleData.borders.length === 0) {
        curBorder.borderId = 0;
        wb.styleData.borders.push(this);
    } else {
        var isMatched = false;
        var curborderId = 0;
        var border2 = JSON.parse(JSON.stringify(this));

        while (isMatched === false && curborderId < wb.styleData.borders.length) {
            var border1 = JSON.parse(JSON.stringify(wb.styleData.borders[curborderId]));

            border1.borderId = null;
            border2.borderId = null;
            isMatched = lodash.isEqual(border1, border2);
            if (isMatched) {
                curBorder.borderId = curborderId;
            } else {
                curborderId += 1;
            }
        }
        if (!isMatched) {
            curBorder.borderId = wb.styleData.borders.length;
            wb.styleData.borders.push(this);
        }
    }

    function genXMLObj() {
        var data = {
            border: []
        };
        ordinals.forEach(function (o) {
            if (curBorder[o]) {
                var tmpObj = {};
                tmpObj[o] = [
                    {
                        '@style': curBorder[o].style ? curBorder[o].style : 'thin'
                    },
                    {
                        'color': {
                            '@rgb': curBorder[o].color ? exports.cleanColor(curBorder[o].color) : 'FF000000'
                        }
                    }
                ];
                data.border.push(tmpObj);
            } else {
                data.border.push(o);
            }
        });
        return data;
    }
};

exports.cellXfs = function (wb, opts) {
    opts = opts ? opts : {};
    this.applyAlignment = opts.applyAlignment ? opts.applyAlignment : 0;
    this.applyBorder = opts.applyBorder ? opts.applyBorder : 0;
    this.applyNumberFormat = opts.applyNumberFormat ? opts.applyNumberFormat : 0;
    this.applyFill = opts.applyFill ? opts.applyFill : 0;
    this.applyFont = opts.applyFont ? opts.applyFont : 0;
    this.borderId = opts.borderId ? opts.borderId : 0;
    this.fillId = opts.fillId ? opts.fillId : 0;
    this.fontId = opts.fontId ? opts.fontId : 0;
    this.numFmtId = opts.numFmtId ? opts.numFmtId : 164;
    if (opts.alignment) {
        this.alignment = opts.alignment;
    }
    this.generateXMLObj = genXMLObj;

    if (wb.styleData.cellXfs.length === 0) {
        this.xfId = 0;
        wb.styleData.cellXfs.push(this);
    } else {
        var isMatched = false;
        var curXfId = 0;
        var xf2 = JSON.parse(JSON.stringify(this));

        while (isMatched === false && curXfId < wb.styleData.cellXfs.length) {
            var xf1 = JSON.parse(JSON.stringify(wb.styleData.cellXfs[curXfId]));

            xf1.xfId = null;
            xf2.xfId = null;
            isMatched = lodash.isEqual(xf1, xf2);
            if (isMatched) {
                this.xfId = curXfId;
            } else {
                curXfId += 1;
            }
        }
        if (!isMatched) {
            this.xfId = wb.styleData.cellXfs.length;
            wb.styleData.cellXfs.push(this);
        }
    }

    function genXMLObj() {
        var data = {
            xf: {
                '@applyAlignment': this.applyAlignment,
                '@applyNumberFormat': this.applyNumberFormat,
                '@applyFill': this.applyFill,
                '@applyFont': this.applyFont,
                '@applyBorder': this.applyBorder,
                '@borderId': this.borderId,
                '@fillId': this.fillId,
                '@fontId': this.fontId,
                '@numFmtId': this.numFmtId
            }
        };
        if (this.alignment) {
            data.xf.alignment = [];
            if (this.alignment.vertical) {
                data.xf.alignment.push({ '@vertical': this.alignment.vertical });
            }
            if (this.alignment.horizontal) {
                data.xf.alignment.push({ '@horizontal': this.alignment.horizontal });
            }
            if (this.alignment.wrapText) {
                data.xf.alignment.push({ '@wrapText': this.alignment.wrapText });
            }
            if (this.alignment.textRotation) {
                data.xf.alignment.push({ '@textRotation': this.alignment.textRotation });
            }
        }
        return data;
    }

    return this;
};

exports.font = function (wb, opts) {
    opts = opts ? opts : {};
    this.bold = opts.bold ? opts.bold : false;
    this.italics = opts.italics ? opts.italics : false;
    this.underline = opts.underline ? opts.underline : false;
    this.sz = opts.sz ? opts.sz : 12;
    this.color = opts.color ? exports.cleanColor(opts.color) : 'FF000000';
    this.name = opts.name ? opts.name : 'Calibri';
    if (opts.alignment) {
        this.alignment = {};
        if (opts.alignment.vertical) {
            this.alignment.vertical = opts.alignment.vertical;
            this.applyAlignment = 1;
        }
        if (opts.alignment.horizontal) {
            this.alignment.horizontal = opts.alignment.horizontal;
            this.applyAlignment = 1;
        }
        if (opts.alignment.wrapText) {
            this.alignment.wrapText = opts.alignment.wrapText;
            this.applyAlignment = 1;
        }
        if (opts.alignment.textRotation) {
            this.alignment.textRotation = opts.alignment.textRotation;
            this.applyAlignment = 1;
        }
    }
    this.generateXMLObj = genXMLObj;

    if (wb.styleData.fonts.length === 0) {
        this.fontId = 0;
        wb.styleData.fonts.push(this);
    } else {
        var isMatched = false;
        var curFontId = 0;
        var font2 = JSON.parse(JSON.stringify(this));

        while (isMatched === false && curFontId < wb.styleData.fonts.length) {
            var font1 = JSON.parse(JSON.stringify(wb.styleData.fonts[curFontId]));

            font1.fontId = null;
            font2.fontId = null;
            isMatched = lodash.isEqual(font1, font2);
            if (isMatched) {
                this.fontId = curFontId;
            } else {
                curFontId += 1;
            }
        }
        if (!isMatched) {
            this.fontId = wb.styleData.fonts.length;
            wb.styleData.fonts.push(this);
        }
    }

    function genXMLObj() {
        var data = {
            font: [
                {
                    sz: {
                        '@val': this.sz
                    }
                },
                {
                    color: {
                        '@rgb': this.color
                    }
                },
                {
                    name: {
                        '@val': this.name
                    }
                }
            ]
        };
        if (this.underline) {
            data.font.splice(0, 0, 'u');
        }
        if (this.italics) {
            data.font.splice(0, 0, 'i');
        }
        if (this.bold) {
            data.font.splice(0, 0, 'b');
        }

        if (this.alignment) {
            var alignment = {};
            if (this.alignment.vertical) {
                alignment['@vertical'] = this.alignment.vertical;
            }
            if (this.alignment.horizontal) {
                alignment['@horizontal'] = this.alignment.horizontal;
            }
            if (this.alignment.wrapText) {
                alignment['@wrapText'] = this.alignment.wrapText;
            }
            if (this.alignment.textRotation) {
                alignment['@textRotation'] = this.alignment.textRotation;
            }
            data.font.push({ alignment: alignment });
        }
        return data;
    }

    return this;
};

exports.numFmt = function (wb, opts) {
    var opts = opts ? opts : {};
    this.formatCode = opts.formatCode ? opts.formatCode : '$ #,##0.00;$ #,##0.00;-';
    this.generateXMLObj = genXMLObj;

    if (wb.styleData.numFmts.length === 0) {
        this.numFmtId = 165;
        wb.styleData.numFmts.push(this);
    } else {
        var isMatched = false;
        var curNumFmtId = 165;
        var fmt2 = JSON.parse(JSON.stringify(this));

        while (isMatched === false && curNumFmtId < wb.styleData.numFmts.length + 165) {
            var fmt1 = JSON.parse(JSON.stringify(wb.styleData.numFmts[curNumFmtId - 165]));

            fmt1.numFmtId = null;
            fmt2.numFmtId = null;
            isMatched = lodash.isEqual(fmt1, fmt2);
            if (isMatched) {
                this.numFmtId = curNumFmtId;
            } else {
                curNumFmtId += 1;
            }
        }
        if (!isMatched) {
            this.numFmtId = wb.styleData.numFmts.length + 165;
            wb.styleData.numFmts.push(this);
        }
    }

    function genXMLObj() {
        var data = {
            numFmt: [
                { '@formatCode': this.formatCode },
                { '@numFmtId': this.numFmtId }
            ]
        };
        return data;
    }
    //<numFmt numFmtId="100" formatCode="$ #,##0.00;$ #,##0.00;-" />
};

exports.fill = function (wb, opts) {
    opts = opts ? opts : {};
    this.patternType = opts.patternType ? opts.patternType : 'solid';
    if (opts.fgColor) {
        this.fgColor = exports.cleanColor(opts.fgColor);
    }
    if (opts.bgColor) {
        this.bgColor = exports.cleanColor(opts.bgColor);
    }
    this.generateXMLObj = genXMLObj;

    if (wb.styleData.fills.length === 0) {
        this.fillId = 0;
        wb.styleData.fills.push(this);
    } else {
        var isMatched = false;
        var curFillId = 0;
        var fill2 = JSON.parse(JSON.stringify(this));

        while (isMatched === false && curFillId < wb.styleData.fills.length) {
            var fill1 = JSON.parse(JSON.stringify(wb.styleData.fills[curFillId]));

            fill1.fillId = null;
            fill2.fillId = null;
            isMatched = lodash.isEqual(fill1, fill2);
            if (isMatched) {
                this.fillId = curFillId;
            } else {
                curFillId += 1;
            }
        }
        if (!isMatched) {
            this.fillId = wb.styleData.fills.length;
            wb.styleData.fills.push(this);
        }
    }

    function genXMLObj() {
        var data = { fill: { patternFill: [] } };
        data.fill.patternFill.push({ '@patternType': this.patternType });
        if (this.fgColor) {
            data.fill.patternFill.push({ fgColor: { '@rgb': this.fgColor } });
        }
        if (this.bgColor) {
            data.fill.patternFill.push({ bgColor: { '@rgb': this.bgColor } });
        }
        return data;
    }
};

exports.cleanColor = function (val) {
    // check for RGB, RGBA or Excel Color Names and return RGBA
    var excelColors = {
        'black': 'FF000000',
        'brown': 'FF993300',
        'olive green': 'FF333300',
        'dark green': 'FF003300',
        'dark teal': 'FF003366',
        'dark blue': 'FF000080',
        'indigo': 'FF333399',
        'gray-80': 'FF333333',
        'dark red': 'FF800000',
        'orange': 'FFFF6600',
        'dark yellow': 'FF808000',
        'green': 'FF008000',
        'teal': 'FF008080',
        'blue': 'FF0000FF',
        'blue-gray': 'FF666699',
        'gray-50': 'FF808080',
        'red': 'FFFF0000',
        'light orange': 'FFFF9900',
        'lime': 'FF99CC00',
        'sea green': 'FF339966',
        'aqua': 'FF33CCCC',
        'light blue': 'FF3366FF',
        'violet': 'FF800080',
        'gray-40': 'FF969696',
        'pink': 'FFFF00FF',
        'gold': 'FFFFCC00',
        'yellow': 'FFFFFF00',
        'bright green': 'FF00FF00',
        'turquoise': 'FF00FFFF',
        'sky blue': 'FF00CCFF',
        'plum': 'FF993366',
        'gray-25': 'FFC0C0C0',
        'rose': 'FFFF99CC',
        'tan': 'FFFFCC99',
        'light yellow': 'FFFFFF99',
        'light green': 'FFCCFFCC',
        'light turquoise': 'FFCCFFFF',
        'pale blue': 'FF99CCFF',
        'lavender': 'FFCC99FF',
        'white': 'FFFFFFFF'
    };

    if (Object.keys(excelColors).indexOf(val.toLowerCase()) >= 0) {
        // val was a named color that matches predefined list. return corresponding color
        return excelColors[val.toLowerCase()];
    } else if (val.length === 8 && val.substr(0, 2) === 'FF' && /^[a-fA-F0-9()]+$/.test(val)) {
        // val is already a properly formatted color string, return upper case version of itself
        return val.toUpperCase();
    } else if (val.length === 6 && /^[a-fA-F0-9()]+$/.test(val)) {
        // val is color code without Alpha, add it and return
        return 'FF' + val.toUpperCase();
    } else if (val.length === 7 && val.substr(0, 1) === '#' && /^[a-fA-F0-9()]+$/.test(val.substr(1))) {
        // val was sent as html style hex code, remove # and add alpha
        return 'FF' + val.substr(1).toUpperCase();
    } else if (val.length === 9 && val.substr(0, 1) === '#' && /^[a-fA-F0-9()]+$/.test(val.substr(1))) {
        // val sent as html style hex code with alpha. revese alpha position and return
        return val.substr(7).toUpperCase() + val.substr(1, 6).toUpperCase();
    } else {
        // I don't know what this is, return valid color and console.log error
        console.log('%s is an invalid color option. changing to white', val);
        console.log('valid color options are html style hex codes or these colors by name: %s', Object.keys(excelColors).join(', '));
        return 'FFFFFFFF';
    }
};

exports.Style = function (opts) {
    var opts = opts ? opts : {};
    if (this.styleData.fonts.length === 0) {
        var defaultFont = new exports.font(this);
    }
    if (this.styleData.fills.length === 0) {
        var defaultFill = new exports.fill(this, { patternType: 'none' });
        var defaultFill2 = new exports.fill(this, { patternType: 'gray125' });
    }
    if (this.styleData.borders.length === 0) {
        var defaultBorder = new exports.border(this);
    }
    if (this.styleData.cellXfs.length === 0) {
        var defaultXF = new exports.cellXfs(this);
    }

    var newStyle = new style(this, opts);
    this.styleData.cellXfs[newStyle.xf.xfId] = newStyle.xf;

    return newStyle;
};

exports.getStyleById = function (wb, id) {
    var xf = undefined;
    wb.styleData.cellXfs.forEach(function (s) {
        if (s.xfId === id) {
            xf = s;
        }
    });
    return xf;
};

exports.getBorderById = function (wb, id) {
    return wb.styleData.borders[id];
};
