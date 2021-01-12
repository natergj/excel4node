'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Drawing = require('./drawing.js');
var xmlbuilder = require('xmlbuilder');
var EMU = require('../classes/emu.js');

var _require = require('uuid'),
    uuidv4 = _require.v4;

var Chart = function (_Drawing) {
    _inherits(Chart, _Drawing);

    /**
     * Element representing an Excel Picture subclass of Drawing
     * @property {String} kind Kind of picture
     */
    function Chart(opts) {
        _classCallCheck(this, Chart);

        var _this = _possibleConstructorReturn(this, (Chart.__proto__ || Object.getPrototypeOf(Chart)).call(this));

        _this.kind = 'chart';
        _this.type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
        _this._title;
        _this.chartData = opts.chartData;

        if (['oneCellAnchor', 'twoCellAnchor'].indexOf(opts.position.type) >= 0) {
            _this.anchor(opts.position.type, opts.position.from, opts.position.to);
        } else if (opts.position.type === 'absoluteAnchor') {
            _this.position(opts.position.x, opts.position.y);
        } else {
            throw new TypeError('Invalid option for anchor type. anchorType must be one of oneCellAnchor, twoCellAnchor, or absoluteAnchor');
        }
        return _this;
    }

    _createClass(Chart, [{
        key: 'addToXMLele',


        /**
         * @alias Picture.addToXMLele
         * @desc When generating Workbook output, attaches pictures to the drawings xml file
         * @func Picture.addToXMLele
         * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
         */
        value: function addToXMLele(ele) {

            var anchorEle = ele.ele('xdr:' + this.anchorType);

            // if (this.editAs !== null) {
            //     anchorEle.att('editAs', this.editAs);
            // }

            if (this.anchorType === 'absoluteAnchor') {
                anchorEle.ele('xdr:pos').att('x', this._position.x).att('y', this._position.y);
            }

            if (this.anchorType !== 'absoluteAnchor') {
                var af = this.anchorFrom;
                var afEle = anchorEle.ele('xdr:from');
                afEle.ele('xdr:col').text(af.col);
                afEle.ele('xdr:colOff').text(af.colOff);
                afEle.ele('xdr:row').text(af.row);
                afEle.ele('xdr:rowOff').text(af.rowOff);
            }

            if (this.anchorTo && this.anchorType === 'twoCellAnchor') {
                var at = this.anchorTo;
                var atEle = anchorEle.ele('xdr:to');
                atEle.ele('xdr:col').text(at.col);
                atEle.ele('xdr:colOff').text(at.colOff);
                atEle.ele('xdr:row').text(at.row);
                atEle.ele('xdr:rowOff').text(at.rowOff);
            }

            if (this.anchorType === 'oneCellAnchor' || this.anchorType === 'absoluteAnchor') {
                anchorEle.ele('xdr:ext').att('cx', this.width).att('cy', this.height);
            }

            var graphicFrame = anchorEle.ele('xdr:graphicFrame');
            graphicFrame.att("macro", "");
            var nvGraphicFramePr = graphicFrame.ele('xdr:nvGraphicFramePr');
            var cNvPrEle = nvGraphicFramePr.ele('xdr:cNvPr').att("id", this.id).att("name", "Chart " + this.id);
            // this.axExtId1 = uuidv4()
            // this.axExtId2 = uuidv4();

            // let extLst = cNvPrEle.ele("a:extLst")
            // let aext = extLst.ele("a:ext").att("uri", this.axExtId1)
            // let a16creat = aext.ele("a16:creationId").att("xmlns:a16", "http://schemas.microsoft.com/office/drawing/2014/main")
            // a16creat.att("id", this.axExtId2)
            nvGraphicFramePr.ele("xdr:cNvGraphicFramePr").ele("a:graphicFrameLocks");

            var xfrm = graphicFrame.ele("xdr:xfrm");
            xfrm.ele("a:off").att("x", 0).att("y", 0);
            xfrm.ele("a:ext").att("cx", 0).att("cy", 0);

            var graphicData = graphicFrame.ele("a:graphic").ele("a:graphicData").att("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            graphicData.ele("c:chart").att("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart").att("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").att("r:id", this.rId);

            anchorEle.ele('xdr:clientData');
        }
    }, {
        key: 'chartPartXML',
        value: function chartPartXML(xmlcx) {
            xmlcx.ele('c:date1904').att('val', 0);
            xmlcx.ele('c:lang').att('val', 'en-US');
            xmlcx.ele('c:roundedCorners').att('val', 0);
            var xchart = xmlcx.ele('c:chart');
            var title = xchart.ele('c:title');
            var title_rich = title.ele('c:tx').ele('c:rich');
            title_rich.ele('a:bodyPr');
            title_rich.ele('a:lstStyle');
            var titlep = title_rich.ele('a:p');
            titlep.ele('a:pPr').ele('a:defRPr').att('sz', 1400).ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
            var titlepr = titlep.ele('a:r');
            titlepr.ele("a:rPr").att("sz", 1400).att('lang', 'ar-SA').ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
            titlepr.ele('a:t', this.chartData.title);
            titlep.ele("a:endParaRPr").att("sz", 1400).att('lang', 'en-US').ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
            title.ele("c:overlay").att("val", 1);

            xchart.ele("c:autoTitleDeleted").att("val", 0);

            var plotArea = xchart.ele("c:plotArea");
            var layout = plotArea.ele("c:layout");
            var ml = layout.ele('c:manualLayout');
            ml.ele('c:layoutTarget').att('val', 'inner');
            ml.ele('c:xMode').att('val', 'edge');
            ml.ele('c:yMode').att('val', 'edge');
            ml.ele('c:x').att('val', 0.1);
            ml.ele('c:y').att('val', 0.1);
            ml.ele('c:w').att('val', 0.75);
            ml.ele('c:h').att('val', 0.6);

            var barChart = plotArea.ele("c:barChart");
            var catAx = plotArea.ele("c:catAx");
            var valAx = plotArea.ele("c:valAx");

            barChart.ele("c:barDir").att("val", "col");
            barChart.ele("c:grouping").att("val", "clustered");
            barChart.ele("c:varyColors").att("val", 0);
            for (var sno in this.chartData.dataSeries) {
                var sx = this.chartData.dataSeries[sno];
                // console.log(sno, sx)
                var ser = barChart.ele("c:ser");
                ser.ele("c:idx").att("val", sno);
                ser.ele("c:order").att("val", sno);
                ser.ele('c:tx').ele('c:v', sx.label);
                var ptrn = ser.ele('c:spPr').ele('a:pattFill').att('prst', sx.pattern);

                ptrn.ele('a:fgClr').ele('a:schemeClr').att('val', 'tx1');
                ptrn.ele('a:bgClr').ele('a:schemeClr').att('val', 'bg1');
                ser.ele('c:invertIfNegative').att('val', 0);

                if (sx.range) {
                    ser.ele('c:val').ele('c:numRef').ele('c:f', sx.range);
                }
                if (sx.catRange) {
                    ser.ele('c:cat').ele('c:strRef').ele('c:f', sx.catRange);
                }
                // ser.ele('c:extLst').ele('xmlns:c16', 'http://schemas.microsoft.com/office/drawing/2014/chart')
                //     .att('uri', this.axExtId1)
                //     .ele('c16:uniqueId').att('val', this.axExtId2)
            }
            var dLbl = barChart.ele('c:dLbls');
            dLbl.ele('c:showLegendKey').att('val', 0);
            dLbl.ele('c:showVal').att('val', 0);
            dLbl.ele('c:showCatName').att('val', 0);
            dLbl.ele('c:showSerName').att('val', 0);
            dLbl.ele('c:showPercent').att('val', 0);
            dLbl.ele('c:showBubbleSize').att('val', 0);

            barChart.ele('c:gapWidth').att('val', 150);
            barChart.ele('c:axId').att('val', 361690304);
            barChart.ele('c:axId').att('val', 1);

            catAx.ele('c:axId').att('val', 361690304);
            catAx.ele('c:scaling').ele('c:orientation').att('val', 'minMax');
            catAx.ele('c:delete').att('val', 0);
            catAx.ele('c:axPos').att('val', 'b');
            catAx.ele('c:numFmt').att('formatCode', 'General').att('sourceLinked', 0);
            catAx.ele('c:majorTickMark').att('val', 'out');
            catAx.ele('c:minorTickMark').att('val', 'none');
            catAx.ele('c:tickLblPos').att('val', 'nextTo');
            if (this.chartData.xlabel.font) genTxPrXML(catAx, this.chartData.xlabel.font);
            catAx.ele('c:crossAx').att('val', 1);
            catAx.ele('c:crosses').att('val', "autoZero");
            catAx.ele('c:auto').att('val', 1);
            catAx.ele('c:lblAlgn').att('val', "ctr");
            catAx.ele('c:lblOffset').att('val', 100);
            catAx.ele('c:noMultiLvlLbl').att('val', 0);
            genTitleXML(catAx, this.chartData.xlabel);

            valAx.ele('c:axId').att('val', 1);
            valAx.ele('c:scaling').ele('c:orientation').att('val', 'minMax');
            valAx.ele('c:delete').att('val', 0);
            valAx.ele('c:axPos').att('val', 'l');
            valAx.ele('c:numFmt').att('formatCode', 'General').att('sourceLinked', 1);
            valAx.ele('c:majorTickMark').att('val', 'out');
            valAx.ele('c:minorTickMark').att('val', 'none');
            valAx.ele('c:tickLblPos').att('val', 'nextTo');
            valAx.ele('c:crossAx').att('val', 361690304);
            valAx.ele('c:crosses').att('val', "autoZero");
            valAx.ele('c:crossBetween').att('val', "between");
            genTitleXML(valAx, this.chartData.ylabel);

            function genTitleXML(xmlc, txtRun) {
                var xz_xlabc = xmlc.ele('c:title');
                var xz_xlab = xz_xlabc.ele('c:tx').ele('c:rich');
                xz_xlab.ele('a:bodyPr');
                xz_xlab.ele('a:lstStyle');
                var xz_xlabp = xz_xlab.ele('a:p');
                xz_xlabp.ele('a:pPr').ele('a:defRPr').att('sz', txtRun.fontSize * 100).ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
                var xz_xlabpr = xz_xlabp.ele('a:r');
                var title_txt = xz_xlabpr.ele("a:rPr").att("sz", txtRun.fontSize * 100).att('lang', 'ar-SA');
                title_txt.ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
                if (txtRun.bold) title_txt.att('b', txtRun.bold);
                if (txtRun.italic) title_txt.att('i', txtRun.italic);
                title_txt.att('baseline', 0);

                xz_xlabpr.ele('a:t', txtRun.value);
                xz_xlabp.ele("a:endParaRPr").att("sz", txtRun.fontSize * 100).att('lang', 'en-US').ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
                xz_xlab.ele("c:overlay").att("val", 1);
            }

            function genTxPrXML(xmlc, txtRun) {
                var xz_xlab = xmlc.ele('c:txPr');

                xz_xlab.ele('a:bodyPr');
                xz_xlab.ele('a:lstStyle');
                var xz_xlabp = xz_xlab.ele('a:p');
                var fx = xz_xlabp.ele('a:pPr').ele('a:defRPr').att('sz', txtRun.fontSize * 100);
                fx.ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
                if (txtRun.bold) fx.att('b', txtRun.bold);
                if (txtRun.italic) fx.att('i', txtRun.italic);
            }

            var legend = xchart.ele("c:legend");
            legend.ele("c:legendPos").att("val", "r");
            legend.ele("c:overlay").att("val", "0");
            var txPr = legend.ele("c:txPr");txPr.ele("a:bodyPr");
            txPr.ele("a:lstStyle");
            var txPrp = txPr.ele("a:p");
            txPrp.ele('a:pPr').ele('a:defRPr').att('sz', 1200).ele('a:cs').att('typeface', 'Sakkal Majalla').att('pitchFamily', 2).att('charset', '-78');
            txPrp.ele("a:endParaRPr").att('lang', 'en-US');
            xchart.ele("c:plotVisOnly").att("val", 1);
            xchart.ele("c:dispBlanksAs").att("val", "gap");
            xchart.ele("c:showDLblsOverMax").att("val", 0);
        }
    }, {
        key: 'name',
        get: function get() {
            return this._name;
        },
        set: function set(newName) {
            this._name = newName;
        }
    }, {
        key: 'id',
        get: function get() {
            return this._id;
        },
        set: function set(id) {
            this._id = id;
        }
    }, {
        key: 'rId',
        get: function get() {
            return 'rId' + this._id;
        }
    }, {
        key: 'description',
        get: function get() {
            return this._descr !== null ? this._descr : this._name;
        },
        set: function set(desc) {
            this._descr = desc;
        }
    }, {
        key: 'title',
        get: function get() {
            return this._title !== null ? this._title : this._name;
        },
        set: function set(title) {
            this._title = title;
        }
    }, {
        key: 'extension',
        get: function get() {
            return this._extension;
        }
    }, {
        key: 'width',
        get: function get() {
            var inWidth = this._pxWidth / 96;
            var emu = new EMU(inWidth + 'in');
            return emu.value;
        }
    }, {
        key: 'height',
        get: function get() {
            var inHeight = this._pxHeight / 96;
            var emu = new EMU(inHeight + 'in');
            return emu.value;
        }
    }]);

    return Chart;
}(Drawing);

module.exports = Chart;
//# sourceMappingURL=chart.js.map