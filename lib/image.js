var fs = require('fs'),
    mime = require('mime'),
    imgsz = require('image-size');

var drawing = function (imgURI, anchor) {
    anchor = anchor || module.exports.Image.ONE_CELL;

    var xml = {},
        key = 'xdr:' + anchor;
    xml[key] = {};

    var  d = {
        props: {
            image: imgURI,
            index: 0,
            imageId: -1,
            mimeType: '',
            extension: '',
            width: 0,
            height: 0,
            dpi: 96
        },
        xml: xml,

        Position: function (r, c, offY, offX, to_offY, to_offX) {
            var offsetX = offX ? offX : 0;
            var offsetY = offY ? offY : 0;
            var xdr;

            if (anchor === module.exports.Image.ABSOLUTE) {
                xdr = d.xml[key]['xdr:pos'] = {};
                xdr['@y'] = r * 9525 * (96 / d.props.dpi); //r - represents Y pixels
                xdr['@x'] = c * 9525 * (96 / d.props.dpi); //c - represents X pixels
            } else {
                xdr = d.xml[key]['xdr:from'] = {};

                xdr['xdr:col'] = c - 1;
                xdr['xdr:colOff'] = offsetX * 9525 * (96 / d.props.dpi);
                xdr['xdr:row'] = r - 1;
                xdr['xdr:rowOff'] = offsetY * 9525 * (96 / d.props.dpi);

                if (anchor === module.exports.Image.TWO_CELL) {
                    xdr = d.xml[key]['xdr:to'] = {};
                    offsetX = to_offY ? to_offY : 0;
                    offsetY = to_offX ? to_offX : 0;

                    xdr['xdr:col'] = offX || c; //offX - represents end column index
                    xdr['xdr:colOff'] = 0;
                    xdr['xdr:row'] = offY || r; //offY - represents end row index
                    xdr['xdr:rowOff'] = 0;
                }
            }
            return d;
        },
        Properties: function (props) {
            Object.keys(props).forEach(function (k) {
                d.props[k] = props[k];
            });
            return d;
        },
        ImageProperties: function (index, imageId, name, descr) {
            var pic = d.xml[key]['xdr:pic'] = {
                    'xdr:nvPicPr': {
                        'xdr:cNvPr': {},
                        'xdr:cNvPicPr': {
                            'a:picLocks': {
                                '@noChangeArrowheads': 1,
                                '@noChangeAspect': 1
                            }
                        }
                    },
                    'xdr:blipFill': {
                        'a:blip': {
                            '@r:embed': 'rId' + imageId,
                            '@xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                        },
                        'a:srcRect': {},
                        'a:stretch': {
                            'a:fillRect': {}
                        }
                    },
                    'xdr:spPr': {
                        '@bwMode': 'auto',
                        'a:xfrm': {},
                        'a:prstGeom': {
                            '@prst': 'rect'
                        }
                    }
                },
                cNvPr = pic['xdr:nvPicPr']['xdr:cNvPr']; 
            d.xml[key]['xdr:clientData'] = {};

            cNvPr['@descr'] = descr || 'image';
            cNvPr['@id'] = index;
            cNvPr['@name'] = name || 'Picture';
            
            d.props.imageId = imageId;
            d.props.index = index;
            return d;
        },
        Size: function (width, height) {
            if (anchor !== module.exports.Image.TWO_CELL) {
                width = isNaN(parseFloat(width)) ? 100 : width;
                height = isNaN(parseFloat(height)) ? 15 : height;
                d.xml[key]['xdr:ext'] = {
                    '@cx': width * 9525 * (96 / d.props.dpi),
                    '@cy': height * 9525 * (96 / d.props.dpi)
                }
            }
            d.props.width = width;
            d.props.height = height;
            return d;
        }
    };

    return d;
};

module.exports.Image = function (imgURI, anchor) {

    var wb = this.wb.workbook;
    var wSs = this.wb.worksheets;
    var ws = this;

    // add entry to [Content_Types].xml
    var mimeType = mime.lookup(imgURI);
    var extension = mimeType.split('/')[1];

    var contentTypeAdded = false;
    wb.Content_Types.Types.forEach(function (t) {
        if (t['Default']) {
            if (t['Default']['@ContentType'] === mimeType) {
                contentTypeAdded = true;
            }
        }
    });
    if (!contentTypeAdded) {
        wb.Content_Types.Types.push({
            'Default': {
                '@ContentType': mimeType,
                '@Extension': extension
            }
        });
    }

    // create drawingn.xml file
    // create drawingn.xml.rels file
    if (!ws.drawings) {
        ws.drawings = {
            'rels': {
                'Relationships': [
                    {
                        '@xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
                    }
                ]
            },
            'xml': {
                'xdr:wsDr': [
                    {
                        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        '@xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
                    }
                ]
            },
            drawings: []
        };
    }
    if (!ws.rels) {
        ws.rels = {
            'Relationships': [
                {
                    '@xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
                }
            ]
        };
    }

    var d = new drawing(imgURI, anchor);


    d.Properties({
        'mimeType': mimeType,
        'extension': extension
    });

    var dim = imgsz(imgURI);
    d.Position(1, 1);
    d.Size({
        'width': dim.width,
        'height': dim.height
    });

    var imgID = 0;
    var fontInWS = null;
    wSs.forEach(function (s) {
        if (s.drawings) {
            imgID += s.drawings.drawings.length;
            for (var i in s.drawings.drawings) {
                var d = s.drawings.drawings[i];
                if (d.props && d.props.imageId !== -1 &&  d.props.image === imgURI) {
                    imgID = d.props.imageId;
                    fontInWS = s;
                }
            }
        }
        return fontInWS === null;
    });
    ws.drawings.drawings.push(d);
    d.ImageProperties(ws.drawings.drawings.length, imgID);

    if (fontInWS !== ws) {
        ws.drawings.rels.Relationships.push({
            'Relationship': {
                '@Id': 'rId' + imgID,
                '@Target': '../media/image' + imgID + '.' + extension,
                '@Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            }
        });
    }

    var relExists = false;
    ws.rels['Relationships'].forEach(function (r) {
        if (r['Relationship']) {
            if (r['Relationship']['@Id'] === 'rId1') {
                relExists = true;
            }
        }
    });
    if (!relExists) {
        ws.rels['Relationships'].push({
            'Relationship': {
                '@Id': 'rId1',
                '@Target': '../drawings/drawing' + ws.sheetId + '.xml',
                '@Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
            }
        });
    }
    ws.sheet.drawing = {
        '@r:id': 'rId1'
    };

    return d;
};

module.exports.Image.ONE_CELL = 'oneCellAnchor';
module.exports.Image.TWO_CELL = 'twoCellAnchor';
module.exports.Image.ABSOLUTE = 'absoluteAnchor';