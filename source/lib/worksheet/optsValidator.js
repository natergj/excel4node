const types = require('../constants/index.js');

const optsTypes = {
    'margins': {
        'bottom': 'Float',
        'footer': 'Float',
        'header': 'Float',
        'left': 'Float',
        'right': 'Float',
        'top': 'Float'
    },
    'printOptions': {
        'centerHorizontal': 'Boolean',
        'centerVertical': 'Boolean',
        'printGridLines': 'Boolean',
        'printHeadings': 'Boolean'
    
    },
    'pageSetup': {
        'blackAndWhite': 'Boolean',
        'cellComments': 'CELL_COMMENTS',
        'copies': 'Integer',
        'draft': 'Boolean',
        'errors': 'PRINT_ERROR',
        'firstPageNumber': 'Boolean',
        'fitToHeight': 'Integer',
        'fitToWidth': 'Integer',
        'horizontalDpi': 'Integer',
        'orientation': 'ORIENTATION',
        'pageOrder': 'PAGE_ORDER',
        'paperHeight': 'POSITIVE_UNIVERSAL_MEASURE',
        'paperSize': 'PAPER_SIZE',
        'paperWidth': 'POSITIVE_UNIVERSAL_MEASURE',
        'scale': 'Integer',
        'useFirstPageNumber': 'Boolean',
        'usePrinterDefaults': 'Boolean',
        'verticalDpi': 'Integer'
    },
    'headerFooter': {
        'evenFooter': 'String',
        'evenHeader': 'String',
        'firstFooter': 'String',
        'firstHeader': 'String',
        'oddFooter': 'String',
        'oddHeader': 'String',
        'alignWithMargins': 'Boolean',
        'differentFirst': 'Boolean',
        'differentOddEven': 'Boolean',
        'scaleWithDoc': 'Boolean'
    },
    'sheetView': {
        'pane': {
            'activePane': 'PANE',
            'state': 'PANE_STATE',
            'topLeftCell': null,
            'xSplit': null,
            'ySplit': null
        },
        'tabSelected': null,
        'workbookViewId': null,
        'rightToLeft': null,
        'zoomScale': null,
        'zoomScaleNormal': null,
        'zoomScalePageLayoutView': null
    },
    'sheetFormat': {
        'baseColWidth': null,
        'customHeight': null,
        'defaultColWidth': null,
        'defaultRowHeight': null,
        'outlineLevelCol': null,
        'outlineLevelRow': null,
        'thickBottom': null,
        'thickTop': null,
        'zeroHeight': null
    },
    'sheetProtection': {
        'autoFilter': null,
        'deleteColumns': null,
        'deleteRow': null,
        'formatCells': null,
        'formatColumns': null,
        'formatRows': null,
        'hashValue': null,
        'insertColumns': null,
        'insertHyperlinks': null,
        'insertRows': null,
        'objects': null,
        'password': null,
        'pivotTables': null,
        'scenarios': null,
        'selectLockedCells': null,
        'selectUnlockedCell': null,
        'sheet': null,
        'sort': null
    },
    'outline': {
        'summaryBelow': null
    },
    'autoFilter': {
        'startRow': null,
        'endRow': null,
        'startCol': null,
        'endCol': null,
        'filters': null
    }
};

let getObjItem = (obj, key) => {
    let returnObj = obj;
    let levels = key.split('.');

    while (levels.length > 0) {
        let thisLevelKey = levels.shift();
        try {
            returnObj = returnObj[thisLevelKey];
        } catch (e) {
            //returnObj = undefined;
        }
    }
    return returnObj;
};

let validator = function (key, val, type) {
    switch (type) {

    case 'PAPER_SIZE': 
        let sizes = Object.keys(types.PAPER_SIZE);
        if (sizes.indexOf(val) < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + sizes.join(', '));
        }
        break;

    case 'PAGE_ORDER':
        let orders = ['downThenOver', 'overThenDown'];
        if (orders.indexOf(val) < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + orders.join(', '));
        }
        break;

    case 'ORIENTATION':
        let orientations = ['default', 'portrait', 'landscape'];
        if (orientations.indexOf(val) < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + orientations.join(', '));
        }
        break;

    case 'POSITIVE_UNIVERSAL_MEASURE': 
        let re = new RegExp('[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)');
        if (re.test(val) !== true) {
            throw new TypeError('Invalid value for ' + key + '. Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. 10.5cm');
        }
        break;

    case 'CELL_COMMENTS':
        let comments = ['none', 'asDisplayed', 'atEnd'];
        if (comments.indexOf('val') < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + comments.join(', '));
        }
        break;

    case 'PRINT_ERROR': 
        let printErrors = ['displayed', 'blank', 'dash', 'NA'];
        if (printErrors.indexOf(val) < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + printErrors.join(', '));
        }
        break;

    case 'PANE':
        let panes = ['bottomLeft', 'bottomRight', 'topLeft', 'topRight'];
        if (panes.indexOf(val) < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + panes.join(', '));
        }
        break;

    case 'PANE_STATE':
        let paneStates = ['split', 'frozen', 'frozenSplit'];
        if (paneStates.indexOf(val) < 0) {
            throw new TypeError('Invalid value for ' + key + '. Value must be one of ' + paneStates.join(', '));
        }
        break;

    case 'Boolean':
        if ([true, false, 1, 0].indexOf(val) < 0) {
            throw new TypeError(key + ' expects value of true, false, 1 or 0');
        }
        break;

    case 'Float': 
        if (parseFloat(val) !== val) {
            throw new TypeError(key + ' expects value as a Float number');
        }
        break;

    case 'Integer':
        if (parseInt(val) !== val) {
            throw new TypeError(key + ' expects value as an Integer');
        }
        break;

    case 'String': 
        if (typeof val !== 'string') {
            throw new TypeError(key + ' expects value as a String');
        }
        break;

    default:
        break;
    }
};

let traverse = function (o, keyParts, func) {
    for (let i in o) {
        let thisKeyParts = keyParts.concat(i);
        let thisKey = thisKeyParts.join('.');
        let thisType = getObjItem(optsTypes, thisKey);

        if (typeof thisType === 'string') {
            let thisItem = o[i];
            func(thisKey, thisItem, thisType); 
        }
        if (o[i] !== null && typeof o[i] === 'object') {
            traverse(o[i], thisKeyParts, func);
        }
    }
};

module.exports = (opts) => {
    traverse(opts, [], validator);
};