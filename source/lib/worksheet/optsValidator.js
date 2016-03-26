const types = require('../constants/index.js');

const optsTypes = {
    'margins': {
        'bottom': null,
        'footer': null,
        'header': null,
        'left': null,
        'right': null,
        'top': null
    },
    'printOptions': {
        'centerHorizontal': null,
        'centerVertical': null,
        'printGridLines': null,
        'printHeadings': null
    
    },
    'pageSetup': {
        'blackAndWhite': null,
        'copies': null,
        'draft': null,
        'firstPageNumber': null,
        'fitToHeight': null,
        'fitToWidth': null,
        'horizontalDpi': null,
        'orientation': null,
        'paperHeight': null,
        'paperSize': 'PAPER_SIZE',
        'paperWidth': null,
        'useFirstPageNumber': null,
        'usePrinterDefaults': null,
        'verticalDpi': null
    },
    'sheetView': {
        'pane': {
            'activePane': null,
            'state': null,
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