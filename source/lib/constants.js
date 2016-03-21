const colorSchemes = [ //§20.1.6.2 clrScheme (Color Scheme)
    'dark 1',
    'light 1',
    'dark 2', 
    'light 2',
    'accent 1',
    'accent 2', 
    'accent 3', 
    'accent 4', 
    'accent 5', 
    'accent 6', 
    'hyperlink',
    'followed hyperlink'
];

const excelColors = { // subset of §20.1.10.48 ST_PresetColorVal (Preset Color Value)
    'aqua': 'FF33CCCC',
    'black': 'FF000000',
    'blue': 'FF0000FF',
    'blue-gray': 'FF666699',
    'bright green': 'FF00FF00',
    'brown': 'FF993300',
    'dark blue': 'FF000080',
    'dark green': 'FF003300',
    'dark red': 'FF800000',
    'dark teal': 'FF003366',
    'dark yellow': 'FF808000',
    'gold': 'FFFFCC00',
    'gray-25': 'FFC0C0C0',
    'gray-40': 'FF969696',
    'gray-50': 'FF808080',
    'gray-80': 'FF333333',
    'green': 'FF008000',
    'indigo': 'FF333399',
    'lavender': 'FFCC99FF',
    'light blue': 'FF3366FF',
    'light green': 'FFCCFFCC',
    'light orange': 'FFFF9900',
    'light turquoise': 'FFCCFFFF',
    'light yellow': 'FFFFFF99',
    'lime': 'FF99CC00',
    'olive green': 'FF333300',
    'orange': 'FFFF6600',
    'pale blue': 'FF99CCFF',
    'pink': 'FFFF00FF',
    'plum': 'FF993366',
    'red': 'FFFF0000',
    'rose': 'FFFF99CC',
    'sea green': 'FF339966',
    'sky blue': 'FF00CCFF',
    'tan': 'FFFFCC99',
    'teal': 'FF008080',
    'turquoise': 'FF00FFFF',
    'violet': 'FF800080',
    'white': 'FFFFFFFF',
    'yellow': 'FFFFFF00'
};

const fillPatternTypes = [ //§18.18.55 ST_PatternType (Pattern Type)
    'darkDown',
    'darkGray',
    'darkGrid',
    'darkHorizontal',
    'darkTrellis',
    'darkUp',
    'darkVerical',
    'gray0625',
    'gray125',
    'lightDown',
    'lightGray',
    'lightGrid',
    'lightHorizontal',
    'lightTrellis',
    'lightUp',
    'lightVertical',
    'mediumGray',
    'none',
    'solid'
];

const borderStyles = [ //§18.18.3 ST_BorderStyle (Border Line Styles)
    'none',
    'thin',
    'medium',
    'dashed',
    'dotted',
    'thick',
    'double',
    'hair',
    'mediumDashed',
    'dashDot',
    'mediumDashDot',
    'dashDotDot',
    'mediumDashDotDot',
    'slantDashDot'
];

const gradientFillTypes = [
    'linear',
    'path'
];

const alignmentTypes = {
    horizontal: ['center', 'centerContinuous', 'distributed', 'fill', 'general', 'justify', 'left', 'right'], // §18.18.40 ST_HorizontalAlignment (Horizontal Alignment Type)
    vertical: ['bottom', 'center', 'distributed', 'justify', 'top'] //§18.18.88 ST_VerticalAlignment (Vertical Alignment Types)
};

const readingOrders = [
    'contextDependent',
    'leftToRight', 
    'rightToLeft'
];

var defaultFont = {
    'color': 'FF000000',
    'name': 'Calibri',
    'size': '12'
};

module.exports = {colorSchemes, excelColors, fillPatternTypes, borderStyles, alignmentTypes, readingOrders, defaultFont};