# excel4node
A full featured xlsx file generation library allowing for the creation of advanced Excel files.

excel4node conforms to the ECMA-376 OOXML specification 2nd edition   

REFERENCES   
[OpenXML White Paper](http://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf)   
[ECMA-376 Website](http://www.ecma-international.org/publications/standards/Ecma-376.htm)   
[OpenOffice Excel File Format Reference](http://www.openoffice.org/sc/excelfileformat.pdf)   
[OpenOffice Anatomy of OOXML explanation](http://officeopenxml.com/anatomyofOOXML-xlsx.php)   
[MS-XSLX spec (pdf)] (http://download.microsoft.com/download/D/3/3/D334A189-E51B-47FF-B0E8-C0479AFB0E3C/%5BMS-XLSX%5D.pdf)    

Code references specifications sections from ECMA-376 2nd edition doc   
ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference.pdf   
found in ECMA-376 2nd edition Part 1 download at [http://www.ecma-international.org/publications/standards/Ecma-376.htm](http://www.ecma-international.org/publications/standards/Ecma-376.htm)   

### Basic Usage
```javascript
// Require library
var xl = require('excel4node');

// Create a new instance of a WorkBook class
var wb = new xl.WorkBook();

// Add Workseets to the workbook
var ws = wb.addWorksheet('Sheet 1');
var ws2 = wb.addWorksheet('Sheet 2');

// Create a reusable style
var style = wb.Style({
	font: {
		color: '#FF0800',
		size: 12
	},
	numberFormat: '$#,##0.00; ($#,##0.00); -'
});

// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1,1).number(100).style(style);

// Set value of cell B1 to 300 as a number type styled with paramaters of style
ws.cell(1,2).number(200).style(style);

// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1,3).formula('A1 + B1').style(style);

// Set value of cell A2 to 'string' styled with paramaters of style
ws.cell(2,1).string('string').style(style);

// Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
ws.cell(3,1).bool(true).style(style).style({font: {size: 14}});

wb.write('Excel.xlsx');
```
## excelnode
excel4node comes with some generic functions and types

xl.getExcelRowCol(cellRef)   
Accepts cell reference (i.e. 'A1') and returns object with corresponding row and column

```javascript 
xl.getExcelRowCol('B5);
// returns { row: 5, col: 2} 
```

xl.getExcelAlpha(column)   
Accepts column as integer and returns corresponding column reference as alpha

```javascript 
xl.getExcelAlpha(10);
// returns 'J'
```

xl.getExcelCellRef(row, column)   
Accepts row and column as integers and returns Excel cell reference

```javascript 
xl.getExcelCellRef(5, 3);
// returns 'C5'
```

xl.getExcelTS(date)   
Accepts Date object and returns an Excel timestamp

```javascript 
var newDate = new Date('2015-01-01T00:00:00.0000Z');
xl.getExcelTS(newDate);
// Returns 42004.791666666664
```

xl.PAPER_SIZE


## WorkBook
An instance of the WorkBook class contains all data and parameters for the Excel WorkBook.

#### Constructor
WorkBook constructor accepts a configuration object.

```javascript
var xl = require('excel4node');
var wb = new xl.WorkBook({
    jszip: {
        compression: 'DEFLATE'
    },
    defaultFont: {
        size: 12,
        name: 'Calibri',
        color: 'FFFFFFFF'
    }
});
```

#### Methods   
wb.addWorksheet(name, options);   
Adds a new WorkSheet to the WorkBook   
Accepts name of new WorkSheet and options object (see WorkSheet section)   
Returns a WorkSheet instance

wb.setSelectedTab(id);   
Sets which tab will be selected when the WorkBook is opened   
Accepts Sheet ID (1-indexed sheet in order that sheets were added)

wb.createStyle(opts);  
Creates a new Style instance   
Accepts Style configuration object (see Style section)
Returns a new Style instance   


## WorkSheet
An instance of the WorkSheet class contains all information specific to that worksheet

#### Contstructor
WorkSheet contructor is called via WorkBook class and accepts a name and configuration object

```javascript
var xl = require('excel4node');
var wb = new xl.WorkBook();

var options = {
	margins: {
		left: 1.5,
		right: 1.5
	}
};

var ws = wb.addWorksheet(options);
```

Full WorkSheet options

```
{
    'margins': { // Accepts a Double in Inches
        'bottom': Double,
        'footer': Double,
        'header': Double,
        'left': Double,
        'right': Double,
        'top': Double
    },
    'printOptions': {
        'centerHorizontal': Boolean,
        'centerVertical': Boolean,
        'printGridLines': Boolean,
        'printHeadings': Boolean
    
    },
    'headerFooter': { // Set Header and Footer strings and options. 
        'evenFooter': String,
        'evenHeader': String,
        'firstFooter': String,
        'firstHeader': String,
        'oddFooter': String,
        'oddHeader': String,
        'alignWithMargins': Boolean,
        'differentFirst': Boolean,
        'differentOddEven': Boolean,
        'scaleWithDoc': Boolean
    },
    'pageSetup': {
        'blackAndWhite': Boolean,
        'cellComments': xl.ST_CellComments,
        'copies': null,
        'draft': null,
        'errors': xl.ST_PrintError,
        'firstPageNumber': null,
        'fitToHeight': null,
        'fitToWidth': null,
        'horizontalDpi': null,
        'orientation': null,
        'pageOrder': null,
        'paperHeight': null,
        'paperSize': null,
        'paperWidth': null,
        'scale': null,
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
        'tabSelected': 0,
        'workbookViewId': 0,
        'rightToLeft': 0,
        'zoomScale': 100,
        'zoomScaleNormal': 100,
        'zoomScalePageLayoutView': 100
    },
    'sheetFormat': {
        'baseColWidth': 10,
        'customHeight': null,
        'defaultColWidth': null,
        'defaultRowHeight': 16,
        'outlineLevelCol': null,
        'outlineLevelRow': null,
        'thickBottom': null,
        'thickTop': null,
        'zeroHeight': null
    },
    'sheetProtection': {                 // same as "Protect Sheet" in Review tab of Excel 
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
        'summaryBelow': false
    },
    'autoFilter': {
        'startRow': null,
        'endRow': null,
        'startCol': null,
        'endCol': null,
        'filters': []
    }
}
```

Notes:   
headerFooter strings accept [Dynamic Formatting Strings](https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/extensions/XSSFHeaderFooter.html). i.e. '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N'   


