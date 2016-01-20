# excel4node

An OOXML (xlsx) generator that supports formatting options.


## Installation

    npm install excel4node


## Usage Example

```javascript
var xl = require('excel4node');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('My Worksheet');

var myCell = ws.Cell(1, 1);
myCell.String('Test Value');

wb.write('MyExcel.xlsx');
```


## Sample

A sample.js script is provided in the code. Running this will output a sample excel workbook named Excel.xlsx

    node sample.js
    open Excel.xlsx


## Workbook

A Workbook represents an Excel document.

```javascript
var xl = require('excel4node');
var wb2 = new xl.WorkBook({      // optional params object
    jszip: {
        compression: 'DEFLATE'   // change the zip compression method
    },
    fileSharing: {               // equates to "password to modify option"
        password: 'Password',    // This does not encrypt the workbook,
        userName: 'John Doe'     // and users can still open the workbook as read-only.
    }
});
```
Set a default font for the workbook

```javascript  
var xl = require('excel4node');
var wb = new xl.WorkBook();
wb.updateDefaultFont({
	size : 12,
	bold : false,
	italics : false,
	underline : true,
	color : 'FF000000',
	font : 'Calibri'
});
```


## Worksheet

A Worksheet represents a tab within an excel document.

```javascript
var ws = wb.WorkSheet('My Worksheet', {
    margins: {                         // page margins
        left: 0.75,
        right: 0.75,
        top: 1.0,
        bottom: 1.0,
        footer: 0.5,
        header: 0.5
    },
    printOptions: {                    // page print options
        centerHorizontal: true,
        centerVertical: false
    },
    view: {                            // page zoom
        zoom: 100
    },
    outline: {
        summaryBelow: true
    },
    fitToPage: {
        fitToHeight: 100,
        orientation: 'landscape',
    },
    sheetProtection: {                 // same as "Protect Sheet" in Review tab of Excel
        autoFilter: false,
        deleteColumns: false,
        deleteRow : false,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertColumns: false,
        insertHyperlinks: false,
        insertRows: false,
        objects: false,
        password: 'Password',
        pivotTables: false,
        scenarios: false,
        sheet: true,
        sort: false
    }
});

// https://support.office.com/en-US/article/Set-a-specific-print-area-BEEBCEB7-0D43-4E07-8895-5AFE0AEDFB32
ws.printArea({
	rows: {
		begin: 1,
		end: 2
	},
	columns: {
		begin: 1,
		end: 3
	}
});

// https://support.office.com/en-us/article/Repeat-specific-rows-or-columns-on-every-printed-page-0d6dac43-7ee7-4f34-8b08-ffcc8b022409
ws.printTitles({
	rows: {
	    begin: 1,
	    end: 2
	},
	columns: {
	    begin: 1,
	    end: 3
	}
});

// Uses https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/extensions/XSSFHeaderFooter.html
ws.headerFooter({
	firstHeader: '&LFirst Page of Report Header&R&D',
	firstFooter: '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N',
	evenHeader: '&LReport Header&R&D',
	evenFooter: '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N',
	oddHeader: '&LReport Header&R&D',
	oddFooter: '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N'
});
```

The `sheetProtection` options are the same as the "Protect Sheet" functions in the Review tab of Excel to prevent certain user editing.  Setting a value to true means that that particular function is protected and the user will not be able to do that thing. All options are false by default except for 'sheet' which defaults to true if the sheetProtection attribute is set in the worksheet options, but false if it is not.


### Worksheet Validations

Optionally, you can set validations for a WorkSheet.

```javascript
ws.setValidation({
    type: 'list',
    allowBlank: 1,
    showInputMessage: 1,
    showErrorMessage: 1,
    sqref: 'X2:X10',
    formulas: [
        'value1,value2'
    ]
});

ws.setValidation({
    type: 'list',
    allowBlank: 1,
    sqref: 'B2:B10',
    formulas: [
        '=sheet2!$A$1:$A$2'
    ]
});
```


## Rows & Columns

Set dimensions of rows or columns:

```javascript
ws.Row(1).Height(30);
ws.Column(1).Width(100);
```


Freeze rows and columns:

```javascript
ws.Column(3).Freeze();   // freeze the first two columns (everything prior to the specified column)
ws.Column(3).Freeze(8);  // freeze the first two columns and scroll to the 8th column
ws.Row(3).Freeze();      // freeze the first two rows (everything prior to the specified row)
ws.Row(3).Freeze(8);     // freeze the first two rows and scroll to the 8th row.
```

See also "Series with frozen Row" tab in sample output workbook.


Set a row to be a filter row:

```javascript
ws.Row(1).Filter();    // no arguments passed will add filter to any populated columns
ws.Row(1).Filter(1,8); // optional start and end columns
```

See also "Departmental Spending Report" tab in sample output workbook.


Hide a specific oow or column:

```javascript
ws.Row(2).Hide();
ws.Column(2).Hide();
```


Set groupings on rows and optionally collapse them:

```
ws.Row(rowNum).Group(level,isCollapsed)
ws.Row(1).Group(1,true)
```

See also "Groupings Summary Top" and "Groupings Summary Bottom" in sample output.


## Cells

Represents a cell within a worksheet.

Cell can take 6 data types: `String`, `Complex`, `Number`, `Formula`, `Date`, `Link`, and `Bool`.

Cell takes two required arguments: row, col and 3 optional arguments

Add a cell to a WorkSheet with some data:

```javascript
ws.Cell({startRow}, {startCol} [, {endRow}, {endCol}, {isMerged}]);
ws.Cell(1, 1).String('My String');
var complexString = [
	{
		bold:true,
		underline: true,
		italic: true,
		color: 'FF0000',
		size: 18,
		value: 'Hello'
	}, 
	' World!', 
	{
		color: '000000'
	},  
	' All', 
	' these', 
	' strings', 
	' are', 
	' black', 
	{
		color: '0000FF',
		value: ', but'
	},  
	' now', 
	' are', 
	' blue' 
];
ws.Cell(1, 2).Complex(complexString);
ws.Cell(2, 1).Number(5);
ws.Cell(2, 2).Number(10);
ws.Cell(2, 3).Formula('A2+B2');
ws.Cell(2, 4).Formula('A2/C2');
ws.Cell(2, 5).Date(new Date());
ws.Cell(2, 6).Link('http://google.com');
ws.Cell(2, 6).Link('http://google.com', 'Link name');
ws.Cell(2, 7).Bool(true);
```


## Styles

Style objects can be applied to cells:

```javascript
var myStyle = wb.Style();
myStyle.Font.Bold();
myStyle.Font.Italics();
myStyle.Font.Underline();
myStyle.Font.Family('Times New Roman');
myStyle.Font.Color('FF0000');
myStyle.Font.Size(16);
myStyle.Font.Alignment.Vertical('top');
myStyle.Font.Alignment.Horizontal('left');
myStyle.Font.Alignment.Rotation('90');
myStyle.Font.WrapText(true);

var myStyle2 = wb.Style();
myStyle2.Font.Size(14);
myStyle2.Number.Format('$#,##0.00;($#,##0.00);-');

var myStyle3 = wb.Style();
myStyle3.Font.Size(14);
myStyle3.Number.Format('##%');
myStyle3.Fill.Pattern('solid');
mystyle3.Fill.Color('CCCCCC');
myStyle3.Border({
    top: {
        style:'thin',
        color:'CCCCCC'
    },
    bottom: {
        style:'thick'
    },
    left: {
        style:'thin'
    },
    right: {
        style:'thin'
    }
});

ws.Cell(1, 1).Style(myStyle);
ws.Cell(1, 2).String('My 2nd String').Style(myStyle);
ws.Cell(2, 1).Style(myStyle2);
ws.Cell(2, 2).Style(myStyle2);
ws.Cell(2, 3).Style(myStyle2);
ws.Cell(2, 4).Style(myStyle3);
```

Available styles:

- `Font.Bold()` bolds text
- `Font.Italics()` italicizes text
- `Font.Underline()` underlines text
- `Font.Family('Arial')` name of font family
- `Font.Color('DDEEFF')` hex rgb font color
- `Font.Size(12)` font size in Pts
- `Font.WrapText()` set text wrapping
- `Font.Alignment.Vertical('top')` options are `top`, `center`, `bottom`
- `Font.Alignment.Horizontal('left')` options are `left`, `center`, `right`
- `Font.Alignment.Rotation(15)` degrees to rotate
- `Number.Format('style')` number style string
- `Fill.Color('DDEEFF')` background color in rgb
- `Fill.Pattern('solid')` pattern style `solid`, `lightUp`, etc.
- `Border({...})` border styles (see below)

Border Styles:

Takes one argument: object defining border. Each ordinal (top, right, etc) are only required if you want to define a border. If omitted, no border will be added to that side.  Style is required if oridinal is defined. If color is omitted, it will default to black.

```javascript
myStyle3.Border({
    top: {
        style: 'thin',
        color: 'CCCCCC'
    },
    right: {
        style: 'thin',
        color: 'CCCCCC'
    },
    bottom: {
        style: 'thin',
        color: 'CCCCCC'
    },
    left: {
        style: 'thin',
        color: 'CCCCCC'
    },
    diagonal: {
        style: 'thin',
        color: 'CCCCCC'
    }
});
```


Apply formatting directly to a cell (similar syntax to creating styles):

```javascript
ws.Cell(1, 1).Format.Font.Color('FF0000');
ws.Cell(1, 1).Format.Fill.Pattern('solid');
ws.Cell(1, 1).Format.Fill.Color('AEAEAE');
```

Merge cells and apply styles or mormats to ranges:

`ws.Cell(row1, col1, row2, col2, merge)`

```javascript
ws.Cell(1, 1, 2, 5, true).String('Merged Cells');
ws.Cell(3, 1, 4, 5).String('Each Cell in Range Contains this String');
ws.Cell(3, 1, 4, 5).Style(myStyle);
ws.Cell(1, 1, 2, 5).Format.Font.Family('Arial');
```


## Conditional Formatting

Conditional formatting adds custom formats in response to cell reference state. A subset of conditional formatting features is currently supported by excel4node.

Formatting rules apply at the worksheet level.

The following example will highlight all cells between A1 and A10 that contain the string "ok" with bold, green text:

```javascript
var wb = new xl.WorkBook();
var ws = wb.WorkSheet('My Worksheet');

var style = wb.Style();
style.Font.Bold();
style.Font.Color('00FF00');

ws.addConditionalFormattingRule('A1:A10', {      // apply ws formatting ref 'A1:A10'
    type: 'expression',                          // the conditional formatting type
    priority: 1,                                 // rule priority order (required)
    formula: 'NOT(ISERROR(SEARCH("ok", A1)))',   // formula that returns nonzero or 0
    style: style                                 // a style object containing styles to apply
});
```

**The only conditional formatting type that is currently supported is `expression`.**

When the formula returns zero, conditional formatting is NOT displayed. When the formula returns a nonzero value, conditional formatting is displayed.


## Images

Images can be inserted into a worksheet.

`img.Position(row, col, [rowOffset], [colOffset])`

```javascript
var imgPath = './my-image.jpg'; // relative path from node script
var img1 = ws.Image(imgPath);
img1.Position(1,1);
```

Set image position directly:

```javascript
var img2 = ws.Image(imgPath2).Position(
    3,       // row to anchor top left corner of image
    3,       // col to anchor top left corner of image
    1000000, // offset from top of row in EMUs
    2000000  // offset from left of col in EMUs
);
``` 
Position images across single or multiple cells

```javascript
//Absolute position near D3
//arguments: y-pixels, x-pixels
ws.Image('sampleFiles/image1.png', ws.Image.ABSOLUTE).Position(218, 400).Size(255, 50); 

//A3
//arguments: row, column, {offsetY, offsetX} (in pixels optional)
ws.Image('sampleFiles/image1.png', ws.Image.ONE_CELL).Position(3, 1, 10, 40).Size(255, 50); 

//A1-F2
//arguments: begin-row, begin-column, end-row, end-column, {offsetY, offsetX} (in pixels - optional)
ws.Image('sampleFiles/image1.png', ws.Image.TWO_CELL).Position(1, 1, 2, 6, 2, 5); 

//D5
//arguments: row, column
ws.Image('sampleFiles/image1.png', ws.Image.TWO_CELL).Position(5, 4); 
```

Currently images should be saved at a resolution of 96dpi.

--------------------------------------------------------------------------------

## Writing Output

Write the Workbook to local file synchronously or
Write the Workbook to local file asynchrously or
Send file via node response

```javascript
wb.write('My Excel File.xlsx'); // write synchronously

wb.write('My Excel File.xlsx', function (err) {
    // done writing
});

wb.write('My Excel File.xlsx', res); // write to http response
```


## Notes

- [MS-XSLX spec (pdf)](http://download.microsoft.com/download/D/3/3/D334A189-E51B-47FF-B0E8-C0479AFB0E3C/[MS-XLSX].pdf)
