# excel4node

An OOXML (xlsx) generator that supports formatting options

### Installation:

```
add dependency to your node package.json file
```

### Usage:

Instantiate a new workook

```
var xl = require('excel4node');
var wb = new xl.WorkBook();
```

Add a new WorkSheet to the workbook

```
var ws = wb.WorkSheet('New Worksheet');
```

Add a cell to a WorkSheet with some data. 
Cell can take 3 data types: String, Number, Formula  
Cell takes two arguments: row, col

```
ws.Cell(1,1).String('My String');
ws.Cell(2,1).Number(5);
ws.Cell(2,2).Number(10);
ws.Cell(2,3).Formula("A2+B2");
ws.Cell(2,4).Formula("A2/C2");
```

Set Dimensions of Rows or Columns

```
ws.Row(1).Height(30);
ws.Column(1).Width(100);
```

Create a Style and apply it to a cell

```
var myStyle = wb.Style();
myStyle.Font.Bold();
myStyle.Font.Italics();
myStyle.Font.Underline();
myStyle.Font.Family('Times New Roman');
myStyle.Font.Color('FF0000');
myStyle.Font.Size(16);

var myStyle2 = wb.Style();
myStyle2.Font.Size(14);
myStyle2.Number.Format("$#,##0.00;($#,##0.00);-");

var myStyle3 = wb.Style();
myStyle3.Font.Size(14);
myStyle3.Number.Format("##%");
myStyle3.Fill.Pattern('solid');
mystyle3.Fill.Color('CCCCCC');

ws.Cell(1,1).Style(myStyle);
ws.Cell(1,2).String('My 2nd String').Style(myStyle);
ws.Cell(2,1).Style(myStyle2);
ws.Cell(2,2).Style(myStyle2);
ws.Cell(2,3).Style(myStyle2);
ws.Cell(2,4).Style(myStyle3);
```

Write the Workbook to file or node response

```
wb.write("My Excel File.xlsx");
wb.write("My Excel File.xlsx",res);

```

### ToDo
* Add Date functions
* Add ability to apply styles to cell range
* Add ability to merge cells
* Add Text formatting options (alignment, cell with more than one font size/color/decoration)
