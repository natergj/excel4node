# excel4node

An OOXML (xlsx) generator that supports formatting options

### Installation:

```
npm install excel4node
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
myStyle.Font.Alignment.Vertical('top');
myStyle.Font.Alignment.Horizontal('left');
myStyle.Font.WrapText(true);

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
Apply Formatting to Cell
Syntax similar to creating styles

```
ws.Cell(1,1).Format.Font.Color('FF0000');
ws.Cell(1,1).Format.Fill.Pattern('solid');
ws.Cell(1,1).Format.Fill.Color('AEAEAE');
```

Merge Cells and apply Styles or Formats to ranges
ws.Cell(row1,col1,row2,col2,merge)

```
ws.Cell(1,1,2,5,true).String('Merged Cells');
ws.Cell(3,1,4,5).String('Each Cell in Range Contains this String');
ws.Cell(3,1,4,5).Style(myStyle);
ws.Cell(1,1,2,5).Format.Font.Family('Arial');
```


Freeze Columns to prevent moving when scrolling horizontally  
First example will freeze the first two columns (everything prior to the specified column);  
Second example will freeze the first two columns and scroll to the 8th column.  

```
ws.Column(3).Freeze();
ws.Column(3).Freeze(8);

```
Insert an image into a WorkSheet  
Image takes one argument which is relative path to image from node script  
Image can be passed optional Position which takes 4 arguments  
img.Position(row, col, [rowOffset], [colOffset])  
row = top left corner of image will be anchored to top of this row  
col = top left corner of image will be anchored to left of this column  
rowOffset = offset from top of row in EMUs  
colOfset = offset from left of col in EMUs 
  
Currently images should be saved at a resolution of 96dpi. 

```
var img1 = ws.Image(imgPath);
img1.Position(1,1);

var img2 = ws.Image(imgPath2).Position(3,3,1000000,2000000);
```

Write the Workbook to file or node response

```
wb.write("My Excel File.xlsx");
wb.write("My Excel File.xlsx",res);

```

### ToDo
* Add Date functions
* Add Text formatting options (cell with more than one font size/color/decoration)
* Add ability to collapse rows
