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

* Font  
  * Bold  
    * Takes no arguments. Bolds text
  * Italics
  	* Takes no arguments. Italicizes text
  * Underline
  	* Takes no arguments. Underlines text
  * Family
  	* Takes one argument: name of font family.
  * Color
  	* Takes one argument: rbg color
  * Size  
  	* Takes one argument: size in Pts
  * WrapText
  	* Takes no arguments. Set text wrapping to true.
  * Alignment
  	* Vertical
  		* Takes one argument of options top, center, bottom
  	* Horizontal
  		* Takes one argument of left, center, right
* Number  
  * Format
  	* Takes one argument: Number style string
* Fill  
  * Color
  	* Takes one argument: Color in rgb
  * Pattern
  	* Takes one argument: pattern style (solid, lightUp, etc)
* Border 
  * Takes one argument: object defining border
  * each ordinal (top, right, etc) are only required if you want to define a border. If omitted, no border will be added to that side. 
  * style is required if oridinal is defined. if color is omitted, it will default to black. 
  * ```
  {
  		top:{
  			style:'thin',
  			color:'CCCCCC'
  		},
  		right:{
  			style:'thin',
  			color:'CCCCCC'
  		},
  		bottom:{
  			style:'thin',
  			color:'CCCCCC'
  		},
  		left:{
  			style:'thin',
  			color:'CCCCCC'
  		},
  		diagonal:{
  			style:'thin',
  			color:'CCCCCC'
  		}
  	}
  ```


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
myStyle3.Border({
	top:{
		style:'thin',
		color:'CCCCCC'
	},
	bottom:{
		style:'thick'
	},
	left:{
		style:'thin'
	},
	right:{
		style:'thin'
	}
});

ws.Cell(1,1).Style(myStyle);
ws.Cell(1,2).String('My 2nd String').Style(myStyle);
ws.Cell(2,1).Style(myStyle2);
ws.Cell(2,2).Style(myStyle2);
ws.Cell(2,3).Style(myStyle2);
ws.Cell(2,4).Style(myStyle3);
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
* Add ability to apply styles to cell range
* Add ability to merge cells
* Add Text formatting options (cell with more than one font size/color/decoration)
