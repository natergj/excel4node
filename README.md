# excel4node

An OOXML (xlsx) generator that supports formatting options

### Installation:

```
npm install excel4node
```

### Sample:
A sample.js script is provided in the code. Running this will output a sample excel workbook named Excel.xlsx

```
node sample.js
```

### Usage:

Instantiate a new workook
Takes optional params object to specify jszip options. More to come.

```
var xl = require('excel4node');
var wb = new xl.WorkBook();

var wbOpts = {
	jszip:{
		compression:'DEFLATE'
	}
}
var wb2 = new xl.WorkBook(wbOpts);
```

Add a new WorkSheet to the workbook
Takes optional params object to specify page margins, zoom and print view centering

```
var ws = wb.WorkSheet('New Worksheet');

var wsOpts = {
	margins:{
		left : .75,
		right : .75,
		top : 1.0,
		bottom : 1.0,
		footer : .5,
		header : .5
	},
	printOptions:{
		centerHorizontal : true,
		centerVertical : false
	},
	view:{
		zoom : 100
	},
	outline:{
		summaryBelow : true
	}
}
var ws2 = wb.WorkSheet('New Worksheet', wsOpts);
```

Add a cell to a WorkSheet with some data.  
Cell can take 4 data types: String, Number, Formula, Date.  
Cell takes two arguments: row, col

```
ws.Cell(1,1).String('My String');
ws.Cell(2,1).Number(5);
ws.Cell(2,2).Number(10);
ws.Cell(2,3).Formula("A2+B2");
ws.Cell(2,4).Formula("A2/C2");
ws.Cell(2,5).Date(new Date());
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


Freeze Columns and Rows to prevent moving when scrolling horizontally  
First example will freeze the first two columns (everything prior to the specified column);  
Second example will freeze the first two columns and scroll to the 8th column.  
Third example will freeze the first two rows (everything prior to the specified row);  
Forth example will freeze the first two rows and scroll to the 8th row.  
See "Series with frozen Row" tab in sample output workbook

```
ws.Column(3).Freeze();
ws.Column(3).Freeze(8);
ws.Row(3).Freeze();
ws.Row(3).Freeze(8);

```
Set a row to be a filter row
Optionally specify start and end columns
If no arguments passed, will add filter to any populated columns
See "Departmental Spending Report" tab in sample output workbook

```
ws.Row(1).Filter();
ws.Row(1).Filter(1,8);
```
Set Groupings on Rows and optionally collapse them.  
See "Groupings Summary Top" and "Groupings Summary Bottom" in sample output.

```
ws.Row(rowNum).Group(level,isCollapsed)
ws.Row(1).Group(1,true)
```

Insert an image into a WorkSheet  for
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

Write the Workbook to local file synchronously or
Write the Workbook to local file asynchrously or
Send file via node response

```
wb.write("My Excel File.xlsx");
wb.write("My Excel File.xlsx",function(err){ ... });
wb.write("My Excel File.xlsx",res);

```

