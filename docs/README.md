# excel4node

A full featured xlsx file generation library allowing for the creation of advanced Excel files.

excel4node conforms to the ECMA-376 OOXML specification 2nd edition

[![NPM version](https://img.shields.io/npm/v/excel4node.svg)](https://www.npmjs.com/package/excel4node)
[![License](https://img.shields.io/badge/License-MIT-brightgreen.svg)](https://opensource.org/licenses/MIT)
[![npm](https://img.shields.io/npm/dt/excel4node.svg)](https://www.npmjs.com/package/excel4node)
[![node](https://img.shields.io/node/v/excel4node.svg)](https://nodejs.org/en/download/)
[![Build Status](https://travis-ci.org/natergj/excel4node.svg?branch=master)](https://travis-ci.org/natergj/excel4node)
[![dependencies Status](https://david-dm.org/natergj/excel4node/status.svg)](https://david-dm.org/natergj/excel4node)
[![devDependency Status](https://david-dm.org/natergj/excel4node/dev-status.svg)](https://david-dm.org/natergj/excel4node#info=devDependencies)


REFERENCES  
[OpenXML White Paper](http://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf)  
[ECMA-376 Website](http://www.ecma-international.org/publications/standards/Ecma-376.htm)  
[OpenOffice Excel File Format Reference](http://www.openoffice.org/sc/excelfileformat.pdf)  
[OpenOffice Anatomy of OOXML explanation](http://officeopenxml.com/anatomyofOOXML-xlsx.php)  
[MS-XSLX spec (pdf)](http://download.microsoft.com/download/D/3/3/D334A189-E51B-47FF-B0E8-C0479AFB0E3C/%5BMS-XLSX%5D.pdf)

Code references specifications sections from ECMA-376 2nd edition doc  
ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference.pdf  
found in ECMA-376 2nd edition Part 1 download at [http://www.ecma-international.org/publications/standards/Ecma-376.htm](http://www.ecma-international.org/publications/standards/Ecma-376.htm)

## Basic Usage

```javascript
const xl = require('excel4node');

const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet 1');

ws.cell(1, 1).string('Hello World');

wb.write('MyExcelFile.xlsx');
```
