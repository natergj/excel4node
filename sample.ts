const xl = require('./src/index');

const wb = new xl.Workbook({
  logLevel: 5,
});
const ws = wb.addWorksheet('sheet 1', {});
ws.cell(1, 1).string('hello');
ws.cell(2, 1).string('hello');
ws.cell(3, 1).string('hello');
ws.cell(4, 1).string('hello');

var complexString = [
  'Workbook default font String\n',
  {
    bold: true,
    underline: true,
    italics: true,
    color: 'FF0000',
    size: 18,
    name: 'Courier',
    value: 'Hello',
  },
  ' World!',
  {
    color: '000000',
    underline: false,
    name: 'Arial',
    vertAlign: 'subscript',
  },
  ' All',
  ' these',
  ' strings',
  ' are',
  ' black subsript,',
  {
    color: '0000FF',
    value: '\nbut',
    vertAlign: 'baseline',
  },
  ' now are blue',
];
ws.cell(5, 1).string(complexString);

wb.write('Sample.xlsx');
