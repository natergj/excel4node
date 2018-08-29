const xl = require('./src/index');

const wb = new xl.Workbook({
  logLevel: 5,
});
wb.addWorksheet('sheet 1', {});

wb.write('Sample.xlsx');
