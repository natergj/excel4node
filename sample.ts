const xl = require('./src/index');

const wb = new xl.Workbook();
wb.addWorksheet('sheet 1', {});

wb.write();
