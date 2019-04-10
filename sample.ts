import { Workbook } from './src';

const wb = new Workbook({
  defaultWorkbookView: {
    windowHeight: 34000,
    activeTab: 1,
  },
  logLevel: 5,
});
const ws = wb.addWorksheet('sheet 1', {});
const ws2 = wb.addWorksheet('sheet 2', {});
ws.cell(1, 1).string('hello');

ws2.cell(1, 1).string('Hello World');

wb.write('Sample.xlsx');
