let test = require('tape');
let xl = require('../distribution/index');
let Column = require('../distribution/lib/column/column.js');

test('Column Tests', (t) => {

    let wb = new xl.Workbook();
    let ws = wb.addWorksheet();

    t.ok(ws.column(2) instanceof Column, 'Successfully accessed a column object');
    t.ok(ws.cols['2'] instanceof Column, 'Column was successfully added to worksheet object');

    ws.column(2).setWidth(40);
    t.equals(ws.column(2).width, 40, 'Column width successfully changed to integer');

    ws.column(2).setWidth(40.5);
    t.equals(ws.column(2).width, 40.5, 'Column width successfully changed to float');

    ws.column(2).freeze(4);
    t.equals(ws.opts.sheetView.pane.xSplit, 2, 'Worksheet set to freeze pane at column 2');
    t.equals(ws.opts.sheetView.pane.topLeftCell, 'D1', 'Worksheet set to freeze pane at column 2 and scrollTo column 4');

    ws.row(4).freeze();
    t.equals(ws.opts.sheetView.pane.topLeftCell, 'D5', 'topLeftCell updated when row was also frozen');

    t.end();
});