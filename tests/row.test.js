let test = require('tape');
let xl = require('../distribution/index');
let Row = require('../distribution/lib/row/row.js');

test('Row Tests', (t) => {

    let rowWb = new xl.Workbook({ logLevel: 5 });
    let rowWS = rowWb.addWorksheet();

    t.ok(rowWS.row(2) instanceof Row, 'Successfully accessed a row object');
    t.ok(rowWS.rows['2'] instanceof Row, 'Row was successfully added to worksheet object');

    rowWS.row(2).setHeight(40);
    t.equals(rowWS.row(2).height, 40, 'Row height successfully changed');

    rowWS.row(2).freeze(4);
    t.equals(rowWS.opts.sheetView.pane.ySplit, 2, 'Worksheet set to freeze pane at row 2');
    t.equals(rowWS.opts.sheetView.pane.topLeftCell, 'A4', 'Worksheet set to freeze pane at row 2 and scrollTo row 4');

    rowWS.column(4).freeze();
    t.equals(rowWS.opts.sheetView.pane.topLeftCell, 'E4', 'topLeftCell updated when column was also frozen');

    t.end();
});