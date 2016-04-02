let test = require('tape');
let xl = require('../source/index');

test('Generate multiple sheets', (t) => {
    let wb = new xl.WorkBook();
    let ws = wb.addWorksheet('test');
    let ws2 = wb.addWorksheet('test2');
    let ws3 = wb.addWorksheet('test3');
    
    t.ok(wb.sheets.length === 3, 'Correctly generated multiple sheets');

    wb.setSelectedTab(2);
    t.ok(
        wb.sheets[0].opts.sheetView.tabSelected === 0 && 
        wb.sheets[1].opts.sheetView.tabSelected === 1 &&
        wb.sheets[2].opts.sheetView.tabSelected === 0, '2nd Tab set to be default tab selected');

    t.end();
});