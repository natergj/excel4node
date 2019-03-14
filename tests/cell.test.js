const test = require('tape');
const DOMParser = require('xmldom').DOMParser;
const xl = require('../source/index');

test('Cell coverage', (t) => {
    t.plan(1);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cellAccessor = ws.cell(1, 1);
    t.ok(cellAccessor, 'Correctly generated cellAccessor object');
});

test('Cell returns correct number of cell references', (t) => {
    t.plan(1);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cellAccessor = ws.cell(1, 1, 5, 2);
    t.ok(cellAccessor.excelRefs.length === 10, 'cellAccessor returns correct number of cellRefs');
});

test('Add String to cell', (t) => {
    t.plan(3);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cell = ws.cell(1, 1).string('my test string');
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(thisCell.t === 's', 'cellType set to sharedString');
    t.ok(typeof (thisCell.v) === 'number', 'cell Value is a number');
    t.ok(wb.sharedStrings[thisCell.v] === 'my test string', 'Cell sharedString value is correct');
});

test('Replace null or undefined value with empty string', (t) => {
    t.plan(3);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cell = ws.cell(1, 1).string(null);
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(thisCell.t === 's', 'cellType set to sharedString');
    t.ok(typeof (thisCell.v) === 'number', 'cell Value is a number');
    t.ok(wb.sharedStrings[thisCell.v] === '', 'Cell is empty string');
});

test('Add Number to cell', (t) => {
    t.plan(3);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cell = ws.cell(1, 1).number(10);
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(thisCell.t === 'n', 'cellType set to number');
    t.ok(typeof (thisCell.v) === 'number', 'cell Value is a number');
    t.ok(thisCell.v === 10, 'Cell value value is correct');
});

test('Add Boolean to cell', (t) => {
    t.plan(3);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cell = ws.cell(1, 1).bool(true);
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(thisCell.t === 'b', 'cellType set to boolean');
    t.ok(typeof (thisCell.v) === 'string', 'cell Value is a string');
    t.ok(thisCell.v === 'true' || thisCell.v === 'false', 'Cell value value is correct');
});

test('Add Formula to cell', (t) => {
    t.plan(4);
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cell = ws.cell(1, 1).formula('SUM(A1:A10)');
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(thisCell.t === null, 'cellType is not set');
    t.ok(thisCell.v === null, 'cellValue is not set');
    t.ok(typeof (thisCell.f) === 'string', 'cell Formula is a string');
    t.ok(thisCell.f === 'SUM(A1:A10)', 'Cell value value is correct');
});

test('Add Comment to cell', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cell = ws.cell(1, 1).comment('My test comment');
    let ref = cell.excelRefs[0];
    t.ok(ws.comments[ref].comment === 'My test comment');
    ws.generateCommentsXML().then((XML) => {
        let doc = new DOMParser().parseFromString(XML);
        let testComment = doc.getElementsByTagName('commentList')[0];
        t.ok(testComment.textContent === 'My test comment', 'Verify comment text is correct');
        t.end()
    });
});

test('Complex String value should NOT be rewritten from Style', (t) => {
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('test');
    const arr = [ '1', '2', '3' ];
    const style = {
        bold: true,
        underline: true,
        italics: false,
        color: '000000'
    };
    for (const i in arr) {
        const el = arr[i];
        style.value = el;
        ws.cell(1, parseInt(i)+1).string([ style ]);
    }
    const cells = ws.cells;
    const values = Object.values(cells).map( e => wb.sharedStrings[e.v][0].value||wb.sharedStrings[e.v] );
    t.ok(values[0]==arr[0], '1 cell value equals 1 array value');
    t.ok(values[1]==arr[1], '2 cell value equals 2 array value');
    t.ok(values[2]==arr[2], '3 cell value equals 3 array value');
    t.end();
});
