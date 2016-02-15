let test = require('tape');
let xl = require('../source/index');

test('Cell coverage', (t) => {
    t.plan(1);
    let wb = new xl.WorkBook();
    let ws = wb.WorkSheet('test');
    let cellAccessor = ws.Cell(1, 1);
    t.ok(cellAccessor, 'Correctly generated cellAccessor object');
});

test('Cell returns correct number of cell references', (t) => {
	t.plan(1);
	let wb = new xl.WorkBook();
	let ws = wb.WorkSheet('test');
	let cellAccessor = ws.Cell(1, 1, 5, 2);
	t.ok(cellAccessor.excelRefs.length === 10, 'cellAccessor returns correct number of cellRefs');
});

test('Add String to cell', (t) => {
	t.plan(3);
    let wb = new xl.WorkBook();
    let ws = wb.WorkSheet('test');
    let cell = ws.Cell(1, 1).String('my test string');
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(thisCell.t === 's', 'cellType set to sharedString')
    t.ok(typeof(thisCell.v) === 'number', 'cell Value is a number');
    t.ok(wb.sharedStrings[thisCell.v] === 'my test string', 'Cell sharedString value is correct');
});