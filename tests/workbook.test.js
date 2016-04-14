let test = require('tape');
let xl = require('../distribution/index');
let Font = require('../distribution/lib/style/classes/font.js');

test('Change default workbook options', (t) => {

    let wb = new xl.Workbook();
    let wb2 = new xl.Workbook({
        jszip: {
            compression: 'DEFLATE'
        },
        defaultFont: {
            size: 14,
            name: 'Arial',
            color: 'FFFFFFFF'
        }
    });
    
    let wb1Font = wb.styleData.fonts[0];
    let wb2Font = wb2.styleData.fonts[0];

    t.ok(wb1Font instanceof Font, 'Default Font successfully created');
    t.ok(wb2Font instanceof Font, 'Updated Default Font successfully created');

    t.ok(wb1Font.color === 'FF000000', 'Default font color correctly set');
    t.ok(wb1Font.name === 'Calibri', 'Default font name correctly set');
    t.ok(wb1Font.size === 12, 'Default font size correctly set');
    t.ok(wb1Font.family === 'roman', 'Default font family correctly set');


    t.ok(wb2Font.color === 'FFFFFFFF', 'Default font color correctly updated');
    t.ok(wb2Font.name === 'Arial', 'Default font name correctly updated');
    t.ok(wb2Font.size === 14, 'Default font size correctly updated');
    t.ok(wb2Font.family === 'roman', 'Default font family correctly updated');

    t.end();
});