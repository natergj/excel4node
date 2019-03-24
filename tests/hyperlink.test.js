let test = require('tape');
let xl = require('../source/index');

test('Create Hyperlink', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    ws.cell(1, 1).link('http://iamnater.com', 'iAmNater', 'iAmNater.com');
    t.ok(ws.hyperlinkCollection.links[0].location === 'http://iamnater.com', 'Link location set correctly');
    t.ok(ws.hyperlinkCollection.links[0].display === 'iAmNater', 'Link display set correctly');
    t.ok(ws.hyperlinkCollection.links[0].tooltip === 'iAmNater.com', 'Link tooltip set correctly');
    t.ok(typeof ws.hyperlinkCollection.links[0].id === 'number', 'ID correctly set');
    t.ok(ws.hyperlinkCollection.links[0].rId === 'rId' + ws.hyperlinkCollection.links[0].id, 'Link Ref ID set correctly');
    t.end();
});