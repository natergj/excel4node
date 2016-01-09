var select = require('xpath.js');
var dom = require('xmldom').DOMParser;


var xml = '\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<worksheet mc:Ignorable="x14ac"\
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"\
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"\
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"\
  xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">\
  <sheetPr>\
    <outlinePr summaryBelow="1"/>\
  </sheetPr>\
  <sheetViews>\
    <sheetView workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" rightToLeft="0"/>\
  </sheetViews>\
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0"/>\
  <sheetData/>\
  <conditionalFormatting sqref="A1:A10">\
    <cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="??">\
      <formula>NOT(ISERROR(SEARCH("??", A1)))</formula>\
    </cfRule>\
  </conditionalFormatting>\
  <conditionalFormatting sqref="B1:B10">\
    <cfRule type="containsText" dxfId="0" priority="2" operator="containsText" text="??">\
      <formula>NOT(ISERROR(SEARCH("??", A1)))</formula>\
    </cfRule>\
  </conditionalFormatting>\
  <printOptions horizontalCentered="0" verticalCentered="1"/>\
  <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>\
  <pageSetup/>\
</worksheet>\
';


// var xml = '<book><title>Harry Potter</title></book>';
var doc = new dom().parseFromString(xml);
// var nodes = select(doc, '//title');
var nodes = select(doc, 'conditionalFormatting');

console.log(nodes);

// console.log(nodes[0].localName + ': ' + nodes[0].firstChild.data);
// console.log('node: ' + nodes[0].toString());
