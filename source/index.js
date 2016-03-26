/* REFERENCES
    http://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf
    http://www.ecma-international.org/publications/standards/Ecma-376.htm
    http://www.openoffice.org/sc/excelfileformat.pdf 
    http://officeopenxml.com/anatomyofOOXML-xlsx.php
*/

/* 
    Code references specifications sections from ECMA-376 2nd edition doc
    ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference.pdf
    found in ECMA-376 2nd edition Part 1 download at http://www.ecma-international.org/publications/standards/Ecma-376.htm
    Sections are referenced in code comments with § 
*/

const utils = require('./lib/utils.js');
module.exports = {
    WorkBook: require('./lib/workbook/index.js'),
    getExcelRowCol: utils.getExcelRowCol,
    getExcelAlpha: utils.getExcelAlpha,
    getExcelTS: utils.getExcelTS,
    PAPER_SIZE: require('./lib/constants/PAPER_SIZE.js')
};
