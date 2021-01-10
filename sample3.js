const fs = require('fs');
const path = require('path');
const unzipper = require('unzipper')

const xl = require('./source')

const aca_data = {
    'FacultyName': 'كلية الهندسة',
    'AcademicYear': '2018/2019',
    'AcademicGrade': 'الرابع',
    'Department': 'الهندسة الكهربائية و الإلكترونية',
    'Discipline': 'غير متخصص',
    'BoardMeetingNo': '2019/06',
    'BoardMeetingDate': '07/12/2020',

    'AcademicYear_M_1': '2017/2018',
    'AcademicYear_M_2': '2016/2017',

    RegistrarName: 'أ. المحيا الأمين خلف الله محمد خير',
    DeputyDeanName: 'د. مجدي محمد زمراوي',
    DeanName: 'د.علي محمد علي السيوري',

    'HonorsOrGraduates': 'Honours',

    dataTables: {
        RegStudents: {
            "Total": [100, 90, 85, "العدد الكلي"],
            "Examined": [80, 75, 70, "الجالسون"],
            "Passed": [75, 70, 65, "النجاح"],
            "Subs": [3, 2, 4, "البدائل"],
            "Supp": [2, 3, 1, "ازالة الرسوب (الملاحق)"],
            "SubsSupp": [1, 0, 3, "البدائل وازالة الرسوب"],
            "Repeat": [0, 1, 2, "اعادة العام"],
            "Dismissals": [0, 2, 1, "المفصولون"],
            "Recess": [0, 0, 0, "المجمدون"],
            "SpecialCases": [0, 0, 0, "حالات خاصة"],
            "CheatCases": [0, 1, 2, "حالات مخالفة لائحة الامتحانات"]
        },
        ExtStudents: {
            "Total": [100, 90, 85],
            "Examined": [80, 75, 70],
            "Passed": [75, 70, 65],
            "Subs": [3, 2, 4],
            "Supp": [2, 3, 1],
            "SubsSupp": [1, 0, 3],
            "Failed": [0, 1, 2],
            "SpecialCases": [0, 0, 0],
            "CheatCases": [0, 1, 2]
        },
        RegStudentsHonours: {
            Total: [30, 25, 23, "العدد الكلي"],
            FirstClass: [20, 15, 19, "مرتبة الشرف الأولى"],
            SecondClassFirst: [10, 14, 18, "مرتبة الشرف الثانية - القسم الأول"],
            SecondClassSecond: [1, 2, 3, "مرتبة الشرف الثانية - القسم الثاني"],
            ThirdClass: [4, 5, 6, "مرتبة الشرف الثالثة"]
        },
        ExtStudentsHonours: {
            Total: [30, 25, 23, ""],
            FirstClass: [20, 15, 19, "مرتبة الشرف الأولى"],
            SecondClassFirst: [10, 14, 18, "مرتبة الشرف الثانية - القسم الأول"],
            SecondClassSecond: [1, 2, 3, "مرتبة الشرف الثانية - القسم الثاني"],
            ThirdClass: [4, 5, 6, "مرتبة الشرف الثالثة"]
        }
    }
}

var wb = new xl.Workbook({
    defaultFont: {
        name: 'Sakkal Majalla',
        size: 14
    },
    dateFormat: 'mm/dd/yyyy hh:mm:ss',
    logLevel: 1,
    workbookView: {
        windowWidth: 28800,
        windowHeight: 17620,
        xWindow: 240,
        yWindow: 480,
    },
    author: 'Mohanad Ahmed',
    calculationProperties: {
        fullCalculationOnLoad: true
    }
});

var wx = wb.addWorksheet("Sheet2", {
    pageSetup: {
        fitToWidth: 1,
        paperSize: 'A4_PAPER',
        orientation: 'portrait'
    },
    sheetView: {
        'rightToLeft': true
    }
})

wx.column(1).width = 3.14;
wx.column(2).width = 26.57;
wx.column(3).width = 8.43;
wx.column(4).width = 8.43;
wx.column(5).width = 8.43;
wx.column(6).width = 8.43;
wx.column(7).width = 8.43;
wx.column(8).width = 8.43;

var borderx = {
    left: { style: 'medium', color: 'black' },
    right: { style: 'medium', color: 'black' },
    top: { style: 'medium', color: 'black' },
    bottom: { style: 'medium', color: 'black' }
};
// var bordermed = {
//     left: { style: 'thick', color: 'black' },
//     right: { style: 'thick', color: 'black' },
//     top: { style: 'thick', color: 'black' },
//     bottom: { style: 'thick', color: 'black' }
// };
var bordert_top = {
    top: { style: 'thick', color: 'black' },
};
var sumStyle = {
    font: { bold: true, size: 14 },
    alignment: {
        horizontal: 'center',
        vertical: 'center'
    },
    border: borderx
};

var titleStyle = {
    font: { bold: true, size: 18 },
    alignment: {
        horizontal: 'center',
        vertical: 'center'
    }
};

wx.cell(1, 2, 1, 8, true).string('جامعة الخرطوم').style(titleStyle);
wx.cell(2, 2, 2, 8, true).string('أمانة الشؤون العلمية').style(titleStyle)
wx.cell(3, 2, 3, 8, true).string('الإحصائية العامة لنتائج الامتحانات النهائية - الخريجون - مرتبة الشرف')
    .style(titleStyle)

wx.cell(4, 2, 4, 8, true).style({ border: bordert_top })

wx.row(1).height = 20.25;
wx.row(2).height = 20.25;
wx.row(3).height = 20.25;
wx.row(4).height = 3.75;
wx.row(5).height = 9.75;

var topDataStyle = { border: borderx, alignment: { horizontal: 'right' } };
wx.cell(6, 3, 6, 8, true).string(aca_data.FacultyName).style(topDataStyle)
wx.cell(7, 3, 7, 8, true).string(aca_data.AcademicYear).style(topDataStyle)
wx.cell(8, 3, 8, 8, true).string(aca_data.AcademicGrade).style(topDataStyle)
wx.cell(9, 3, 9, 8, true).string(aca_data.Department).style(topDataStyle)
wx.cell(10, 3, 10, 8, true).string(aca_data.Discipline).style(topDataStyle)
wx.cell(11, 3, 11, 8, true).string(aca_data.BoardMeetingNo).style(topDataStyle)
wx.cell(12, 3, 12, 8, true).string(aca_data.BoardMeetingDate).style(topDataStyle)

var topLabelsStyle = { border: borderx, font: { bold: true } };
wx.cell(6, 2).string('الكلية').style(topLabelsStyle)
wx.cell(7, 2).string('العام الدراسي').style(topLabelsStyle)
wx.cell(8, 2).string('المستوى').style(topLabelsStyle)
wx.cell(9, 2).string('القسم').style(topLabelsStyle)
wx.cell(10, 2).string('التخصص').style(topLabelsStyle)
wx.cell(11, 2).string('رقم اجتماع مجلس الكلية').style(topLabelsStyle)
wx.cell(12, 2).string('تاريخ الاجتماع').style(topLabelsStyle)

const BigLabels = ['أولا', 'ثانيا', 'ثالثا', 'رابعا', 'خامسا', 'سادسا'];
var currentRow = 14;
var tabNum = 0;
let vTitle = '';

var regStud = aca_data.dataTables.RegStudents;
wx.cell(currentRow, 2).string('اولاً: الطلاب النظاميون:').style({ font: { bold: true, size: 16 } })
GenerateACADataTable(wx, currentRow + 1, aca_data, regStud)
currentRow += Object.keys(regStud).length + 6;
wx.cell(currentRow - 2, 2).string('المجموع').style(sumStyle)
wx.cell(currentRow - 2, 3).formula(`SUM(${xl.getExcelCellRef(currentRow - 11, 3)}:${xl.getExcelCellRef(currentRow - 3, 3)})`).style(sumStyle)
wx.cell(currentRow - 2, 4).formula(`SUM(${xl.getExcelCellRef(currentRow - 11, 4)}:${xl.getExcelCellRef(currentRow - 3, 4)})`).style(sumStyle)
wx.cell(currentRow - 2, 5).formula(`SUM(${xl.getExcelCellRef(currentRow - 11, 5)}:${xl.getExcelCellRef(currentRow - 3, 5)})`).style(sumStyle)
wx.cell(currentRow - 2, 6).formula(`SUM(${xl.getExcelCellRef(currentRow - 11, 6)}:${xl.getExcelCellRef(currentRow - 3, 6)})`).style(sumStyle)
wx.cell(currentRow - 2, 7).formula(`SUM(${xl.getExcelCellRef(currentRow - 11, 7)}:${xl.getExcelCellRef(currentRow - 3, 7)})`).style(sumStyle)
wx.cell(currentRow - 2, 8).formula(`SUM(${xl.getExcelCellRef(currentRow - 11, 8)}:${xl.getExcelCellRef(currentRow - 3, 8)})`).style(sumStyle)
tabNum += 1;

if (aca_data.HonorsOrGraduates) {
    vTitle = BigLabels[tabNum] + ': ' + 'الحاصلون على مرتبة الشرف (نظاميون)';
    var regStudHonours = aca_data.dataTables.RegStudentsHonours;
    wx.cell(currentRow, 2).string(vTitle).style({ font: { bold: true, size: 16 } })
    GenerateACADataTable(wx, currentRow + 1, aca_data, regStudHonours)
    currentRow += Object.keys(regStudHonours).length + 6;
    var crw = currentRow - 2;
    wx.cell(crw, 2).string('المجموع').style(sumStyle)
    wx.cell(crw, 3).formula(`SUM(${xl.getExcelCellRef(crw - 4, 3)}:${xl.getExcelCellRef(crw - 1, 3)})`).style(sumStyle)
    wx.cell(crw, 4).formula(`SUM(${xl.getExcelCellRef(crw - 4, 4)}:${xl.getExcelCellRef(crw - 1, 4)})`).style(sumStyle)
    wx.cell(crw, 5).formula(`SUM(${xl.getExcelCellRef(crw - 4, 5)}:${xl.getExcelCellRef(crw - 1, 5)})`).style(sumStyle)
    wx.cell(crw, 6).formula(`SUM(${xl.getExcelCellRef(crw - 4, 6)}:${xl.getExcelCellRef(crw - 1, 6)})`).style(sumStyle)
    wx.cell(crw, 7).formula(`SUM(${xl.getExcelCellRef(crw - 4, 7)}:${xl.getExcelCellRef(crw - 1, 7)})`).style(sumStyle)
    wx.cell(crw, 8).formula(`SUM(${xl.getExcelCellRef(crw - 4, 8)}:${xl.getExcelCellRef(crw - 1, 8)})`).style(sumStyle)
    tabNum += 1;
}

vTitle = BigLabels[tabNum] + ': ' + 'الممتحنون من الخارج';
var extStud = aca_data.dataTables.ExtStudents;
wx.cell(currentRow, 2).string(vTitle).style({ font: { bold: true, size: 16 } })
GenerateACADataTable(wx, currentRow + 1, aca_data, extStud)
currentRow += Object.keys(extStud).length + 6;
var crw = currentRow - 2;
wx.cell(crw, 2).string('المجموع').style(sumStyle)
wx.cell(crw, 3).formula(`SUM(${xl.getExcelCellRef(crw - 7, 3)}:${xl.getExcelCellRef(crw - 1, 3)})`).style(sumStyle)
wx.cell(crw, 4).formula(`SUM(${xl.getExcelCellRef(crw - 7, 4)}:${xl.getExcelCellRef(crw - 1, 4)})`).style(sumStyle)
wx.cell(crw, 5).formula(`SUM(${xl.getExcelCellRef(crw - 7, 5)}:${xl.getExcelCellRef(crw - 1, 5)})`).style(sumStyle)
wx.cell(crw, 6).formula(`SUM(${xl.getExcelCellRef(crw - 7, 6)}:${xl.getExcelCellRef(crw - 1, 6)})`).style(sumStyle)
wx.cell(crw, 7).formula(`SUM(${xl.getExcelCellRef(crw - 7, 7)}:${xl.getExcelCellRef(crw - 1, 7)})`).style(sumStyle)
wx.cell(crw, 8).formula(`SUM(${xl.getExcelCellRef(crw - 7, 8)}:${xl.getExcelCellRef(crw - 1, 8)})`).style(sumStyle)

tabNum += 1;

if (aca_data.HonorsOrGraduates) {
    vTitle = BigLabels[tabNum] + ': ' + 'الحاصلون على مرتبة الشرف (خارجيون) ';
    var extStudHonours = aca_data.dataTables.ExtStudentsHonours;
    wx.cell(currentRow, 2).string(vTitle).style({ font: { bold: true, size: 16 } })
    GenerateACADataTable(wx, currentRow + 1, aca_data, extStudHonours)
    currentRow += Object.keys(extStudHonours).length + 6;
    var crw = currentRow - 2;
    wx.cell(crw, 2).string('المجموع').style(sumStyle)
    wx.cell(crw, 3).formula(`SUM(${xl.getExcelCellRef(crw - 4, 3)}:${xl.getExcelCellRef(crw - 1, 3)})`).style(sumStyle)
    wx.cell(crw, 4).formula(`SUM(${xl.getExcelCellRef(crw - 4, 4)}:${xl.getExcelCellRef(crw - 1, 4)})`).style(sumStyle)
    wx.cell(crw, 5).formula(`SUM(${xl.getExcelCellRef(crw - 4, 5)}:${xl.getExcelCellRef(crw - 1, 5)})`).style(sumStyle)
    wx.cell(crw, 6).formula(`SUM(${xl.getExcelCellRef(crw - 4, 6)}:${xl.getExcelCellRef(crw - 1, 6)})`).style(sumStyle)
    wx.cell(crw, 7).formula(`SUM(${xl.getExcelCellRef(crw - 4, 7)}:${xl.getExcelCellRef(crw - 1, 7)})`).style(sumStyle)
    wx.cell(crw, 8).formula(`SUM(${xl.getExcelCellRef(crw - 4, 8)}:${xl.getExcelCellRef(crw - 1, 8)})`).style(sumStyle)
}

function GenerateACADataTable(ws, srow, acd, dtab, tabNum) {

    if (true) {
        var xsty = {
            font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                left: { style: 'thin', color: 'black' },
                right: { style: 'thin', color: 'black' },
                top: { style: 'thin', color: 'black' },
                bottom: { style: 'thin', color: 'black' }
            }
        };

        var zsty = {
            font: { bold: true }, alignment: { horizontal: 'right', vertical: 'center' },
            border: {
                left: { style: 'medium', color: 'black' },
                right: { style: 'medium', color: 'black' },
                top: { style: 'thin', color: 'black' },
                bottom: { style: 'thin', color: 'black' }
            }
        };

        var r1sty = {
            font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                left: { style: 'medium', color: 'black' },
                right: { style: 'medium', color: 'black' },
                top: { style: 'medium', color: 'black' },
            }
        };

        var r2sty = {
            font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                left: { style: 'medium', color: 'black' },
                right: { style: 'medium', color: 'black' },
                bottom: { style: 'thin', color: 'black' },
            }
        };

        var r3lsty = {
            font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                left: { style: 'thin', color: 'black' },
                right: { style: 'medium', color: 'black' },
                bottom: { style: 'medium', color: 'black' },
                top: { style: 'thin', color: 'black' },
            }
        };
        var r3rsty = {
            font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                left: { style: 'medium', color: 'black' },
                right: { style: 'thin', color: 'black' },
                bottom: { style: 'medium', color: 'black' },
                top: { style: 'thin', color: 'black' },
            }
        };


        ws.cell(srow, 2, srow + 2, 2, true).string('الوصف').style({
            font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                left: { style: 'medium', color: 'black' },
                right: { style: 'medium', color: 'black' },
                top: { style: 'medium', color: 'black' },
                bottom: { style: 'medium', color: 'black' }
            }
        });

        ws.cell(srow, 3, srow, 4, true).string('العام الدراسي').style(r1sty)
        ws.cell(srow, 5, srow, 6, true).string('العام الدراسي').style(r1sty)
        ws.cell(srow, 7, srow, 8, true).string('العام الدراسي').style(r1sty)

        ws.cell(srow + 1, 3, srow + 1, 4, true).string(acd.AcademicYear).style(r2sty)
        ws.cell(srow + 1, 5, srow + 1, 6, true).string(acd.AcademicYear_M_1).style(r2sty)
        ws.cell(srow + 1, 7, srow + 1, 8, true).string(acd.AcademicYear_M_2).style(r2sty)

        ws.cell(srow + 2, 3).string('العدد').style(r3rsty)
        ws.cell(srow + 2, 4).string('النسبة').style(r3lsty)
        ws.cell(srow + 2, 5).string('العدد').style(r3rsty)
        ws.cell(srow + 2, 6).string('النسبة').style(r3lsty)
        ws.cell(srow + 2, 7).string('العدد').style(r3rsty)
        ws.cell(srow + 2, 8).string('النسبة').style(r3lsty)
    }
    var n = 3;
    for (var label in dtab) {
        var val = dtab[label];
        // console.log(val)
        var zlab = val[3] ? val[3] : label;
        ws.cell(srow + n, 2).string(zlab).style(zsty)
        ws.cell(srow + n, 3).number(val[0]).style(xsty)
        ws.cell(srow + n, 5).number(val[1]).style(xsty)
        ws.cell(srow + n, 7).number(val[2]).style(xsty)
        ws.row(srow + n).height = 17.25;
        n++;
    }
}

wx.addChart({
    type: 'chart',
    chartType: 'bar',
    chartData: {
        title: 'الطلاب النظاميون',
        dataSeries: [
            {
                label: 'العام الحالي', range: "'Sheet2'!$D$18:$D$28",
                pattern: 'solidDmnd', catRange: "'Sheet2'!$B$18:$B$28"
            },
            { label: 'العام السابق', pattern: 'smGrid', range: "'Sheet2'!$F$18:$F$28", catRange: "'Sheet2'!$B$18:$B$28" },
            { label: 'العام الحالي', pattern: 'wdDnDiag', range: "'Sheet2'!$H$18:$H$28", catRange: "'Sheet2'!$B$18:$B$28" },
        ],
        dataLegend: [

        ],
        catAxisLabels: {
            range: "'Sheet2'!$B$18:$B$28"
        }
    },
    position: {
        type: 'twoCellAnchor',
        from: {
            col: 2, row: 69, colOff: 19050, rowOff: 76200
        },
        to: {
            col: 8, row: 85, colOff: 600075, rowOff: 190500
        }
    }
})

// wx.addChart({
//     type: 'chart',
//     chartType: 'bar',
//     chartData: {
//         title: 'الطلاب النظاميون',
//         dataSeries: [
//             {
//                 label: 'العام الحالي', range: "'Sheet2'!$D$18:$D$28",
//                 pattern: 'solidDmnd', catRange: "'Sheet2'!$B$18:$B$28"
//             },
//             { label: 'العام السابق', pattern: 'smGrid', range: "'Sheet2'!$F$18:$F$28", catRange: "'Sheet2'!$B$18:$B$28" },
//             { label: 'العام الحالي', pattern: 'wdDnDiag', range: "'Sheet2'!$H$18:$H$28", catRange: "'Sheet2'!$B$18:$B$28" },
//         ],
//         dataLegend: [

//         ],
//         catAxisLabels: {
//             range: "'Sheet2'!$B$18:$B$28"
//         }
//     },
//     position: {
//         type: 'twoCellAnchor',
//         from: {
//             col: 10, row: 1, colOff: 19050, rowOff: 76200
//         },
//         to: {
//             col: 16, row: 17, colOff: 600075, rowOff: 190500
//         }
//     }
// })

wb.write('Excel.xlsx')
var dir = require('path').join(__dirname, '/testzip')
fs.createReadStream('Excel.xlsx').pipe(unzipper.Extract({ path: dir }));
