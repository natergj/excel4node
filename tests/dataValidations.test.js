let test = require('tape');
let xl = require('../distribution/index');
let DataValidation = require('../distribution/lib/worksheet/classes/dataValidation.js');

test('DataValidation Tests', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');

    let val1 = ws.addDataValidation({
        type: 'whole',
        errorStyle: 'warning',
        operator: 'greaterThan',
        showInputMessage: '1',
        showErrorMessage: '1',
        errorTitle: 'Invalid Data',
        error: 'The value must be a whole number greater than 0.',
        promptTitle: 'Whole Number',
        prompt: 'Please enter a whole number greater than 0.',
        sqref: 'A1:B1',
        formulas: [
            0
        ]
    });

    let val2 = ws.addDataValidation({
        type: 'list',
        allowBlank: 1,
        showInputMessage: 1,
        showErrorMessage: 1,
        sqref: 'X2:X10',
        formulas: [
            'value1,value2'
        ]
    });

    let val3 = ws.addDataValidation({
        type: 'whole',
        errorStyle: 'warning',
        operator: 'between',
        showInputMessage: '1',
        showErrorMessage: '1',
        errorTitle: 'Invalid Data',
        error: 'The value must be a whole number greater than 0.',
        promptTitle: 'Whole Number',
        prompt: 'Please enter a whole number greater than 0.',
        sqref: 'A10:D10',
        formulas: [0, 10]
    });

    t.ok(
        val1 instanceof DataValidation.DataValidation && 
        val2 instanceof DataValidation.DataValidation && 
        val3 instanceof DataValidation.DataValidation && 
        ws.dataValidationCollection.length === 3, 
        'Data Validations Created'
    );
    t.ok(val1.formula1 === 0 && val1.formula2 === undefined, 'formula\'s of first validation correctly set');
    t.ok(val2.formula1 === 'value1,value2' && val2.formula2 === undefined, 'formula\'s of 2nd validation correctly set');
    t.ok(val3.formula1 === 0 && val3.formula2 === 10, 'formula\'s of 3rd validation correctly set');
    try {
        let val4 = ws.addDataValidation({
            type: 'list',
            allowBlank: 1,
            showInputMessage: 1,
            showErrorMessage: 1,
            //sqref: 'X2:X10',
            formulas: [
                'value1,value2'
            ]
        });
        t.ok(val4 instanceof DataValidation === false,  'init of DataValidation with missing properties should throw an error');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'init of DataValidation with missing properties should throw an error'
        );
    }

    t.end();
});