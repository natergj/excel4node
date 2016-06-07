const test = require('tape');
const xl = require('../distribution/index');
const DOMParser = require('xmldom').DOMParser;
const DataValidation = require('../distribution/lib/worksheet/classes/dataValidation.js');

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
        type: 'list',
        allowBlank: 1,
        showInputMessage: 1,
        showErrorMessage: 1,
        showDropDown: true,
        sqref: 'X2:X10',
        formulas: [
            'value1,value2'
        ]
    });

    let val4 = ws.addDataValidation({
        type: 'list',
        allowBlank: 1,
        showInputMessage: 1,
        showErrorMessage: 1,
        showDropDown: false,
        sqref: 'X2:X10',
        formulas: [
            'value1,value2'
        ]
    });

    let val5 = ws.addDataValidation({
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
        val4 instanceof DataValidation.DataValidation && 
        val5 instanceof DataValidation.DataValidation && 
        ws.dataValidationCollection.length === 5, 
        'Data Validations Created'
    );
    t.ok(val1.formula1 === 0 && val1.formula2 === undefined, 'formula\'s of first validation correctly set');
    t.ok(val2.formula1 === 'value1,value2' && val2.formula2 === undefined, 'formula\'s of 2nd validation correctly set');
    t.ok(val3.formula1 === 'value1,value2' && val3.formula2 === undefined, 'formula\'s of 3rd validation correctly set');
    t.ok(val5.formula1 === 0 && val5.formula2 === 10, 'formula\'s of 4th validation correctly set');
    try {
        let val6 = ws.addDataValidation({
            type: 'list',
            allowBlank: 1,
            showInputMessage: 1,
            showErrorMessage: 1,
            //sqref: 'X2:X10',
            formulas: [
                'value1,value2'
            ]
        });
        t.ok(val6 instanceof DataValidation === false,  'init of DataValidation with missing properties should throw an error');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'init of DataValidation with missing properties should throw an error'
        );
    }

    ws.generateXML().then((XML) => {
        let doc = new DOMParser().parseFromString(XML);
        let dataValidations = doc.getElementsByTagName('dataValidation');

        t.equals(dataValidations[1].getAttribute('showDropDown'), '', 'showDropDown correclty not set when showDropDown is set to true');
        t.equals(dataValidations[2].getAttribute('showDropDown'), '', 'showDropDown correclty not set when showDropDown is not specified');
        t.equals(dataValidations[3].getAttribute('showDropDown'), '1', 'showDropDown correclty set to 1 when showDropDown is set to false');
        t.end();
    });

    
});