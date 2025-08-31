
global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/customPropertyController')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/index')

global.ExcelCustomProp =            require('../src/excelCustomPropertyController');

const Document =                    require('./helpers/document');
const excelWorkbook =               require('./helpers/excel/excelWorkbook');

global.Excel =                      require('./helpers/excel/excel')

test('Excel test length of array',
    async () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Excel.context.workbook = new excelWorkbook();

        const info = {host: Office.HostType.Excel, };
        await taskpane.init(info);

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
    });

// TODO: Display empty classification

// TODO: Display not empty classification

// TODO: set classification from empty

// TODO: update existed classification

// TODO: clear classification
