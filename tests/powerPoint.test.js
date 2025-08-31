
global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/customProp')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/index')

const Document =                    require('./helpers/document');

test('Display empty classification',
    async () =>
    {
        // TODO: implement
        // global.window = new Window();
        // global.document = new Document();
        // global.Excel.context.workbook = new excelWorkbook();

        // const info = {host: Office.HostType.Excel, };
        // await taskpane.init(info);

        await expect(true).toBe(true);
    });

// TODO: set classification from empty

// TODO: update existed classification

// TODO: clear classification
