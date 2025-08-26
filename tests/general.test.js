global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/taskpane/customProp')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/taskpane/taskpane')

global.WordCustomProp =             require('../src/taskpane/WordCustomProp');
global.ExcelCustomProp =            require('../src/taskpane/ExcelCustomProp');

const Document =                    require('./helpers/document');
const WordDocument =                require('./helpers/word/wordDocument');
const excelWorkbook =               require('./helpers/excel/excelWorkbook');

global.Word =                       require('./helpers/word/word')
global.Excel =                      require('./helpers/excel/excel')

test('Word length of array',
    async () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Word.context.document = new WordDocument();

        const info = {host: Office.HostType.Word, };
        await taskpane.init(info);

        await expect(global.document.elements["app-body"].children.length).toBe(5);

    });

test('Excel test length of array',
    async () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Excel.context.workbook = new excelWorkbook();

        const info = {host: Office.HostType.Excel, };
        await taskpane.init(info);

        await expect(global.document.elements["app-body"].children.length).toBe(5);

    });

test('Word onclick',
    async () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Word.context.document = new WordDocument();

        const info = {host: Office.HostType.Word, };
        await taskpane.init(info);
        
        const button = global.document.elements["app-body"].children.find(
            el => el.id === "Default"
        );

        button.onclick();

        
        //await expect(global.Word.context.document.properties.customProperties.items.name).toBe("Classification");
        //await expect(global.Word.context.document.properties.customProperties.items.value).toBe("Default");
        
        await expect(global.document.elements["app-body"].children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");

    });