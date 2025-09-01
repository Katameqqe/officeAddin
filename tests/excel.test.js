
global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window');

global.CustomPropertyController =   require('../src/CustomPropertyController');
global.CustomClassification =       require('../src/customClassification');

global.fetch =                      require('./helpers/fetch');

const taskpane =                    require('../src/index');

global.ExcelCustomPropertyController = require('../src/excelCustomPropertyController');

const Document =                    require('./helpers/document');
const ExcelWorkbook =               require('./helpers/excel/excelWorkbook');
const CustomProperty =              require('./helpers/customProperty');

global.Excel =                      require('./helpers/excel/excel');
const info = {host: Office.HostType.Excel, };

beforeEach(() => {
    global.window = new Window();
    global.document = new Document();
    global.Excel.context.workbook = new ExcelWorkbook();
});

test('Excel Display empty classification',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items = [];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(0);
    });

test('Excel Display not empty classification',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items =
        [
            new CustomProperty("Classification", "Default"),
            new CustomProperty("ClassifiedBy", "User"),
            new CustomProperty("ClassificationHost", "Word"),
            new CustomProperty("ClassificationDate", "Date"),
            new CustomProperty("ClassificationGUID", "GUID"),
        ];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
    });

test('Excel set classification from empty',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items = [];
        await taskpane.init(info);

        taskpane.classificationSelected("Default");

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
    });

test('Excel update existed classification',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items =
        [
            new CustomProperty("Classification", "Default"),
            new CustomProperty("ClassifiedBy", "User"),
            new CustomProperty("ClassificationHost", "Word"),
            new CustomProperty("ClassificationDate", "Date"),
            new CustomProperty("ClassificationGUID", "GUID"),
        ];
        await taskpane.init(info);
        
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
        
        await taskpane.classificationSelected("Restricted");

        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Restricted");
    });

test('Excel clear classification',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items =
        [
            new CustomProperty("Classification", "Default"),
            new CustomProperty("ClassifiedBy", "User"),
            new CustomProperty("ClassificationHost", "Word"),
            new CustomProperty("ClassificationDate", "Date"),
            new CustomProperty("ClassificationGUID", "GUID"),
        ];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
        
        await taskpane.removeClassification();
        
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(0);
    });