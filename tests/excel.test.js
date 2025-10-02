
global.Office =                             require('./helpers/office');

const Window =                              require('./helpers/window');

global.CustomPropertyController =           require('../src/CustomPropertyController');
global.CustomClassification =               require('../src/customClassification');

global.fetch =                              require('./helpers/fetch');

const taskpane =                            require('../src/index');
global.defaultClassificationFont =          require('../src/defaultVars');

global.CustomXMLController =                require('../src/CustomXMLController');
global.CustomXMLProcessor =                 require('../src/CustomXMLProcessor');
global.CustomPropertyProcessor =            require('../src/CustomPropertyProcessor');

const Document =                            require('./helpers/document');
const ExcelWorkbook =                       require('./helpers/excel/excelWorkbook');
const CustomProperty =                      require('./helpers/customProperty');
const CustomXmlPart =                       require('./helpers/customXmlPart');

global.Excel =                              require('./helpers/excel/excel');
const { DOMParser } =                       require('xmldom');

const info = {host: Office.HostType.Excel, };

const testXMLpart_one =                     require('./helpers/exampleXML');

beforeEach(
    () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Excel.context.workbook = new ExcelWorkbook();
        global.DOMParser = DOMParser;
    });

test('Excel Display empty classification',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items = [];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(0);
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(0);
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
        global.Excel.context.workbook.customXmlParts.items =
        [
            new CustomXmlPart(testXMLpart_one),
        ];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(1);
        await expect(global.Excel.context.workbook.customXmlParts.items[0].getXml().value).toBe(testXMLpart_one);
    });

test('Excel set classification from empty',
    async () =>
    {
        global.Excel.context.workbook.properties.custom.items = [];
        global.Excel.context.workbook.customXmlParts.items = [];
        await taskpane.init(info);

        taskpane.classificationSelected("Default");

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(1);
        await expect(global.Excel.context.workbook.customXmlParts.items[0].getXml().value.includes('<attrValue xml:space="preserve">Default</attrValue>')).toBe(true);
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
        global.Excel.context.workbook.customXmlParts.items =
        [
            new CustomXmlPart(testXMLpart_one),
        ];
        await taskpane.init(info);

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(1);
        await expect(global.Excel.context.workbook.customXmlParts.items[0].getXml().value).toBe(testXMLpart_one);

        await taskpane.classificationSelected("Restricted");

        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Restricted");
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(1);
        await expect(global.Excel.context.workbook.customXmlParts.items[0].getXml().value.includes('<attrValue xml:space="preserve">Restricted</attrValue>')).toBe(true);
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
        global.Excel.context.workbook.customXmlParts.items =
        [
            new CustomXmlPart(testXMLpart_one),
        ];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(5);
        await expect(global.Excel.context.workbook.properties.custom.items[0].value).toBe("Default");
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(1);
        await expect(global.Excel.context.workbook.customXmlParts.items[0].getXml().value).toBe(testXMLpart_one);

        await taskpane.removeClassification();

        await expect(global.Excel.context.workbook.properties.custom.items.length).toBe(0);
        await expect(global.Excel.context.workbook.customXmlParts.items.length).toBe(0);
    });
