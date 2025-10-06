
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
global.RequestController =                  require('../src/RequestController');

const Document =                            require('./helpers/document');
const WordDocument =                        require('./helpers/word/wordDocument');
const CustomProperty =                      require('./helpers/customProperty');
const CustomXmlPart =                       require('./helpers/customXmlPart');


global.Word =                               require('./helpers/word/word');
const { DOMParser } =                       require('xmldom');

const info = {host: Office.HostType.Word, };

const testXMLpart_one =                     require('./helpers/exampleXML');

// TODO: implement tests

beforeEach(
    () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Word.context.document = new WordDocument();
        global.DOMParser = DOMParser;
    });

test('Word Display empty classification',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items = [];
        global.Word.context.document.customXmlParts.items = [];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(0);
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(0);
    });

test('Word Display not empty classification',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items =
        [
            new CustomProperty("Classification", "Default"),
            new CustomProperty("ClassifiedBy", "User"),
            new CustomProperty("ClassificationHost", "Word"),
            new CustomProperty("ClassificationDate", "Date"),
            new CustomProperty("ClassificationGUID", "GUID"),
        ];
        console.log(testXMLpart_one);
        global.Word.context.document.customXmlParts.items =
        [
            new CustomXmlPart(testXMLpart_one),
        ];

        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(1);
        await expect(global.Word.context.document.customXmlParts.items[0].getXml().value).toBe(testXMLpart_one);
    });

test('Word set classification from empty',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items = [];
        global.Word.context.document.customXmlParts.items = [];
        await taskpane.init(info);

        taskpane.classificationSelected("Default");

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(1);
        const xml = global.Word.context.document.customXmlParts.items[0].getXml().value;
        await expect(xml.includes(`<attrValue xml:space="preserve">Default</attrValue>`)).toBe(true);
    });

test('Word update existed classification',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items =
        [
            new CustomProperty("Classification", "Default"),
            new CustomProperty("ClassifiedBy", "User"),
            new CustomProperty("ClassificationHost", "Word"),
            new CustomProperty("ClassificationDate", "Date"),
            new CustomProperty("ClassificationGUID", "GUID"),
        ];
        global.Word.context.document.customXmlParts.items =
        [
            new CustomXmlPart(testXMLpart_one),
        ];
        await taskpane.init(info);

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(1);
        await expect(global.Word.context.document.customXmlParts.items[0].getXml().value).toBe(testXMLpart_one);

        await taskpane.classificationSelected("Restricted");

        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Restricted");
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(1);
        
        const xml = global.Word.context.document.customXmlParts.items[0].getXml().value;
        await expect(xml.includes(`<attrValue xml:space="preserve">Restricted</attrValue>`)).toBe(true);
    });

test('Word clear classification',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items =
        [
            new CustomProperty("Classification", "Default"),
            new CustomProperty("ClassifiedBy", "User"),
            new CustomProperty("ClassificationHost", "Word"),
            new CustomProperty("ClassificationDate", "Date"),
            new CustomProperty("ClassificationGUID", "GUID"),
        ];
        global.Word.context.document.customXmlParts.items =
        [
            new CustomXmlPart(testXMLpart_one),
        ];

        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(1);

        await taskpane.removeClassification();

        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(0);
        await expect(global.Word.context.document.customXmlParts.items.length).toBe(0);
    });
