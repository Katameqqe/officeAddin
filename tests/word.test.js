
global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/CustomPropertyController')
global.CustomClassification =       require('../src/customClassification')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/index')

global.WordCustomPropertyController = require('../src/WordCustomPropertyController');

const Document =                    require('./helpers/document');
const WordDocument =                require('./helpers/word/wordDocument');
const CustomProperty = require('./helpers/customProperty');

global.Word =                       require('./helpers/word/word')
global.Excel =                      require('./helpers/excel/excel')
const info = {host: Office.HostType.Word, };

beforeEach(() => {
    global.window = new Window();
    global.document = new Document();
    global.Word.context.document = new WordDocument();
});

// TODO: Display empty classification
test('Word Display empty classification',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items = [];
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(0);
    });

// TODO: Display not empty classification
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
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
    });

// TODO: set classification from empty
test('Word set classification from empty',
    async () =>
    {
        global.Word.context.document.properties.customProperties.items = [];
        await taskpane.init(info);

        taskpane.classificationSelected("Default");

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
    });

// TODO: update existed classification
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
        await taskpane.init(info);
        
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
        
        await taskpane.classificationSelected("Restricted");

        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Restricted");
    });

// TODO: clear classification
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
        await taskpane.init(info);
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(5);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
        
        await taskpane.removeClassification();
        
        await expect(global.Word.context.document.properties.customProperties.items.length).toBe(0);
    });