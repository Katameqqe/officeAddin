
global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/customProp')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/index')

global.WordCustomProp =             require('../src/WordCustomProp');
global.ExcelCustomProp =            require('../src/ExcelCustomProp');

const Document =                    require('./helpers/document');
const WordDocument =                require('./helpers/word/wordDocument');

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

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);

    });

test('Word onclick',
    async () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Word.context.document = new WordDocument();

        const info = {host: Office.HostType.Word, };
        await taskpane.init(info);
        const button = document.querySelector('input[value="Default"]');

        button.onchange();

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        console.log(global.Word.context.document.properties.customProperties.items);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
    });

// TODO: Display empty classification

// TODO: Display not empty classification

// TODO: set classification from empty

// TODO: update existed classification

// TODO: clear classification
