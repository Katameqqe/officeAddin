global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/taskpane/customProp')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/taskpane/taskpane')

const Document =                    require('./helpers/document');
const WordDocument =                require('./helpers/word/wordDocument');

global.Word =                       require('./helpers/word/word')

test('helloWorld',
    async () =>
    {
        global.window = new Window();
        global.document = new Document();
        global.Word.context.document = new WordDocument();

        const info = {host: Office.HostType.Word, };
        await taskpane.init(info);

        expect(global.document.elements.app-body.children.length).toBe(5);
    });
