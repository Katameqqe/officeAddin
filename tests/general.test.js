global.Office =     require('./office');

const Window =      require('./window')

const customProp =  require('../src/taskpane/customProp')
global.initCustomProp = customProp.initCustomProp;

const fetch =       require('./fetch'); 

global.fetch = fetch.fetch;

const taskpane =    require('../src/taskpane/taskpane')

const Document =    require('./document');

test('helloWorld',
    () =>
    {
        global.window = new Window();
        global.document = new Document();

        const info = {host: Office.HostType.Word};
        taskpane.init(info);

        expect(true).toBe(true);
    });
