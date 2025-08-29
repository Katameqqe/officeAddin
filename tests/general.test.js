/**
 * @jest-environment jsdom
 */


global.Office =                     require('./helpers/office');

const Window =                      require('./helpers/window')

global.CustomPropertyController =   require('../src/taskpane/customProp')

global.fetch =                      require('./helpers/fetch');;

const taskpane =                    require('../src/taskpane/taskpane')
taskpane.setDebug(true);

global.WordCustomProp =             require('../src/taskpane/WordCustomProp');
global.ExcelCustomProp =            require('../src/taskpane/ExcelCustomProp');

const Document =                    require('./helpers/document');
const WordDocument =                require('./helpers/word/wordDocument');
const excelWorkbook =               require('./helpers/excel/excelWorkbook');

global.Word =                       require('./helpers/word/word')
global.Excel =                      require('./helpers/excel/excel')

import fs from "fs";
import path from "path";

beforeEach(
    () => 
    {
        const html = fs.readFileSync(path.resolve(__dirname, "../src/taskpane/taskpane.html"), "utf8");
        document.documentElement.innerHTML = html;
    });

test('Word length of array',
    async () =>
    {
        //global.window = new Window();
        //global.document = new Document();
        global.Word.context.document = new WordDocument();

        const info = {host: Office.HostType.Word, };
        await taskpane.init(info);

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);

    });

test('Excel test length of array',
    async () =>
    {
        //global.window = new Window();
        //global.document = new Document();
        global.Excel.context.workbook = new excelWorkbook();

        const info = {host: Office.HostType.Excel, };
        await taskpane.init(info);

        await expect(document.getElementById("classificationGroup").children.length).toBe(5);

    });

test('Word onclick',
    async () =>
    {
        //global.window = new Window();
        //global.document = new Document();
        global.Word.context.document = new WordDocument();

        const info = {host: Office.HostType.Word, };
        await taskpane.init(info);
        const button = document.querySelector('input[value="Default"]');
        
        button.onchange();
        
        await expect(document.getElementById("classificationGroup").children.length).toBe(5);
        console.log(global.Word.context.document.properties.customProperties.items);
        await expect(global.Word.context.document.properties.customProperties.items[0].value).toBe("Default");
    });