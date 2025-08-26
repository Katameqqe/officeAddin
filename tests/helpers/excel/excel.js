const ExcelContext = require("./excelContext");

class Excel
{
    static context = new ExcelContext();

    static async run(aFunction)
    {
        return aFunction(Excel.context);
    }
}
module.exports = Excel;
