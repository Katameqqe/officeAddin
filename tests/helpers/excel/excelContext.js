class ExcelContext
{
    constructor()
    {
        this.workbook = null;
    }

    async sync()
    {
        this.workbook.sync();
    }
}
module.exports = ExcelContext;
