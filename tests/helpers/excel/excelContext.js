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

    async load(option)
    {
        await this.workbook.load(this, option);
    }
}
module.exports = ExcelContext;
