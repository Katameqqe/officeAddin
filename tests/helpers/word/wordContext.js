class WordContext
{
    constructor()
    {
        this.document = null;
    }

    async sync()
    {
        this.document.sync();
    }

    async load(option)
    {
        await this.document.load(this, option);
    }
}
module.exports = WordContext;
