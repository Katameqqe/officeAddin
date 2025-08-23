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
}
module.exports = WordContext;
