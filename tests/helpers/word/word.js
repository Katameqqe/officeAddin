const WordContext = require("./wordContext");

class Word
{
    static context = new WordContext();

    static async run(aFunction)
    {
        return aFunction(Word.context);
    }
}
module.exports = Word;
