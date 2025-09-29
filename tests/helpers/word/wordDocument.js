const CustomXmlParts =      require("../customXmlParts");
const Properties =          require("../properties");
const CustomProperties =    require("../customProperties");

class WordDocument
{
    constructor()
    {
        this.customXmlParts = new CustomXmlParts();
        this.properties = new Properties();
        this.properties.customProperties = new CustomProperties();
    }

    async sync()
    {
        this.customXmlParts.sync();
        this.properties.customProperties.sync();
    }

    async load(context, option)
    {
        await this.customXmlParts.load(context, "items");
        await this.properties.customProperties.load();
    }
}
module.exports = WordDocument;
