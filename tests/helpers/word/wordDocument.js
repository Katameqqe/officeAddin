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
}
module.exports = WordDocument;
