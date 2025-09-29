const CustomXmlParts =      require("../customXmlParts");
const Properties =          require("../properties");
const CustomProperties =    require("../customProperties");

class ExcelDocument
{
    constructor()
    {
        this.customXmlParts = new CustomXmlParts();
        this.properties = new Properties();
        this.properties.custom = new CustomProperties();
    }

    async sync()
    {
        this.customXmlParts.sync();
        this.properties.custom.sync();
    }

    async load(context, option)
    {
        await this.customXmlParts.load(context, "items");
        await this.properties.custom.load();
    }
}
module.exports = ExcelDocument;
