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
}
module.exports = ExcelDocument;
