class CustomXmlPart
{
    constructor(anXml)
    {
        this.id = "customXmlPartId"
        this.xml = anXml;
        this.context = {};
        this.context.sync = async () => {};
    }

    getXml()
    {
        return { value: this.xml, };
    }

    setXml(anXml)
    {
        this.xml = anXml;
    }

    delete()
    {
        this.toDelete = true;
    }
}
module.exports = CustomXmlPart;
