class CustomXMLProcessor
{
    constructor(documentType)
    {
        this.documentType = documentType;
    }

    async addCustomProperty(context, classificationObj)
    {
        const xmlParts = context[this.documentType].customXmlParts;
        context.load(xmlParts, "items");
        await context.sync();

        const obj = await CustomClassification.readByNameFromCustomXmlParts(classificationObj.name, xmlParts);
        if (obj && obj.name == classificationObj.name)
        {
            console.log(`Custom property XML "${classificationObj.name}" already exists.`);
            xmlParts.getItem(obj.id).delete();
        }
        
        const xmlString = classificationObj.toXmlString();

        xmlParts.add(xmlString);
        await context.sync();
        console.log(`Custom XML part "${JSON.stringify(classificationObj, null, 2)}" added.`);
    }

    async readCustomProperty(context, aName)
    {
        const xmlParts = context[this.documentType].customXmlParts;
        context.load(xmlParts, "items");
        await context.sync();
        const result = await CustomClassification.readByNameFromCustomXmlParts(aName, xmlParts)
        if (result == null)
        {
            console.log(`One or more classification properties does not exist.`);
        }
        return result;
    }

    async removeCustomProperty(context, aName)
    {
        const xmlParts = context[this.documentType].customXmlParts;
        context.load(xmlParts, "items");
        await context.sync();
        const classificationObj = await CustomClassification.readByNameFromCustomXmlParts(aName, xmlParts);
        if (classificationObj)
        {
            xmlParts.getItem(classificationObj.id).delete();
        }
        else
        {
            console.log(`Custom property "${aName}" does not exist.`);
        }
        await context.sync();
    }
}

module.exports = CustomXMLProcessor;