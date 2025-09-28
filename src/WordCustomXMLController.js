
class WordCustomXMLController
{
    constructor()
    {
    }

    // TODO: classifLabels in plural form, but it  uses as a single object. What is it? Is it a single object or an array of objects?
    async addCustomProperty(classificationObj, classifLabels)
    {
        return Word.run(
            async (context) =>
            {
                const xmlParts = context.document.customXmlParts;
                context.load(xmlParts, "items");
                await context.sync();

                const obj = await CustomClassification.readByNameFromCustomXmlParts(classificationObj.name, xmlParts);
                if (obj && obj.name == classificationObj.name)
                {
                    console.log(`Custom property XML "${classificationObj.name}" already exists.`);
                    xmlParts.getItem(obj.id).delete();
                }

                const xmlString = classificationObj.toXmlString(classifLabels.hdr, classifLabels.ftr, classifLabels.wm);

                xmlParts.add(xmlString);
                await context.sync();
                console.log(`Custom XML part "${JSON.stringify(classificationObj, null, 2)}" added.`);
            });
    }

    async readCustomProperty(aName)
    {
        return Word.run(
            async (context) =>
            {
                const xmlParts = context.document.customXmlParts;
                context.load(xmlParts, "items");
                await context.sync();
                const result = await CustomClassification.readByNameFromCustomXmlParts(aName, xmlParts)
                if (result == null)
                {
                    console.log(`One or more classification properties does not exist.`);
                }
                return result;
            });
    }

    async removeCustomProperty(aName)
    {
        return Word.run(
            async (context) =>
            {
                const xmlParts = context.document.customXmlParts;
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
            });
    }
}

module.exports = WordCustomXMLController;
