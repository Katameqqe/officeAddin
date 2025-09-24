class WordCustomXMLController
{
    constructor()
    {
    }

    async addCustomProperty(classificationObj)
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

                const FontArr = [{ fontName: "Arial", fontColor: "000000", fontSize: "14", text: "Sample Watermark" }, { fontName: "Verdana", fontColor: "FF0000", fontSize: "12", text: "Second Line" }];
                const wmObj = { fontName: "Arial", fontColor: "000000", fontSize: "36", rotation: "315", transparency: "0.5", text: classificationObj.value };
                const xmlString = classificationObj.toXmlString(FontArr, FontArr, wmObj);

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
