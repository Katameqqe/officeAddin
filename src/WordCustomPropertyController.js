class WordCustomPropertyController
{
    constructor()
    {
    }

    async addCustomProperty(classificationObj)
    {
        return Word.run(
            async (context) =>
            {
                const customProps = context.document.properties.customProperties;
                customProps.load("items");
                await context.sync();

                classificationObj.addClassificationInfo(customProps);

                await context.sync();
                console.log(`Custom property "${JSON.stringify(classificationObj, null, 2)}" added.`);
            });
    }

    async readCustomProperty(aName)
    {
        return Word.run(
            async (context) =>
            {
                const customProperties = context.document.properties.customProperties;
                customProperties.load("items");
                await context.sync();

                const result = CustomClassification.readByNameFromCustomProperties(aName, customProperties)

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
                const customProps = context.document.properties.customProperties;
                customProps.load("items");
                await context.sync();

                const classificationObj = CustomClassification.readByNameFromCustomProperties(aName, customProps);

                if (classificationObj)
                {
                    CustomClassification.deleteFromCustomProperties(aName, customProps);
                    await context.sync();
                    console.log(`Custom property "${aName}" and related classification properties removed.`);
                }
                else
                {
                    console.log(`Custom property "${aName}" does not exist.`);
                }
            });
    }
}

module.exports = WordCustomPropertyController;
