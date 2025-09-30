class CustomPropertyProcessor
{
    constructor(documentType)
    {
        this.documentType = documentType;
    }

    async addCustomProperty(context, classificationObj)
    {
        const customProps = context[this.documentType].properties[this.propertyName];
        customProps.load("items");
        await context.sync();

        classificationObj.addClassificationInfo(customProps);

        await context.sync();
        console.log(`Custom property "${JSON.stringify(classificationObj, null, 2)}" added.`);
    }

    async readCustomProperty(context, aName)
    {
        const customProps = context[this.documentType].properties[this.propertyName];
        customProps.load("items");
        await context.sync();

        const result = CustomClassification.readByNameFromCustomProperties(aName, customProps)

        if (result == null)
        {
            console.log(`One or more classification properties does not exist.`);
        }

        return result;
    }

    async removeCustomProperty(context, aName)
    {
        const customProps = context[this.documentType].properties[this.propertyName];
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
    }
}

module.exports = CustomPropertyProcessor;
