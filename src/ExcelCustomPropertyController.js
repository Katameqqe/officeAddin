
class ExcelCustomPropertyController
{
    constructor()
    {
    }

    async addCustomProperty(name, value)
    {
        return Excel.run(
            async (context) =>
            {
                const customProperties = context.workbook.properties.custom;
                customProperties.load("items");
                await context.sync();

                customProperties.add(name, value);
                customProperties.add("ClassifiedBy", this.userName);
                customProperties.add("ClassificationHost", this.HostName);
                customProperties.add("ClassificationDate", new Date().toLocaleString());
                customProperties.add("ClassificationGUID", this.GUID);
                await context.sync();
                console.log(`Custom property "${name}" added with value: ${value}`);
            });
    }

    async readCustomProperty(aName)
    {
        return Excel.run(
            async (context) =>
            {
                const customProperties = context.workbook.properties.custom;
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

    async removeCustomProperty(name)
    {
        return Excel.run(
            async (context) =>
            {
                const customProps = context.workbook.properties.custom;
                customProps.load("items");
                await context.sync();

                const mainProp = customProps.items.find(item => item.key === name)

                const classifiedBy = customProps.items.find(item => item.key === "ClassifiedBy");
                const classificationHost = customProps.items.find(item => item.key === "ClassificationHost");
                const classificationDate = customProps.items.find(item => item.key === "ClassificationDate");
                const classificationGUID = customProps.items.find(item => item.key === "ClassificationGUID");

                if (mainProp && classifiedBy && classificationHost && classificationDate && classificationGUID)
                {
                    mainProp.delete();
                    classifiedBy.delete();
                    classificationHost.delete();
                    classificationDate.delete();
                    classificationGUID.delete();
                    await context.sync();
                    console.log(`Custom property "${name}" and related classification properties removed.`);
                }
                else
                {
                    console.log(`Custom property "${name}" does not exist.`);
                }
            });
    }
}

module.exports = ExcelCustomPropertyController;
