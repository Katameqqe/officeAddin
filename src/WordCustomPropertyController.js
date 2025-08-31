class WordCustomPropertyController
{
    constructor(userName, HostName, GUID)
    {
        this.userName = userName;
        this.HostName = HostName;
        this.GUID = GUID;
    }

    async addCustomProperty(name, value)
    {
        return Word.run(
            async (context) =>
            {
                const customProps = context.document.properties.customProperties;
                customProps.load("items");
                await context.sync();

                customProps.add(name, value);
                customProps.add("ClassifiedBy", this.userName);
                customProps.add("ClassificationHost", this.HostName);
                customProps.add("ClassificationDate", new Date().toLocaleString());
                customProps.add("ClassificationGUID", this.GUID);
                await context.sync();
                console.log(`Custom property "${name}" added with value: ${value}`);
            });
    }

    async readCustomProperty(name)
    {
        return Word.run(
            async (context) =>
            {
                const customProps = context.document.properties.customProperties;
                customProps.load("items");
                await context.sync();

                const mainProp = customProps.items.find(item => item.key === name);

                const classifiedBy = customProps.items.find(item => item.key === "ClassifiedBy");
                const classificationHost = customProps.items.find(item => item.key === "ClassificationHost");
                const classificationDate = customProps.items.find(item => item.key === "ClassificationDate");
                const classificationGUID = customProps.items.find(item => item.key === "ClassificationGUID");

                if (mainProp && classifiedBy && classificationHost && classificationDate && classificationGUID)
                {
                    return {
                        [name]:             mainProp.value,
                        ClassifiedBy:       classifiedBy.value,
                        ClassificationHost: classificationHost.value,
                        ClassificationDate: classificationDate.value,
                        ClassificationGUID: classificationGUID.value
                    };
                }
                else
                {
                    console.log(`One or more classification properties do not exist.`);
                    return 
                    {
                        [name]: null
                    };
                }
            });
    }

    async removeCustomProperty(name)
    {
        return Word.run(
            async (context) =>
            {
                const customProps = context.document.properties.customProperties;
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

module.exports = WordCustomProp;
