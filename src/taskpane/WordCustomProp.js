class WordCustomProp
{
    async addCustomProperty(name, value)
    {
        return Word.run(
            async (context) =>
            {
                const customProps = context.document.properties.customProperties;
                customProps.load("items");
                await context.sync();

                customProps.add(name, value);
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

                const exists = customProps.items.find(item => item.key === name)

                if (!exists)
                {
                    console.log(`Custom property "${name}" does not exist.`);
                    return null;
                }
                console.log("Custom Properties:", exists.value);
                return exists.value;
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

                const exists = customProps.items.find(item => item.key === name)
                
                if (exists)
                {
                    exists.delete();
                    await context.sync();
                    console.log(`Custom property "${name}" removed.`);
                }
                else
                {
                    console.log(`Custom property "${name}" does not exist.`);
                }
            });
    }
}

module.exports = WordCustomProp;