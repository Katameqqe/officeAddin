class CustomPropertyController
{
    constructor(aHost)
    {
        if (aHost === Office.HostType.Word)
        {
            this.propertyController = new WordCustomProp();
        }
        else if (aHost === Office.HostType.Excel)
        {
            this.propertyController = new ExcelCustomProp();
        }
        else if (aHost === Office.HostType.PowerPoint)
        {
            //TODO: Implement PowerPoint implementation
            console.error("Unsupported host application.");
        }
        else
        {
            console.error("Unsupported host application.");
        }
    }

    async addCustomProperty(name, value)
    {
        return this.propertyController.addCustomProperty(name, value);
    }

    async readCustomProperty(name)
    {
        return this.propertyController.readCustomProperty(name);
    }

    async removeCustomProperty(value)
    {
        return this.propertyController.removeCustomProperty(value);
    };
}

// TODO: this file should be divided to separate files. One class - one file
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

class ExcelCustomProp
{
    async addCustomProperty(name, value)
    {
        return Excel.run(
            async (context) =>
            {
                const customProps = context.workbook.properties.custom;
                customProps.load("items");
                await context.sync();

                customProps.add(name, value);
                await context.sync();
                console.log(`Custom property "${name}" added with value: ${value}`);
            });
    }

    async readCustomProperty(name)
    {
        return Excel.run(
            async (context) =>
            {
                const customProps = context.workbook.properties.custom;
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
        return Excel.run(
            async (context) =>
            {
                const customProps = context.workbook.properties.custom;
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

module.exports = CustomPropertyController;
