
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

function initCustomProp()
{
    if (window.INFO.host === Office.HostType.Word)
    {
        const wordCustomProp = new WordCustomProp();
        window.addCustomProperty =
            async (name, value) =>
            {
                await wordCustomProp.addCustomProperty(name, value);
            };

        window.readCustomProperty =
            async (name) =>
            {
                return await wordCustomProp.readCustomProperty(name);
            };

        window.removeCustomProperty =
            async (value) =>
            {
                await wordCustomProp.removeCustomProperty(value);
            };
    }
    else if (window.INFO.host === Office.HostType.Excel)
    {
        const excelCustomProp = new ExcelCustomProp();

        window.addCustomProperty =
            async (name, value) =>
            {
                await excelCustomProp.addCustomProperty(name, value);
            };

        window.readCustomProperty =
            async (name) =>
            {
                return await excelCustomProp.readCustomProperty(name);
            };

        window.removeCustomProperty =
            async (value) =>
            {
                await excelCustomProp.removeCustomProperty(value);
            };
    }
    else if (window.INFO.host === Office.HostType.PowerPoint)
    {
        //TODO: Implement PowerPoint implementation
    }
    else
    {
        console.error("Unsupported host application.");
    }
}

module.exports.initCustomProp = initCustomProp;
