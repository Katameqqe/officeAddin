class CustomPropertyController
{
    constructor(aHost, userName, HostName, GUID)
    {
        if (aHost === Office.HostType.Word)
        {
            this.propertyController = new WordCustomProp(userName, HostName, GUID);
        }
        else if (aHost === Office.HostType.Excel)
        {
            this.propertyController = new ExcelCustomProp(userName, HostName, GUID);
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

module.exports = CustomPropertyController;
