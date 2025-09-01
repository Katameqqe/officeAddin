class CustomPropertyController
{
    constructor(aHost)
    {
        if (aHost === Office.HostType.Word)
        {
            this.propertyController = new WordCustomPropertyController();
        }
        else if (aHost === Office.HostType.Excel)
        {
            this.propertyController = new WordCustomPropertyController();
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

    async addCustomProperty(ClassificationObj)
    {
        return this.propertyController.addCustomProperty(ClassificationObj);
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
