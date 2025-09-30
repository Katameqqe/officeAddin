class CustomPropertyController
{
    constructor(aHost)
    {
        this.PropertyController = new CustomPropertyProcessor();
        if (aHost === Office.HostType.Word)
        {
            this.PropertyController.documentType = "document";
            this.PropertyController.propertyName = "customProperties";
            this.platform = Word;
        }
        else if (aHost === Office.HostType.Excel)
        {
            this.PropertyController.documentType = "workbook";
            this.PropertyController.propertyName = "custom";
            this.platform = Excel;
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
        return this.platform.run(async (context) => {this.PropertyController.addCustomProperty(context, ClassificationObj)});
    }

    async readCustomProperty(name)
    {
        return this.platform.run(async (context) => {this.PropertyController.readCustomProperty(context, name)});
    }

    async removeCustomProperty(value)
    {
        return this.platform.run(async (context) => {this.PropertyController.removeCustomProperty(context, value)});
    };
}

module.exports = CustomPropertyController;
