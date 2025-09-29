class CustomXMLController 
{
    constructor(aHost)
    {
        this.XMLcontroller = new CustomXMLProcessor();
        if (aHost === Office.HostType.Word)
        {
            this.contextType = "document";
            this.platform = Word;
        }
        else if (aHost === Office.HostType.Excel)
        {
            this.contextType = "workbook";
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
        return this.platform.run(async (context) => {this.XMLcontroller.addCustomProperty(context[this.contextType], ClassificationObj)});
    }

    async readCustomProperty(name)
    {
        return this.platform.run(async (context) => {this.XMLcontroller.readCustomProperty(context[this.contextType], name)});
    }

    async removeCustomProperty(value)
    {
        return this.platform.run(async (context) => {this.XMLcontroller.removeCustomProperty(context[this.contextType], value)});
    };

}

module.exports = CustomXMLController;