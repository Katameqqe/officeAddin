const CustomXmlPart = require("./customXmlPart");

class CustomXmlParts
{
    constructor()
    {
        this.items = [ ];
    }

    add(anXml)
    {
        this.items.push(new CustomXmlPart(anXml));
    }

    async load()
    {

    }

    async sync()
    {
        let newItems = [];
        for (let i = 0; i < this.items.length; i++)
        {
            const item = this.items[i];
            if (!item.toDelete)
            {
                newItems.push(item);
            }
        }
        this.items = newItems;
    }
}
module.exports = CustomXmlParts;
