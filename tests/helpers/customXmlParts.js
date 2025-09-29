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

    async load(context, option)
    {
        return this.items;
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
    getItem(aId)
    {
        return this.items.find(item => item.id === aId);
    }
}
module.exports = CustomXmlParts;
