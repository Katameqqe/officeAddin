const CustomProperty = require("./customProperty");

class CustomProperties
{
    constructor()
    {
        this.items = [];
    }

    getItemOrNullObject(aName)
    {
        let result = { isNullObject: true, };
        for (let i = 0; i < this.items.length; i++)
        {
            let customProperty = this.items[i];
            if (customProperty.name === aName)
            {
                result = customProperty;
            }
        }
        return result;
    }

    add(aName, aValue)
    {
        this.items.push(new CustomProperty(aName, aValue))
    }

    load()
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
module.exports = CustomProperties;
