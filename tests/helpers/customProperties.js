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
            if (customProperty.key === aName)
            {
                result = customProperty;
            }
        }
        return result;
    }

    add(aName, aValue)
    {
        if (this.getItemOrNullObject(aName).isNullObject)
        {
            this.items.push(new CustomProperty(aName, aValue))
        }
        else
        {
            this.getItemOrNullObject(aName).set(
                {
                    value: aValue, 
                });
        }
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

    find(aFunc)
    {
        for (let i = 0; i < this.items.length; i++)
        {
            const item = this.items[i];
            const res = aFunc(item);
            if (res)
            {
                return item;
            }
        }
        return null;
    }
}
module.exports = CustomProperties;
