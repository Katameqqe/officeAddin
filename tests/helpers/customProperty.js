class CustomProperty
{
    constructor(aName, aValue)
    {
        this.key = aName;
        this.value = aValue;
        this.isNullObject = false;
        this.toDelete = false;
    }

    set(aData)
    {
        this.value = aData.value;
    }

    load()
    {

    }

    delete()
    {
        this.toDelete = true;
    }

}
module.exports = CustomProperty;
