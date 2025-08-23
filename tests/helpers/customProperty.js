class CustomProperty
{
    constructor(aName, aValue)
    {
        this.name = aName;
        this.value = aValue;
        this.isNullObject = false;
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
