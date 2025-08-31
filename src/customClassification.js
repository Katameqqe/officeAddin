class CustomClassification
{
    static readByNameFromCustomProperties(aName, aCustomProperties)
    {
        let result = null;

        const mainProperty = aCustomProperties.items.find(item => item.key === aName);

        if (mainProperty)
        {
            result = new CustomClassification(aName, mainProperty.value);
            result.classifiedBy = aCustomProperties.items.find(item => item.key === "ClassifiedBy");
            result.classificationHost = aCustomProperties.items.find(item => item.key === "ClassificationHost");
            result.classificationDate = aCustomProperties.items.find(item => item.key === "ClassificationDate");
            result.classificationGUID = aCustomProperties.items.find(item => item.key === "ClassificationGUID");
        }
        return result;
    }

    constructor(aName, aValue)
    {
        this.name = aName;
        this.value = aValue;
    }

}
module.exports = CustomClassification;
