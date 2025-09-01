class CustomClassification
{
    static readByNameFromCustomProperties(aName, aCustomProperties)
    {
        const mainProp = aCustomProperties.items.find(item => item.key === aName);
        const classifiedBy = aCustomProperties.items.find(item => item.key === "ClassifiedBy");
        const classificationHost = aCustomProperties.items.find(item => item.key === "ClassificationHost");
        const classificationDate = aCustomProperties.items.find(item => item.key === "ClassificationDate");
        const classificationGUID = aCustomProperties.items.find(item => item.key === "ClassificationGUID");

        if (!mainProp || !classifiedBy || !classificationHost || !classificationDate || !classificationGUID)
        {
            return null;
        }

        return new CustomClassification
        (
            aName,
            mainProp.value,
            classifiedBy.value,
            classificationHost.value,
            classificationGUID.value
        );
    }

    addClassificationInfo(customProperties)
    {
        customProperties.add(this.name, this.value);
        customProperties.add("ClassifiedBy", this.classifiedBy);
        customProperties.add("ClassificationHost", this.classificationHost);
        customProperties.add("ClassificationDate", this.classificationDate);
        customProperties.add("ClassificationGUID", this.classificationGUID);
    }

    static deleteFromCustomProperties(aName, customProps)
    {
        const keysToRemove =
        [
            aName,
            "ClassifiedBy",
            "ClassificationHost",
            "ClassificationDate",
            "ClassificationGUID"
        ];
        for (const key of keysToRemove)
        {
            const prop = customProps.items.find(item => item.key === key);
            if (prop)
            {
                prop.delete();
            }
        }
    }

    constructor(aName, aValue, aUserName, aHostName, aGUID)
    {
        this.name = aName;
        this.value = aValue;
        this.classifiedBy = aUserName;
        this.classificationHost = aHostName;
        this.classificationDate = new Date().toLocaleString();
        this.classificationGUID = aGUID;
    }

}
module.exports = CustomClassification;
