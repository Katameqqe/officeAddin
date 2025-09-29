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
            classificationDate.value,
            classificationGUID.value
        );
    }

    static async readByNameFromCustomXmlParts(aName, aCustomXmlParts)
    {
        if (aCustomXmlParts.items.length === 0)
        {
            return null;
        }

        for (const part of aCustomXmlParts.items)
        {
            const xml = part.getXml();
            await part.context.sync();

            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xml.value, "application/xml");

            const propName = xmlDoc.getElementsByTagName("customPropName")[0]?.textContent;
            if (propName === aName) {
                return {
                    value: xmlDoc.getElementsByTagName("attrValue")[0]?.textContent ?? "",
                    name: propName,
                    classificationDate: xmlDoc.getElementsByTagName("timestamp")[0]?.textContent ?? "",
                    classifiedBy: xmlDoc.getElementsByTagName("userName")[0]?.textContent ?? "",
                    classificationHost: xmlDoc.getElementsByTagName("computerName")[0]?.textContent ?? "",
                    classificationGUID: xmlDoc.getElementsByTagName("guid")[0]?.textContent ?? "",
                    id: part.id,
                };
            }
        }

        return null;
    }

    addClassificationInfo(customProperties)
    {
        customProperties.add(this.name, this.value);
        customProperties.add("ClassifiedBy", this.classifiedBy);
        customProperties.add("ClassificationHost", this.classificationHost);
        customProperties.add("ClassificationDate", this.classificationDate);
        customProperties.add("ClassificationGUID", this.classificationGUID);
    }

    toXmlString(hdrArr, ftrArr, wmObj)
    {
        function createR(fontName, fontColor, fontSize, text)
        {
            return `<r>
    <fontName>${fontName}</fontName>
    <fontColor>${fontColor}</fontColor>
    <fontSize>${fontSize}</fontSize>
    <b/>
    <text xml:space="preserve">${text}</text>
</r>`;
        }
        function addTwoTabsToEachLine(str) {
            return str.replace(/^/gm, '        ');
        }

        const hdrBlocks = hdrArr.map(obj =>
            createR(obj.fontName, obj.fontColor, obj.fontSize, obj.text)
        ).join('\n');
        const ftrBlocks = ftrArr.map(obj =>
            createR(obj.fontName, obj.fontColor, obj.fontSize, obj.text)
        ).join('\n');
        return `<GTBClassification>
    <attrValue xml:space="preserve">${this.value}</attrValue>
    <customPropName>${this.name}</customPropName>
    <timestamp>${this.classificationDate}</timestamp>
    <userName>${this.classifiedBy}</userName>
    <computerName>${this.classificationHost}</computerName>
    <guid>${this.classificationGUID}</guid>
    <hdr>${`\n` + addTwoTabsToEachLine(hdrBlocks)}
    </hdr>
    <ftr>${`\n` + addTwoTabsToEachLine(ftrBlocks)}
    </ftr>
    <wm>
        <fontName>${wmObj.fontName}</fontName>
        <fontColor>${wmObj.fontColor}</fontColor>
        <fontSize>${wmObj.fontSize}</fontSize>
        <b/>
        <rotation>${wmObj.rotation}</rotation>
        <transparency>${wmObj.transparency}</transparency>
        <text xml:space="preserve">${wmObj.text}</text>
    </wm>
</GTBClassification>`;
    }

    static deleteFromCustomProperties(aName, customProps)
    {
        const keysToRemove =
        [
            aName,
            "ClassifiedBy",
            "ClassificationHost",
            "ClassificationDate",
            "ClassificationGUID",
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

    constructor(aName, aValue, aUserName, aHostName, aDate = new Date().toLocaleString(), aGUID)
    {
        this.name = aName;
        this.value = aValue;
        this.classifiedBy = aUserName;
        this.classificationHost = aHostName;
        this.classificationDate = aDate;
        this.classificationGUID = aGUID;
    }
}
module.exports = CustomClassification;
