const MetaPrefix = "Classification";
const address = "https://192.168.128.4:443"
const HostName = "GTB.test.com";
const userName = "USERRR";
const GUID = "{123e4567-e89b-12d3-a456-426614174000}";

let classifLabels = null
let propertyController = null;
let XMLController = null;

Office.onReady(
    async (info) =>
    {
        console.log("*******************************************************");
        init(info);
    });

function isShouldProceed(anInfo)
{
    if (anInfo.host === Office.HostType.Word ||
        anInfo.host === Office.HostType.Excel ||
        (anInfo.host === Office.HostType.PowerPoint && Office.context.requirements.isSetSupported("PowerPointApi", "1.7")))
    {
        return true;
    }

    return false;
}

async function init(info)
{
    if (!isShouldProceed(info))
    {
        return;
    }

    propertyController = new CustomPropertyController(info.host);
    XMLController = new CustomXMLController(info.host);

    const ListSuffix = await getLabels();
    classifLabels = await getClassifLabels();


    // We better get classification from document before. And then "createButtons" with selected classification
    const classification = await propertyController.readCustomProperty(MetaPrefix);
    console.log(`Read custom property "${MetaPrefix}": ${JSON.stringify(classification, null, 2)}`);

    createButtons(ListSuffix, classification);
};

function createButtons(ListSuffix, aSelectedClassification)
{
    const clearSelected = aSelectedClassification == null;
    for (const suffix of ListSuffix)
    {
        let isSelected = false;
        if (!clearSelected)
        {
            isSelected = suffix === aSelectedClassification.value;
        }
        const node = generateClassificationItem(suffix, isSelected);
        document.getElementById("classificationGroup").appendChild(node);
    }
    const resetNode = clearClassificationItem(clearSelected);
    document.getElementById("classificationGroup").appendChild(resetNode);
}

// TODO: the function name "getLabels", but the api name is "list", function name and api name should be same.
async function getLabels()
{
    const List = await fetch(`${address}/api/v1/classification-labels`)
        .then(res => res.json())
        .then(resJson => resJson.names)
        .catch(
            err =>
            {
                console.error("Error fetching suffix list:", err);
                return ["Document", "Default", "Restricted", "Protected",];
            });

    console.log(JSON.stringify(List,null,2));
    return List;
}

async function getClassifLabels()
{
    const List = await fetch(`${address}/api/v1/xml-fonts`)
        .then(res => res.json())
        .then(resJson => resJson)
        .catch(
            err =>
            {
                // TODO: function name get classification labels, but in log "suffix".
                // What do we get or fetch? suffixes?
                console.error("Error fetching suffix list:", err);

                // TODO: It is bad practice to return such long something.
                // Move it to object, give it a name. What is it? Do we really need it?
                // If you have something default such long - move it to separate file.
                return defaultClassificationFont;
            });

    console.log(JSON.stringify(List,null,2));
    return List;
}

function generateClassificationItem(itemText, itemIsChecked)
{
    let isChecked = "";
    if (itemIsChecked)
    {
        isChecked = `checked="checked"`;
    }

    const itemHTML =
    `<div class="ms-ChoiceField">
        <label class="ms-ChoiceField-field">
            <input class="ms-ChoiceField-input" type="radio" name="classificationRadio" value="${itemText}" ${isChecked} onchange="classificationSelected(this.value);">
            <span class="ms-Label">${itemText}</span>
        </label>
    </div>`;

    const temp = document.createElement('div');
    temp.innerHTML = itemHTML;
    const node = temp.firstElementChild;
    return node;
}

function clearClassificationItem(itemIsChecked)
{
    let isChecked = "";
    if (itemIsChecked)
    {
        isChecked = `checked="checked"`;
    }

    const itemHTML =
    `
    <hr/>
    <div class="ms-ChoiceField">
        <label for="_clear_classification_" class="ms-ChoiceField-field">
            <input id="_clear_classification_" class="ms-ChoiceField-input" type="radio" name="classificationRadio" ${isChecked} onchange="removeClassification();">
            <span class="ms-Label">Not classified</span>
        </label>
    </div>`;

    const temp = document.createElement('div');
    temp.innerHTML = itemHTML;
    const node = temp;

    return node;
}

async function classificationSelected(aClassificationValue)
{
    let classificationObject = new CustomClassification(MetaPrefix, aClassificationValue, userName, HostName,new Date().toLocaleString(),GUID);
    XMLController.addCustomProperty(classificationObject, classifLabels);
    propertyController.addCustomProperty(classificationObject);
}

async function removeClassification()
{
    await propertyController.removeCustomProperty(MetaPrefix);
    await XMLController.removeCustomProperty(MetaPrefix);
}

module.exports.init = init;
module.exports.classificationSelected = classificationSelected;
module.exports.removeClassification = removeClassification;
