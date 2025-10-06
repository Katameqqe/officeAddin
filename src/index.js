const MetaPrefix = "Classification";
const address = "https://192.168.128.4:443"
const HostName = "GTB.test.com";
const userName = "USERRR";
const GUID = "{123e4567-e89b-12d3-a456-426614174000}";

let ClassificationFonts = null;
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

    reqCtrl = new RequestController();

    console.log(`RequestController address: ${JSON.stringify(await reqCtrl.getClassifcationLabels())}`);

    const ClassificationLabels = await reqCtrl.getClassifcationLabels();
    ClassificationFonts = await reqCtrl.getClassificationFonts();


    // We better get classification from document before. And then "createButtons" with selected classification
    const classification = await propertyController.readCustomProperty(MetaPrefix);
    const classificationFonts = await XMLController.readCustomProperty(MetaPrefix);
    console.log(`Read custom property "${MetaPrefix}": ${JSON.stringify(classification, null, 2)}`);

    createButtons(ClassificationLabels, classification);
};

function createButtons(ClassificationLabels, aSelectedClassification)
{
    const clearSelected = aSelectedClassification == null;
    for (const suffix of ClassificationLabels)
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
    let classificationObject = new CustomClassification(MetaPrefix, aClassificationValue, userName, HostName,new Date().toLocaleString(),GUID, ClassificationFonts);
    XMLController.addCustomProperty(classificationObject);
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
