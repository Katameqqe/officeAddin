
const MetaPrefix = "Classification";
const address = "https://192.168.128.4:443/list"
let propertyController = null;
const isButtons = false; // false - radio buttons, true - buttons

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

// TODO: Such a long function is a bad coding style. The code should be divided to separate functions and classes.
async function init(info)
{
    if (!isShouldProceed(info))
    {
        return;
    }

    propertyController = new CustomPropertyController(info.host);
    global.propertyController = propertyController; // or window.propertyController in browser
    global.MetaPrefix = MetaPrefix;
    const ListSuffix = await getLabels();
    createButtons(ListSuffix);

    var classifValue = await propertyController.readCustomProperty(MetaPrefix);

    console.log(`Read custom property "${MetaPrefix}": ${classifValue}`);

    readClassif(ListSuffix, classifValue);
};

function readClassifButton(ListSuffix, prefix)
{
    // If the custom property value is in the list
    if (ListSuffix.includes(prefix))
    {
        console.log(`Custom property "${MetaPrefix}" exists with value: ${prefix}`);
        document.getElementById(prefix).classList.add("meta-button-active");
    }
    // If the custom property does not exist or is empty
    else if (prefix === null || prefix === "")
    {
        console.log(`Custom property "${MetaPrefix}" exists with value: "NoLabel"`);
        document.getElementById("NoLabel").classList.add("meta-button-active");
    }
    //TODO: If custom property exists, but is not in the list - then clear custom property for document and select "No Label"
    // If the custom property exists but is not in the list
    else
    {
        console.log(`Custom property "${MetaPrefix}" exists with value: ${prefix}`);
        const newButton = createButton(prefix)
        newButton.classList.add("meta-button-active");
        document.getElementById("app-body").insertBefore(newButton, document.getElementById("NoLabel"));
    };
}
function readClassif(ListSuffix, suffix)
{
    if (ListSuffix.includes(suffix))
    {
        console.log(`Custom property "${MetaPrefix}" exists with value: ${suffix}`);
        document.querySelector(`input[value="${suffix}"]`).checked = true;
    }
    else 
    {
        console.log(`Custom property "${MetaPrefix}" exists out of list with value: ${suffix}`);
        propertyController.removeCustomProperty(MetaPrefix);
        document.getElementById("_clear_classification_").checked = true;
    }
}

function recolorButtons(activeButtonId, color = "meta-button-active")
{
    const buttons = document.querySelectorAll(".pref-button");

    buttons.forEach(
        btn =>
        {
            if (btn.id === activeButtonId)
            {
                btn.classList.add(color);
            }
            else
            {
                btn.classList.remove(color);
            }
        });
}

function createButton(suffix = "")
{
    const button = document.createElement("button");
    button.id = suffix;
    button.className = "pref-button meta-button";
    button.textContent = suffix;
    button.onclick =
        () =>
        {
            recolorButtons(button.id);
            propertyController.addCustomProperty(MetaPrefix, suffix);
            //console.log(`Button "${button.id}" clicked.`);
        };
    return button;
}

function createButtons(ListSuffix)
{
    if (isButtons){
        for (const suffix of ListSuffix)
        {
            const newButton = createButton(suffix);
            document.getElementById("app-body").appendChild(newButton);
        }
    } else {
        for (const suffix of ListSuffix)
        {
            const node = generateClassificationItem(suffix, false);
            document.getElementById("classificationGroup").appendChild(node);
        }
        const resetNode = clearClassificationItem(false);
        document.getElementById("classificationGroup").appendChild(resetNode);
    }
}

async function getLabels()
{
    const List = await fetch(address)
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
            <input class="ms-ChoiceField-input" type="radio" name="classificationRadio" value="${itemText}" ${isChecked} onchange="propertyController.addCustomProperty(MetaPrefix, this.value);">
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
            <input id="_clear_classification_" class="ms-ChoiceField-input" type="radio" name="classificationRadio" ${isChecked} onchange="propertyController.removeCustomProperty(MetaPrefix)">
            <span class="ms-Label">Not classified</span>
        </label>
    </div>`;

    const temp = document.createElement('div');
    temp.innerHTML = itemHTML;
    const node = temp;
    return node;
}

module.exports.init = init;
