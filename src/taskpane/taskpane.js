
const MetaPrefix = "Classification";
const address = "https://192.168.128.4:443/list"
let propertyController = null;

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

    // TODO: getLabels function
    const ListSuffix = await fetch(address)
        .then(res => res.json())
        .then(resJson => resJson.names)
        .catch(
            err =>
            {
                console.error("Error fetching suffix list:", err);
                return ["Document", "Default", "Restricted", "Protected", "NoLabel",];
            });

    console.log(JSON.stringify(ListSuffix,null,2));

    if (!ListSuffix.includes("NoLabel"))
    {
        ListSuffix.push("NoLabel");
    }
    // ---------------------

    document.getElementById("app-body").style.display = "flex";

    // TODO: create buttons function
    for (const suffix of ListSuffix)
    {
        const newButton = createButton(suffix);
        document.getElementById("app-body").appendChild(newButton);
    }

    document.getElementById("NoLabel").onclick =
    () =>
    {
        recolorButtons("NoLabel");
        propertyController.addCustomProperty(MetaPrefix, "", "NoLabel");
    };
    // ----------------

    var prefixValue = await propertyController.readCustomProperty(MetaPrefix);

    console.log(`Read custom property "${MetaPrefix}": ${prefixValue}`);

    readClassif(document, ListSuffix, prefixValue);
};

function readClassif(document, ListSuffix, prefix)
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

module.exports.init = init;
