
const MetaPrefix = "Classification";
const address = "https://192.168.128.4:443/list"
let propertyController = null;

Office.onReady(
    async (info) =>
    {
        console.log("*******************************************************");
        init(info);
    });

// TODO: Such a long function is a bad coding style. The code should be divided to separate functions and classes.
async function init(info)
{
    propertyController = new CustomPropertyController(info.host);

    if (info.host === Office.HostType.Word ||
        info.host === Office.HostType.Excel ||
        (info.host === Office.HostType.PowerPoint && Office.context.requirements.isSetSupported("PowerPointApi", "1.7")))
    {
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

        document.getElementById("app-body").style.display = "flex";
        //document.getElementById("run").onclick = runDocument;

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

        var prefixValue = await propertyController.readCustomProperty(MetaPrefix);

        console.log(`Read custom property "${MetaPrefix}": ${prefixValue}`);

        // TODO: If it is needed to write a comment - then it means that this code should be extracted to separate function
        // If the custom property value is in the list
        if (ListSuffix.includes(prefixValue))
        {
            console.log(`Custom property "${MetaPrefix}" exists with value: ${prefixValue}`);
            document.getElementById(prefixValue).classList.add("meta-button-active");
        }
        // If the custom property does not exist or is empty
        else if (prefixValue === null || prefixValue === "")
        {
            console.log(`Custom property "${MetaPrefix}" exists with value: "NoLabel"`);
            document.getElementById("NoLabel").classList.add("meta-button-active");
        }
        // If the custom property exists but is not in the list
        else
        {
            console.log(`Custom property "${MetaPrefix}" exists with value: ${prefixValue}`);
            const newButton = createButton(prefixValue)
            newButton.classList.add("meta-button-active");
            document.getElementById("app-body").insertBefore(newButton, document.getElementById("NoLabel"));
        };

        /*
document.getElementById("add-prop").onclick = () => {
  const value = document.getElementById("custom-prop-value").value;
  addCustomProperty(MetaPrefix, value, );
};
document.getElementById("remove-prop").onclick = () => {
  const value = document.getElementById("remove-prop-value").value;
  removeCustomProperty(value);
};
*/
    }
};

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
        };
    return button;
}

/*
export async function runDocument() {
  return Word.run(async (context) => {
    const properties = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    console.log(JSON.stringify(properties.items, null, 2));
});
}

export async function runWorkbook() {
  return Excel.run(async (context) => {
    const properties = context.workbook.properties.custom;
    properties.load("key,type,value");

    await context.sync();
    console.log(JSON.stringify(properties.items, null, 2));
});
}

export async function runPresentation() {
  return PowerPoint.run(async (context) => {
    const properties = context.presentation.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    console.log(JSON.stringify(properties.items, null, 2));
});
}
*/

module.exports.init = init;
