/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const MetaPrefix = "Classification";
const ListSuffix = ["Document", "Default", "Restricted","Protected"];
ListSuffix.push("NoLabel");

function recolorButtons(activeButtonId, color = "#00FF00", defaultColor = "#000000") {
  const buttons = document.querySelectorAll(".meta-a_button");
  buttons.forEach(btn => {
    btn.style.color = (btn.id === activeButtonId) ? color : defaultColor;
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    for (const suffix of ListSuffix) {
      const newButton = document.createElement("button");
      newButton.id = suffix;
      newButton.className = "meta-a_button ms-font-xl";
      newButton.textContent = suffix;
      newButton.onclick = () => {
        recolorButtons(newButton.id);
        addCustomProperty(MetaPrefix, suffix, newButton.id);
      };
      document.getElementById("app-body").appendChild(newButton);
    }
    document.getElementById("NoLabel").onclick = () => {
      recolorButtons("NoLabel");
      addCustomProperty(MetaPrefix, "", "NoLabel");
    };
    (async () => {
      var rb = await readCustomProperty(MetaPrefix);
      console.log(`Read custom property "${MetaPrefix}": ${rb}`);
      if (ListSuffix.includes(rb)) {
        console.log(`Custom property "${MetaPrefix}" exists with value: ${rb}`);
        document.getElementById(rb).style.color = "#00FF00";
      } else if (rb === null || rb === "") {
        console.log(`Custom property "${MetaPrefix}" exists with value: "NoLabel"`);
        document.getElementById("NoLabel").style.color = "#00FF00";
      } else {
        const newButton = document.createElement("button");
        newButton.id = rb;
        newButton.className = "meta-a_button ms-font-xl";
        newButton.textContent = rb;
        newButton.style.color = "#00FF00";
        newButton.onclick = () => {
          recolorButtons(newButton.id);
          addCustomProperty(MetaPrefix, rb, newButton.id);
        };
        const referenceElement = document.getElementById("NoLabel");
        document.getElementById("app-body").insertBefore(newButton, referenceElement);
      };
    })();
    document.getElementById("add-prop").onclick = () => {
      const value = document.getElementById("custom-prop-value").value;
      addCustomProperty(MetaPrefix, value, );
    };
    document.getElementById("remove-prop").onclick = () => {
      const value = document.getElementById("remove-prop-value").value;
      removeCustomProperty(value);
    };
  }
});

export async function removeCustomProperty(name) {
  return Word.run(async (context) => {
    const customProps = context.document.properties.customProperties;
    customProps.load("items");
    await context.sync();

    const propToRemove = customProps.items.find(prop => prop.key === name);
    if (propToRemove) {
      propToRemove.delete();
      await context.sync();
      console.log(`Custom property "${name}" removed.`);
    } else {
      console.log(`Custom property "${name}" not found.`);
    }
  });
}

export async function readCustomProperty(name) {
  return Word.run(async (context) => {
    const customProps = context.document.properties.customProperties;
    customProps.load("items");
    await context.sync();

    const exists = customProps.items.find(prop => prop.key === name);
    if (exists) {
      console.log(`Custom property "${name}" updated with value: ${JSON.stringify(exists.value)}`);
      return exists.value;
    } else {
      console.log(`Custom property "${name}" not found.`);
      return null;
    }
  });
}

export async function addCustomProperty(name, value, id) {
  return Word.run(async (context) => {
    const customProps = context.document.properties.customProperties;
    customProps.load("items");
    await context.sync();

    const exists = customProps.items.find(prop => prop.key === name);
    if (exists) {
      exists.value = value; 
      await context.sync();
      console.log(`Custom property "${name}" updated with value: ${value}`);
    } else {  
      customProps.add(name, value);
      await context.sync();
      console.log(`Custom property "${name}" added with value: ${value}`);
    }
  });
}

export async function run() {
  return Word.run(async (context) => {
    const properties = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    console.log(JSON.stringify(properties.items, null, 2));
});
}
