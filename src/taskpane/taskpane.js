/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { initCustomProp } from "./customProp";

const MetaPrefix = "Classification";
let ListSuffix = ["Document", "Default", "Restricted","Protected"];
ListSuffix.push("NoLabel");

function recolorButtons(activeButtonId, color = "#00FF00", defaultColor = "#000000") {
  const buttons = document.querySelectorAll(".meta-a_button");
  buttons.forEach(btn => {
    btn.style.color = (btn.id === activeButtonId) ? color : defaultColor;
  });
}
Office.onReady(async (info) => {
  window.INFO = info;
  initCustomProp();
  if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || (info.host === Office.HostType.PowerPoint && Office.context.requirements.isSetSupported("PowerPointApi", "1.7"))) {
    const res = await fetch('https://192.168.128.4:8000/list');
    const resJson = await res.json();
    ListSuffix = await resJson.names;
    console.log(JSON.stringify(ListSuffix,null,2));
    if(!ListSuffix.includes("NoLabel")){
      ListSuffix.push("NoLabel");
    }

    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runDocument;

    for (const suffix of ListSuffix) {
      const newButton = document.createElement("button");
      newButton.id = suffix;
      newButton.className = "meta-a_button ms-font-xl";
      newButton.textContent = suffix;
      newButton.onclick = () => {
        recolorButtons(newButton.id);
        addCustomProperty(MetaPrefix, suffix);
      };
      document.getElementById("app-body").appendChild(newButton);
    }
    document.getElementById("NoLabel").onclick = () => {
      recolorButtons("NoLabel");
      addCustomProperty(MetaPrefix, "", "NoLabel");
    };
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
