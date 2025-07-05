/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const MetaPrefix = "MyCustomProp";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("add-prop").onclick = () => {
      const value = document.getElementById("custom-prop-value").value;
      addCustomProperty(MetaPrefix, value);
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

export async function addCustomProperty(name, value) {
  return Word.run(async (context) => {
    const customProps = context.document.properties.customProperties;
    customProps.load("items");
    await context.sync();

    // Check if property already exists
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
