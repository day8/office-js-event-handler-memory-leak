/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  console.info("v2");
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function batch(context) {
  let sheet1 = context.workbook.worksheets.getItemOrNullObject("Sheet1");
  await context.sync({"sheet1": sheet1, "context": context});
  if (sheet1.isNullObject) {
    let sheet1 = context.workbook.worksheets.add("Sheet1");
    sheet1.visibility = Excel.SheetVisibility.hidden;
    await context.sync();
    console.info("Created new worksheet.")
  } else {
    console.info("Worksheet already exists.")
  }
}

var previousContext = null;

export async function run() {
  try {
    if (previousContext) {
      await Excel.run(previousContext, batch);
    } else {
      await Excel.run(async context => {
        previousContext = context;
        await batch(context);
      });
    }
  } catch (error) {
    console.error(error);
  }
}
