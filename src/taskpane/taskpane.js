/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  console.info("v1");
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function onWorksheetAdded(event) {
  let context = event.context;
  let worksheetId = event.worksheetId;
  console.info("Added worksheet ID:", worksheetId);
  try {
    await Excel.run(context, async context => {
      let addedWorksheet = context.workbook.worksheets.getItemOrNullObject(worksheetId);
      console.log("here", worksheetId);
      addedWorksheet.load("name");
      await context.sync();
      if (!addedWorksheet.isNullObject) {
        console.info("Added worksheet name:", addedWorksheet.name);
        addedWorksheet.visibility = Excel.SheetVisibility.hidden;
        addedWorksheet.delete();
        await context.sync();
      } else {
        console.warn("Added worksheet is null object.");
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run() {
  try {
    await Excel.run(async context => {
      context.workbook.worksheets.onAdded.add(onWorksheetAdded);
      await context.sync();
      console.info("worksheets.onAdded event handler added.")
    });
  } catch (error) {
    console.error(error);
  }
}
