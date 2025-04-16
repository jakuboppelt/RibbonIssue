/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  console.log("[ddguo][taskpane.ts]: office ready");
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    try {
      Office.addin.showAsTaskpane();
    } catch (e) {
      console.error("Failed to open taskpane.");
      console.log(e);
    }
  }
});

export async function run() {
  console.log("[ddguo][taskpane.ts]: run is executed.");

  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

function runDemo(event: Office.AddinCommands.Event | null) {
  console.log("[ddguo][taskpane.ts]: command received");
  if (event) {
    try {
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = "green";
        await context.sync();
      });
    } catch (error) {
      // Note: In a production add-in, notify the user through your add-in's UI.
      console.error(error);
    }
    event.completed();
  }
}

// Register the function with Office.
Office.actions.associate("runDemo", runDemo);
