/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function runDemo(event: Office.AddinCommands.Event | null) {
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
}

// /* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.log("[ddguo][commands.ts]: office ready");
});

// Register the function with Office.
Office.actions.associate("runDemo", runDemo);

console.log("registered!!!§");
