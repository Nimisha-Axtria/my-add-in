/* global Excel, Office */

Office.onReady(() => {
  console.log("Office is ready");

  // Optional: Only run if you want polling logic when taskpane is shown
  startPollingForTrigger();
});

/**
 * Starts polling Z1 cell every 1 second for the "RunColorChange" trigger.
 */
function startPollingForTrigger() {
  setInterval(checkForTrigger, 1000);
}

/**
 * Checks Sheet2!Z1 for trigger value and runs color logic if matched.
 */
async function checkForTrigger() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemAt(1); // Sheet2
      const triggerCell = sheet.getRange("Z1");
      triggerCell.load("values");

      await context.sync();

      const trigger = triggerCell.values[0][0];
      if (trigger === "RunColorChange") {
        await runColorChange(context);

        // Clear the trigger
        triggerCell.values = [[""]];
        await context.sync();
      }
    });
  } catch (error) {
    console.error("Error checking for trigger:", error);
  }
}

/**
 * Applies conditional fill to the selected range.
 */
async function runColorChange(context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load(["rowCount", "columnCount"]);

  await context.sync();

  const isSingleCell = selectedRange.rowCount === 1 && selectedRange.columnCount === 1;
  selectedRange.format.fill.color = isSingleCell ? "green" : "yellow";
}
