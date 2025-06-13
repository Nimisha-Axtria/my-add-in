/** 
 * This function runs when ribbon button is clicked.
 * It checks if Sheet2!Z1 has the trigger "RunColorChange" and acts accordingly.
 */
function checkAndRunColorChange(event) {
  Excel.run(async (context) => {
    // Get Sheet2 by index (index 1 = second sheet)
    const sheet = context.workbook.worksheets.getItemAt(1);
    const triggerCell = sheet.getRange("Z1");
    triggerCell.load("values");

    await context.sync();

    if (triggerCell.values[0][0] === "RunColorChange") {
      const range = context.workbook.getSelectedRange();
      range.load(["rowCount", "columnCount"]);

      await context.sync();

      const isSingleCell = range.rowCount === 1 && range.columnCount === 1;
      range.format.fill.color = isSingleCell ? "green" : "yellow";

      // Clear the trigger
      triggerCell.values = [[""]];
      await context.sync();
    } else {
      console.log("No trigger found in Sheet2!Z1");
    }
  })
  .catch(error => {
    console.error(error);
  })
  .finally(() => {
    event.completed();
  });
}

// Required to expose function to Office runtime
if (typeof window !== "undefined") {
  window.checkAndRunColorChange = checkAndRunColorChange;
} else {
  // Node environment or no window object
  global.checkAndRunColorChange = checkAndRunColorChange;
}
