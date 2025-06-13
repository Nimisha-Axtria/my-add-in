Office.onReady(() => {
  document.getElementById("app-body").style.display = "block";

  // Start polling every 2 seconds
  setInterval(checkForTrigger, 2000);
});

async function checkForTrigger() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const triggerCell = sheet.getRange("Z1");
      triggerCell.load("values");
      await context.sync();

      const trigger = triggerCell.values[0][0];
      if (trigger === "RunColorChange") {
        await runColorChange(context);
        triggerCell.values = [[""]];
        await context.sync();
      }
    });
  } catch (error) {
    console.error("Error checking for trigger:", error);
  }
}

async function runColorChange(context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load(["rowCount", "columnCount"]);
  await context.sync();

  const isSingleCell = selectedRange.rowCount === 1 && selectedRange.columnCount === 1;
  selectedRange.format.fill.color = isSingleCell ? "green" : "yellow";
}
