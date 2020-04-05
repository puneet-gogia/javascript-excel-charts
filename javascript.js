$("#setup").click(() => tryCatch(setup));
$("#create-line-chart").click(() => tryCatch(createLineChart));

async function createLineChart() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let salesTable = sheet.tables.getItem("SalesTable");

    let dataRange = sheet.getRange("A1:E7");
    let chart = sheet.charts.add("Line", dataRange, "Auto");

    chart.setPosition("A10", "F20");
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.title.text = "Bicycle Parts Quarterly Sales";

    await context.sync();
  });
}

async function setup() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    let expensesTable = sheet.tables.add("A1:B1", true);
    expensesTable.name = "SalesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Quantity"]];
    expensesTable.rows.add(null, [
      ["01-Jan-2020", 1897],
      ["02-Jan-2020", 1010],
      ["03-Jan-2020", 1359],
      ["04-Jan-2020", 1695],
      ["05-Jan-2020", 1374],
      ["06-Jan-2020", 1132],
      ["07-Jan-2020", 1302],
      ["08-Jan-2020", 1738],
      ["09-Jan-2020", 1835],
      ["10-Jan-2020", 1759]
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
    sheet.activate();
    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
