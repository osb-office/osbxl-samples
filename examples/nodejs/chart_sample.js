/**
 * OSBXL Node.js Example: Chart Generation (Column & Line Charts)
 * Demonstrates:
 * - Creating data tables
 * - Adding clustered column chart (addChartObject + chartWizard)
 * - Adding line chart with markers
 * - Chart positioning, titles, and legends
 */

const { nodeosbxl, enums } = require("nodeosbxl");
const fs = require("fs");

function main() {
  const filePath = "chart_sample.xlsx";
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }

  const app = new nodeosbxl.App();
  const wb = app.createWorkBook(filePath);

  try {
    const ws = wb.openWorkSheet("Sheet1");

    // 1. Populate data table
    const headers = ["Month", "Product A", "Product B"];
    headers.forEach((h, c) => ws.getCells(1, c + 1).setValue(h));

    const salesData = [
      ["Jan", 120, 90],
      ["Feb", 135, 95],
      ["Mar", 150, 110],
      ["Apr", 142, 125],
      ["May", 160, 118],
      ["Jun", 175, 130],
    ];

    salesData.forEach((row, i) => {
      ws.getCells(i + 2, 1).setValue(row[0]);
      ws.getCells(i + 2, 2).setNumberValue(row[1]);
      ws.getCells(i + 2, 3).setNumberValue(row[2]);
    });

    // 2. Add Clustered Column Chart (addChartObject in points: left, top, width, height)
    const co1 = ws.addChartObject("SalesColumn", 260, 15, 380, 220);
    const chart1 = co1.getChart();
    chart1.chartWizard(
      "Sheet1!A1:C7",
      enums.XlChartType.ChartTypeColumnClustered,
      enums.XlRowCol.Columns,
      1,
      1,
      1,
      enums.XlChartColorPalette.XlPaletteColorful1,
      true,
      "Monthly Sales Comparison"
    );

    // 3. Add Line Chart with Markers
    const co2 = ws.addChartObject("SalesLine", 260, 250, 380, 220);
    const chart2 = co2.getChart();
    chart2.chartWizard(
      "Sheet1!A1:C7",
      enums.XlChartType.ChartTypeLineMarkers,
      enums.XlRowCol.Columns,
      1,
      1,
      1,
      enums.XlChartColorPalette.XlPaletteColorful1,
      true,
      "Sales Trend"
    );

    wb.save();
    console.log(`Successfully generated: ${filePath}`);
  } finally {
    wb.close();
  }
}

main();
