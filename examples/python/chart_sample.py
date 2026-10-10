"""
OSBXL Python Example: Chart Generation (Column & Line Charts)
Demonstrates:
- Creating data tables with headers
- Adding clustered column chart (addChartObject + chartWizard)
- Adding line chart with markers
- Chart positioning, titles, and legends
"""

import os
import pyosbxl
from pyosbxl import enums

def main():
    file_path = "chart_sample.xlsx"
    if os.path.exists(file_path):
        os.remove(file_path)

    app = pyosbxl.App()
    wb = app.createWorkBook(file_path)

    try:
        ws = wb.openWorkSheet("Sheet1")

        # 1. Populate data table (Monthly sales of Product A & B)
        headers = ["Month", "Product A", "Product B"]
        for c, h in enumerate(headers, start=1):
            ws.getCells(1, c).setValue(h)

        sales_data = [
            ("Jan", 120, 90),
            ("Feb", 135, 95),
            ("Mar", 150, 110),
            ("Apr", 142, 125),
            ("May", 160, 118),
            ("Jun", 175, 130),
        ]

        for r, (month, p_a, p_b) in enumerate(sales_data, start=2):
            ws.getCells(r, 1).setValue(month)
            ws.getCells(r, 2).setNumberValue(p_a)
            ws.getCells(r, 3).setNumberValue(p_b)

        # 2. Add Clustered Column Chart
        # addChartObject(name, left, top, width, height) in points (pt)
        co1 = ws.addChartObject("SalesColumn", 260, 15, 380, 220)
        chart1 = co1.getChart()
        chart1.chartWizard(
            "Sheet1!A1:C7",
            enums.XlChartType.ChartTypeColumnClustered,
            enums.XlRowCol.Columns,   # Plot by columns (Product A / B as series)
            1,                        # seriesLabelLines (row 1 is series header)
            1,                        # categoryLabelLines (col 1 is category header)
            1,                        # chartStyleNo
            enums.XlChartColorPalette.XlPaletteColorful1,
            True,                     # hasLegend
            "Monthly Sales Comparison" # Chart Title
        )

        # 3. Add Line Chart with Markers
        co2 = ws.addChartObject("SalesLine", 260, 250, 380, 220)
        chart2 = co2.getChart()
        chart2.chartWizard(
            "Sheet1!A1:C7",
            enums.XlChartType.ChartTypeLineMarkers,
            enums.XlRowCol.Columns,
            1, 1, 1,
            enums.XlChartColorPalette.XlPaletteColorful1,
            True,
            "Sales Trend"
        )

        wb.save()
        print(f"Successfully generated: {file_path}")
    finally:
        wb.close()

if __name__ == "__main__":
    main()

