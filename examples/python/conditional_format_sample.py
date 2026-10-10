"""
OSBXL Python Example: Conditional Formatting & Color Scales
Demonstrates:
- Cell value rules (e.g., Value > 60 with custom font & fill)
- 3-color scales (Minimum, Percentile 50, Maximum)
- Data bars
- Top/Bottom rules
"""

import os
import pyosbxl
from pyosbxl import enums, dto

def hex_color(hex_str: str) -> dto.ColorObject:
    c = dto.ColorObject()
    c.setHexColor(hex_str)
    return c

def main():
    file_path = "conditional_format_sample.xlsx"
    if os.path.exists(file_path):
        os.remove(file_path)

    app = pyosbxl.App()
    wb = app.createWorkBook(file_path)

    try:
        ws = wb.openWorkSheet("Sheet1")

        # 1. Header and Data Setup
        headers = ["Month", "Cell Value Rule", "Color Scale", "Data Bar"]
        for i, h in enumerate(headers):
            col = "ABCD"[i]
            ws.getRange(f"{col}1:{col}1").setValue(h)

        values = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49]
        for i, val in enumerate(values, start=2):
            ws.getRange(f"A{i}:A{i}").setNumberValue(i - 1)
            ws.getRange(f"B{i}:B{i}").setNumberValue(val)
            ws.getRange(f"C{i}:C{i}").setNumberValue(val)
            ws.getRange(f"D{i}:D{i}").setNumberValue(val)

        # 2. Conditional Formatting Setup
        fc = ws.getFormatConditions()

        # Rule 1: Cell Value > 60 (Highlight with light red fill and dark red bold text)
        cell_val = dto.CellValueObject()
        cell_val.setValueCondition(60, enums.XlFormatConditionOperator.FormatConditionGreater)
        cell_val.setA1C1("B2:B11")

        red_font = dto.FontObject()
        red_font.setBold(True)
        red_font.setColorObject(hex_color("9C0006"))
        cell_val.setFont(red_font)

        pink_fill = dto.PatternFillObject()
        pink_fill.setPattern(enums.XlPatternType.PatternSolid)
        pink_fill.setCellColor(hex_color("FFC7CE"))
        cell_val.setPatternFill(pink_fill)

        fc.setCellValueCondition(1, cell_val)

        # Rule 2: 3-Color Scale on C2:C11 (Red -> Yellow -> Green)
        color_scale = dto.ColorScaleObject()
        color_scale.setA1C1("C2:C11")
        color_scale.setMinimumCondition(hex_color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum)
        color_scale.setMediumCondition(hex_color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50)
        color_scale.setMaximumCondition(hex_color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum)
        fc.setColorScaleFormatCondition(2, color_scale)

        # Rule 3: Data Bar on D2:D11
        data_bar = dto.DataBarObject()
        data_bar.setA1C1("D2:D11")
        data_bar.setDataBarColor(hex_color("638EC6"))
        data_bar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum)
        data_bar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum)
        fc.setDataBarFormatCondition(3, data_bar)

        # Apply conditional formatting rules
        fc.apply()

        # Adjust column widths (1-based index)
        ws.getCol(1).setColumnWidth(10)
        ws.getCol(2).setColumnWidth(18)
        ws.getCol(3).setColumnWidth(16)
        ws.getCol(4).setColumnWidth(16)

        wb.save()
        print(f"Successfully generated: {file_path}")
    finally:
        wb.close()

if __name__ == "__main__":
    main()
