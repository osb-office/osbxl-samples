import os
import pyosbxl
from pyosbxl import enums

def main():
    file_path = "sales_report.xlsx"
    if os.path.exists(file_path):
        os.remove(file_path)

    app = pyosbxl.App()

    # Create new workbook
    wb = app.createWorkBook(file_path)
    try:
        ws = wb.openWorkSheet("Sheet1")

        # Set headers
        ws.getRange("A1").setValue("Item")
        ws.getRange("B1").setValue("Unit Price")
        ws.getRange("C1").setValue("Quantity")
        ws.getRange("D1").setValue("Total")

        # Set data rows
        items = [
            ("Apple", 100, 5),
            ("Orange", 60, 10),
            ("Banana", 80, 8),
            ("Grape", 200, 3),
        ]

        for i, (name, price, qty) in enumerate(items, start=2):
            ws.getRange(f"A{i}").setValue(name)
            ws.getRange(f"B{i}").setNumberValue(price)
            ws.getRange(f"C{i}").setNumberValue(qty)
            ws.getRange(f"D{i}").setFormula(f"B{i}*C{i}")

        # Add Excel Table (ListObjects) with total row
        table = ws.getListObjects().addList("SalesTable", "A1:D5", True, True)
        table.setTotalRowLabel(3, "Grand Total")
        table.setCustomRowFunction(4, "SUM(D2:D5)", False)
        ws.getRange("A6").setValue("")

        # Formatting
        ws.getRange("B2:B6").setNumberFormat("#,##0")
        ws.getRange("D2:D6").setNumberFormat("#,##0")

        wb.save()
        print(f"Created: {file_path}")
    finally:
        wb.close()

if __name__ == "__main__":
    main()
