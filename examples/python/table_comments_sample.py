"""
OSBXL Python Example: Table (ListObject), Formulas & Modern Comments
Demonstrates:
- Creating an Excel Table (ListObject) from a range
- Adding total rows and structured references (e.g., SUM(Sales[Price]))
- Adding threaded modern comments (CommentObject) and memos
"""

import os
import pyosbxl
from pyosbxl import enums, dto

def main():
    file_path = "table_comments_sample.xlsx"
    if os.path.exists(file_path):
        os.remove(file_path)

    app = pyosbxl.App()
    wb = app.createWorkBook(file_path)

    try:
        ws = wb.openWorkSheet("Sheet1")

        # 1. Populate data rows using setRangeData for efficiency
        rows = [
            ["Product", "Unit Price", "Quantity"],
            ["Apple", "120", "15"],
            ["Orange", "80", "25"],
            ["Banana", "100", "10"],
            ["Grape", "300", "8"],
        ]
        ws.setRangeData("A1", rows)

        # 2. Create Table (ListObject)
        # addList(tableName, range, hasHeaders, showTotalRow)
        table = ws.getListObjects().addList("ProductSales", "A1:C5", True, True)

        # Configure Total Row functions (1-based column indices)
        table.setTotalRowLabel(1, "Total")
        table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationAverage)
        table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum)

        # 3. Add modern threaded comment
        comments = ws.getComments()
        c1 = dto.CommentObject()
        c1.setAuthor("Alice")
        c1.setContent("Please review unit prices before the promotion.")

        c2 = dto.CommentObject()
        c2.setAuthor("Bob")
        c2.setContent("Approved for next month.")

        comments.setCommentThread("B2", [c1, c2])

        # 4. Column formatting
        ws.getCol(1).setColumnWidth(16)
        ws.getCol(2).setColumnWidth(14)
        ws.getCol(3).setColumnWidth(14)

        wb.save()
        print(f"Successfully generated: {file_path}")
    finally:
        wb.close()

if __name__ == "__main__":
    main()

