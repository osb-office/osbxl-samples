/**
 * OSBXL Node.js Example: Table (ListObject) & Threaded Comments
 * Demonstrates:
 * - Adding Excel Tables (ListObjects)
 * - Adding total row aggregations
 * - Adding modern threaded comments (CommentObject)
 */

const { nodeosbxl, enums, dto } = require("nodeosbxl");
const fs = require("fs");

function main() {
  const filePath = "table_comments_sample.xlsx";
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }

  const app = new nodeosbxl.App();
  const wb = app.createWorkBook(filePath);

  try {
    const ws = wb.openWorkSheet("Sheet1");

    // 1. Bulk populate using setValueArray
    const rows = [
      ["Product", "Unit Price", "Quantity"],
      ["Apple", 120, 15],
      ["Orange", 80, 25],
      ["Banana", 100, 10],
      ["Grape", 300, 8],
    ];

    const values = [];
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
        const obj = new dto.InputValueObject();
        if (typeof cell === "number") {
          obj.setNumberValue(a1, cell);
        } else {
          obj.setStringValue(a1, cell);
        }
        values.push(obj);
      });
    });
    ws.setValueArray(values);

    // 2. Create Table (ListObject)
    const table = ws.getListObjects().addList("ProductSales", "A1:C5", true, true);
    table.setTotalRowLabel(1, "Total");
    table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationAverage);
    table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);

    // 3. Add modern threaded comment
    const comments = ws.getComments();
    const c1 = new dto.CommentObject();
    c1.setAuthor("Alice");
    c1.setContent("Please review unit prices before the promotion.");

    const c2 = new dto.CommentObject();
    c2.setAuthor("Bob");
    c2.setContent("Approved for next month.");

    comments.setCommentThread("B2", [c1, c2]);

    // 4. Set column widths
    ws.getCol(1).setColumnWidth(16);
    ws.getCol(2).setColumnWidth(14);
    ws.getCol(3).setColumnWidth(14);

    wb.save();
    console.log(`Successfully generated: ${filePath}`);
  } finally {
    wb.close();
  }
}

main();
