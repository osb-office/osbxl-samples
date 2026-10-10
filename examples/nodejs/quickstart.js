const fs = require("fs");
const osbxl = require("nodeosbxl");

function main() {
    const filePath = "monthly_report.xlsx";
    if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
    }

    const app = new osbxl.App();
    const wb = app.createWorkBook(filePath);

    try {
        const ws = wb.addWorkSheet("Summary");

        // Headers
        ws.getRange("A1").setValue("Product");
        ws.getRange("B1").setValue("Q1");
        ws.getRange("C1").setValue("Q2");
        ws.getRange("D1").setValue("Total");

        // Data
        ws.getRange("A2").setValue("Server License");
        ws.getRange("B2").setNumberValue(250000);
        ws.getRange("C2").setNumberValue(320000);
        ws.getRange("D2").setFormula("SUM(B2:C2)");

        ws.getRange("A3").setValue("Support Plan");
        ws.getRange("B3").setNumberValue(50000);
        ws.getRange("C3").setNumberValue(60000);
        ws.getRange("D3").setFormula("SUM(B3:C3)");

        // Table
        const table = ws.getListObjects().addList("ReportTable", "A1:D3", true, true);
        table.setTotalRowLabel(3, "Total");
        table.setCustomRowFunction(4, "SUM(D2:D3)", false);

        wb.save();
        console.log(`Created: ${filePath}`);
    } finally {
        wb.close();
    }
}

main();

