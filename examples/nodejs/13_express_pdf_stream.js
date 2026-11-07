const express = require("express");
const { nodeosbxl, dto, enums } = require("nodeosbxl");

const app = express();
const port = 3000;

app.post("/api/export-pdf", (req, res) => {
    try {
        const xlApp = new nodeosbxl.App();
        const wb = xlApp.createWorkBook("sales_report.xlsx");
        const ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

        // 1. 表タイトル
        ws.getRange("A1").setValue("2026年度 四半期別 売上実績レポート");

        // 2. テーブルデータ用ヘッダー
        const headers = ["四半期", "商品A (万円)", "商品B (万円)", "合計 (万円)"];
        headers.forEach((h, c) => ws.getCells(2, c).setValue(h));

        // 3. 明細データ
        const sales = [
            ["第1四半期", 120,  95],
            ["第2四半期", 145, 110],
            ["第3四半期", 160, 125],
            ["第4四半期", 190, 140],
        ];

        sales.forEach((row, idx) => {
            const r = 3 + idx;
            ws.getCells($r = r, 0).setValue(row[0]);
            ws.getCells(r, 1).setNumberValue(row[1]);
            ws.getCells(r, 2).setNumberValue(row[2]);

            const excelRow = r + 1;
            ws.getCells(r, 3).setFormula(`SUM(B${excelRow}:C${excelRow})`);
        });

        // 4. Excel 構造化テーブル作成
        const lo = ws.getListObjects();
        const table = lo.addList("SalesTable", "A3:D7", true, true);
        table.setTotalRowLabel(0, "通期合計");
        table.setTotalRowFunction(1, enums.XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);

        // 5. 集合縦棒グラフの配置 (left: 380pt, top: 40pt)
        const co = ws.addChartObject("SalesChart", 380, 40, 420, 240);
        const chart = co.getChart();
        chart.chartWizard(
            "Sheet1!A3:C7",
            enums.XlChartType.ChartTypeColumnClustered,
            enums.XlRowCol.Columns,
            1,
            1,
            1,
            enums.XlChartColorPalette.XlPaletteColorful1,
            true,
            "四半期売上推移 (商品別)"
        );

        // 6. PageSetupObject による印刷・PDF レイアウト設定
        const ps = ws.getPageSetupObject();
        ps.setPaperSize(enums.XlPaperSize.PaperA4);
        ps.setOrientation(enums.XlPageOrientation.Landscape);
        ps.setCenterHorizontally(true);
        ps.setCenterVertically(true);
        ps.setFitToPages(1, 1);

        const header = new dto.HeaderObject();
        header.setLeftHeader("株式会社サンプル");
        header.setRightHeader("出力日: 2026/11/07");
        ps.setHeader(header);

        const footer = new dto.FooterObject();
        footer.setCenterFooter("ページ 1 / 1");
        ps.setFooter(footer);

        ws.setPageSetupObject(ps);

        // 7. PdfExportObject によるオプション設定
        const pdfOpt = new dto.PdfExportObject();
        pdfOpt.setOpenPassword("customer2026");       // 開封パスワード
        pdfOpt.setPermissionsPassword("adminSecret"); // 権限パスワード
        pdfOpt.setAllowPrint(true);
        pdfOpt.setAllowCopy(false);
        pdfOpt.setAllowModify(false);
        pdfOpt.setColorMode(enums.XlPdfColorMode.Color);
        pdfOpt.setFitMode(enums.XlPdfFitMode.FitAllColumnsOnOnePage);

        // 8. メモリ直接 PDF 生成
        const pdfBuffer = wb.saveAsPdfToBuffer(
            false,
            ["Sheet1"],
            true,
            pdfOpt
        );

        wb.close();

        // 9. HTTP レスポンス返却
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", 'attachment; filename="sales_report.pdf"');
        res.setHeader("Content-Length", pdfBuffer.length);
        res.send(pdfBuffer);
    } catch (err) {
        console.error("PDF export error:", err);
        res.status(500).json({ error: err.message });
    }
});

const server = app.listen(port, () => {
    console.log(`Express PDF Server ready at http://localhost:${port}`);
});

process.on("SIGTERM", () => server.close());
process.on("SIGINT", () => server.close());

