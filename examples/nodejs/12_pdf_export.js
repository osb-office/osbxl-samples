// pdf-export.js — PDF エクスポートとページ設定
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("report.xlsx");
const ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

// 1. レポート表の作成
ws.getRange("A1").setValue("2026年度 部門別売上実績レポート");
const headers = ["部門名", "第1四半期", "第2四半期", "第3四半期", "第4四半期", "合計"];
headers.forEach((h, c) => ws.getCells(2, c).setValue(h));

const rows = [
    ["営業1課", 1200000, 1450000, 1380000, 1820000],
    ["営業2課", 980000,  1120000, 1050000, 1280000],
    ["開発営業", 2150000, 2300000, 2420000, 2890000]
];
rows.forEach((row, r) => {
    ws.getCells(3 + r, 0).setValue(row[0]);
    for (let c = 1; c <= 4; c++) {
        ws.getCells(3 + r, c).setNumberValue(row[c]);
    }
    // 合計列（SUM 数式）
    ws.getCells(3 + r, 5).setFormula(`SUM(B${4 + r}:E${4 + r})`);
});

// 2. ページ設定（A4・横向き・余白調整・水平中央配置）
const ps = ws.getPageSetupObject();
ps.setPaperSize(enums.XlPaperSize.PaperA4);
ps.setOrientation(enums.XlPageOrientation.Landscape);
ps.setLeftMargin(0.5);   // 左余白 0.5 インチ
ps.setRightMargin(0.5);  // 右余白 0.5 インチ
ps.setTopMargin(0.6);    // 上余白 0.6 インチ
ps.setBottomMargin(0.6); // 下余白 0.6 インチ
ps.setCenterHorizontally(true); // 水平中央に配置
ws.setPageSetupObject(ps);

// 3. ブックの保存
wb.save();

// 4. PDF 出力（ファイル出力）
const pdfResult = wb.saveAsPdf("report.pdf", false); // includeLayout = false
console.log("PDF 出力完了: 総ページ数 =", pdfResult.getPageCount());

// 5. メモリバッファへの出力（Web API レスポンス等に最適）
const pdfBuffer = wb.saveAsPdfToBuffer(false, ["Sheet1"]);
console.log("PDF バッファ生成完了: バイト長 =", pdfBuffer.length);

wb.close();