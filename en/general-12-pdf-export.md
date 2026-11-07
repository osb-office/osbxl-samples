# PDF Export — Page Setup, Margins, Print Area & Paper Size

level: advanced ／ verified: 2026-10-31 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-12-pdf-export.md)


## What you want to do


**Directly export high-quality PDF** from an Excel workbook or worksheet.  

osbxl can convert and output directly to PDF files or memory buffers at high speed without relying on external tools (LibreOffice, headless Chrome, etc.) or virtual printers.


By combining `PageSetupObject`, you can finely control **paper size (A4/A3/Letter, etc.), print orientation (portrait/landscape), margins (top/bottom/left/right), scaling (zoom), and horizontal centering**.


| Output Method | Method | Use Case |
| --- | --- | --- |
| Save directly to file | `wb.saveAsPdf(filePath, includeLayout, sheetNames?, includeHeaderFooter?, options?)` | Saving reports and invoice documents |
| Output to memory buffer | `wb.saveAsPdfToBuffer(includeLayout, sheetNames?, includeHeaderFooter?, options?)` | Web API response or streaming email attachments |
| Save single worksheet | `ws.saveAsPdf(filePath, includeLayout, includeHeaderFooter?, options?)` | Targeted output of a specific worksheet |


## Main Arguments and Options


| Argument | Type | Description |
| --- | --- | --- |
| `savePath` | string | Target PDF output file path. |
| `includeLayout` | boolean | When `true`, outputs row/column headings and gridlines. |
| `sheetNames` | string[] | Array of sheet names to export. An empty array `[]` exports all visible sheets. |
| `includeHeaderFooter` | boolean | Whether to include headers and footers (default: `false`). |
| `options` | `dto.PdfExportObject` | Advanced PDF export options including password encryption (open/permissions), printing/copying permissions, and page range overrides (`fromPage` / `toPage`). |


## Page Setup with PageSetupObject


Retrieve `ws.getPageSetupObject()` from the worksheet, configure desired settings, and apply back with `ws.setPageSetupObject(ps)`.


| Method | Value / Type | Description |
| --- | --- | --- |
| `setPaperSize(size)` | `enums.XlPaperSize` | Specifies paper size (`PaperA4`, `PaperA3`, `PaperLetter`, etc.). |
| `setOrientation(orient)` | `enums.XlPageOrientation` | Page orientation (`Portrait` / `Landscape`). |
| `setLeftMargin(val)` / `setRightMargin(val)` | number (inches) | Left and right margins (Excel default: ~0.7 in, narrow: ~0.25 in). |
| `setTopMargin(val)` / `setBottomMargin(val)` | number (inches) | Top and bottom margins (Excel default: ~0.75 in, narrow: ~0.3 in). |
| `setCenterHorizontally(true)` | boolean | Centers the printed content horizontally on the page. |
| `setZoom(percent)` | number (%) | Sets zoom percentage for printing (e.g. `140` for 140%). |
| `setPrintArea(a1c1)` | string | Defines print area cell range (e.g. `"A1:G15"`). |


## Code


Creates a sales report workbook, applies A4 landscape page setup with custom margins, and exports both to a PDF file and an in-memory buffer.


#### [Node.js]

```javascript
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

// 5. PDF 出力オプション（パスワード保護・権限設定）
const { dto } = require("nodeosbxl");
const pdfOpt = new dto.PdfExportObject();
pdfOpt.setOpenPassword("user2026");       // 閲覧パスワード
pdfOpt.setPermissionsPassword("adm2026"); // 権限パスワード
pdfOpt.setAllowPrint(true);               // 印刷許可
pdfOpt.setAllowCopy(false);               // コピー禁止

// 6. メモリバッファへの出力（Web API レスポンス等に最適）
const pdfBuffer = wb.saveAsPdfToBuffer(false, ["Sheet1"], true, pdfOpt);
console.log("PDF バッファ生成完了: バイト長 =", pdfBuffer.length);

wb.close();
```


#### [Python]

```python
# pdf_export.py — PDF エクスポートとページ設定 (pyosbxl)
import pyosbxl
from pyosbxl import enums

app = pyosbxl.App()
wb = app.createWorkBook("report.xlsx")
ws = wb.openWorkSheet(0)  # 0 始まり (zeroBase)

# 1. レポート表の作成
ws.getRange("A1").setValue("2026年度 部門別売上実績レポート")
headers = ["部門名", "第1四半期", "第2四半期", "第3四半期", "第4四半期", "合計"]
for c, h in enumerate(headers):
    ws.getCells(2, c).setValue(h)

rows = [
    ["営業1課", 1200000, 1450000, 1380000, 1820000],
    ["営業2課", 980000,  1120000, 1050000, 1280000],
    ["開発営業", 2150000, 2300000, 2420000, 2890000]
]
for r, row in enumerate(rows):
    ws.getCells(3 + r, 0).setValue(row[0])
    for c in range(1, 5):
        ws.getCells(3 + r, c).setNumberValue(row[c])
    ws.getCells(3 + r, 5).setFormula(f"SUM(B{4 + r}:E{4 + r})")

# 2. ページ設定（A4・横向き・余白調整・水平中央配置）
ps = ws.getPageSetupObject()
ps.setPaperSize(enums.XlPaperSize.PaperA4)
ps.setOrientation(enums.XlPageOrientation.Landscape)
ps.setLeftMargin(0.5)
ps.setRightMargin(0.5)
ps.setTopMargin(0.6)
ps.setBottomMargin(0.6)
ps.setCenterHorizontally(True)
ws.setPageSetupObject(ps)

# 3. ブックの保存
wb.save()

# 4. PDF 出力（ファイル出力）
pdf_result = wb.saveAsPdf("report.pdf", False)  # includeLayout = False
print(f"PDF 出力完了: 総ページ数 = {pdf_result.getPageCount()}")

# 5. メモリバッファへの出力
pdf_buffer = wb.saveAsPdfToBuffer(False, ["Sheet1"])
print(f"PDF バッファ生成完了: バイト長 = {len(pdf_buffer)}")

wb.close()
```


#### [Java]

```java
// PdfExportExample.java — PDF エクスポートとページ設定
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.enums.*;
import com.osboffice.osbxl.dto.*;
import java.util.Collections;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("report-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

        // 1. レポート表の作成
        ws.getRange("A1").setValue("2026年度 部門別売上実績レポート");
        String[] headers = {"部門名", "第1四半期", "第2四半期", "第3四半期", "第4四半期", "合計"};
        for (int c = 0; c < headers.length; c++) {
            ws.getCells(2, c).setValue(headers[c]);
        }

        String[] depts = {"営業1課", "営業2課", "開発営業"};
        double[][] sales = {
            {1200000, 1450000, 1380000, 1820000},
            {980000,  1120000, 1050000, 1280000},
            {2150000, 2300000, 2420000, 2890000}
        };

        for (int r = 0; r < depts.length; r++) {
            ws.getCells(3 + r, 0).setValue(depts[r]);
            for (int c = 0; c < 4; c++) {
                ws.getCells(3 + r, c + 1).setNumberValue(sales[r][c]);
            }
            ws.getCells(3 + r, 5).setFormula("SUM(B" + (4 + r) + ":E" + (4 + r) + ")");
        }

        // 2. ページ設定
        PageSetupObjectWrapper ps = ws.getPageSetupObject();
        ps.setPaperSize(XlPaperSize.PaperA4);
        ps.setOrientation(XlPageOrientation.Landscape);
        ps.setLeftMargin(0.5);
        ps.setRightMargin(0.5);
        ps.setTopMargin(0.6);
        ps.setBottomMargin(0.6);
        ps.setCenterHorizontally(true);
        ws.setPageSetupObject(ps);

        // 3. ブックの保存
        wb.save();

        // 4. PDF 出力
        PdfExportResultWrapper result = wb.saveAsPdf("report-java.pdf", false);
        System.out.println("PDF 出力完了: 総ページ数 = " + result.getPageCount());

        // 5. メモリバッファへの出力
        byte[] buffer = wb.saveAsPdfToBuffer(false, Collections.singletonList("Sheet1"));
        System.out.println("PDF バッファ生成完了: バイト長 = " + buffer.length);

        wb.close();
    }
}
```


#### [PHP]

```php
<?php
// pdf_export.php — PDF エクスポートとページ設定
use phposbxl\App;
use phposbxl\enums\XlPaperSize;
use phposbxl\enums\XlPageOrientation;

$app = new App();
$wb = $app->createWorkBook("report.xlsx");
$ws = $wb->openWorkSheet(0); // 0 始まり (zeroBase)

// 1. レポート表の作成
$ws->getRange("A1")->setValue("2026年度 部門別売上実績レポート");
$headers = ["部門名", "第1四半期", "第2四半期", "第3四半期", "第4四半期", "合計"];
foreach ($headers as $c => $h) {
    $ws->getCells(2, $c)->setValue($h);
}

$rows = [
    ["営業1課", 1200000, 1450000, 1380000, 1820000],
    ["営業2課", 980000,  1120000, 1050000, 1280000],
    ["開発営業", 2150000, 2300000, 2420000, 2890000]
];
foreach ($rows as $r => $row) {
    $ws->getCells(3 + $r, 0)->setValue($row[0]);
    for ($c = 1; $c <= 4; $c++) {
        $ws->getCells(3 + $r, $c)->setNumberValue($row[$c]);
    }
    $rowNum = 4 + $r;
    $ws->getCells(3 + $r, 5)->setFormula("SUM(B{$rowNum}:E{$rowNum})");
}

// 2. ページ設定
$ps = $ws->getPageSetupObject();
$ps->setPaperSize(XlPaperSize::PaperA4);
$ps->setOrientation(XlPageOrientation::Landscape);
$ps->setLeftMargin(0.5);
$ps->setRightMargin(0.5);
$ps->setTopMargin(0.6);
$ps->setBottomMargin(0.6);
$ps->setCenterHorizontally(true);
$ws->setPageSetupObject($ps);

// 3. ブックの保存
$wb->save();

// 4. PDF 出力
$pdfResult = $wb->saveAsPdf("report.pdf", false);
echo "PDF 出力完了: 総ページ数 = " . $pdfResult->getPageCount() . "\n";

// 5. メモリバッファへの出力
$pdfBuffer = $wb->saveAsPdfToBuffer(false, ["Sheet1"]);
echo "PDF バッファ生成完了: バイト長 = " . strlen($pdfBuffer) . "\n";

$wb->close();
```


#### [WebAssembly]

```javascript
// Note: PDF export (OSBXL_ENABLE_PDF) is currently not supported in the WebAssembly (wasmosbxl) build.
// PageSetupObject configurations (margins, paper size, orientation) and workbook saving are fully functional in WASM.
// For exporting directly to PDF, please use Node.js, Python, Java, or PHP.
```


## Real-World Example: In-Memory PDF Streaming with Table & Column Chart (Express)


Builds a structured table with total row formulas, adds a modern column chart, applies A4 landscape page setup with custom headers/footers, sets AES-256 PDF encryption, and streams the generated PDF binary directly from memory in Express without saving temporary files to disk.


`const express = require("express");
const { nodeosbxl, dto, enums } = require("nodeosbxl");

const app = express();

app.post("/api/export-pdf", (req, res) => {
    const xlApp = new nodeosbxl.App();
    const wb = xlApp.createWorkBook("sales_report.xlsx");
    const ws = wb.openWorkSheet(0); // 0-based indexing (zeroBase)

    // 1. Title and headers
    ws.getRange("A1").setValue("2026 Quarterly Sales Performance Report");
    const headers = ["Quarter", "Product A ($)", "Product B ($)", "Total ($)"];
    headers.forEach((h, c) => ws.getCells(2, c).setValue(h));

    // 2. Data rows and formulas
    const sales = [
        ["Q1", 120,  95],
        ["Q2", 145, 110],
        ["Q3", 160, 125],
        ["Q4", 190, 140],
    ];
    sales.forEach((row, idx) => {
        const r = 3 + idx;
        ws.getCells(r, 0).setValue(row[0]);
        ws.getCells(r, 1).setNumberValue(row[1]);
        ws.getCells(r, 2).setNumberValue(row[2]);
        const excelRow = r + 1; // 1-based row number
        ws.getCells(r, 3).setFormula(`SUM(B${excelRow}:C${excelRow})`);
    });

    // 3. Structured table (ListObjects) with totals row (SUM)
    const lo = ws.getListObjects();
    const table = lo.addList("SalesTable", "A3:D7", true, true);
    table.setTotalRowLabel(0, "Total");
    table.setTotalRowFunction(1, enums.XlTotalsCalculation.TotalsCalculationSum);
    table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
    table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);

    // 4. Clustered column chart positioned to the right (left: 380pt, top: 40pt)
    const co = ws.addChartObject("SalesChart", 380, 40, 420, 240);
    const chart = co.getChart();
    chart.chartWizard(
        "Sheet1!A3:C7",
        enums.XlChartType.ChartTypeColumnClustered,
        enums.XlRowCol.Columns,
        1, 1, 1,
        enums.XlChartColorPalette.XlPaletteColorful1,
        true,
        "Quarterly Sales Trends"
    );

    // 5. PageSetupObject: A4 landscape, centered, fit to 1x1 page
    const ps = ws.getPageSetupObject();
    ps.setPaperSize(enums.XlPaperSize.PaperA4);
    ps.setOrientation(enums.XlPageOrientation.Landscape);
    ps.setCenterHorizontally(true);
    ps.setCenterVertically(true);
    ps.setFitToPages(1, 1);

    const header = new dto.HeaderObject();
    header.setLeftHeader("Sample Corporation");
    header.setRightHeader("Date: 2026/11/07");
    ps.setHeader(header);

    const footer = new dto.FooterObject();
    footer.setCenterFooter("Page 1 / 1");
    ps.setFooter(footer);
    ws.setPageSetupObject(ps);

    // 6. PdfExportObject: password encryption and security options
    const pdfOpt = new dto.PdfExportObject();
    pdfOpt.setOpenPassword("customer2026");       // Open password (AES-256)
    pdfOpt.setPermissionsPassword("adminSecret"); // Master permissions password
    pdfOpt.setAllowPrint(true);
    pdfOpt.setAllowCopy(false);
    pdfOpt.setColorMode(enums.XlPdfColorMode.Color);
    pdfOpt.setFitMode(enums.XlPdfFitMode.FitAllColumnsOnOnePage);

    // 7. Generate in-memory PDF Buffer directly (zero disk writes)
    const pdfBuffer = wb.saveAsPdfToBuffer(false, ["Sheet1"], true, pdfOpt);
    wb.close();

    // 8. Stream directly to client
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'attachment; filename="sales_report.pdf"');
    res.setHeader("Content-Length", pdfBuffer.length);
    res.send(pdfBuffer);
});

app.listen(3000);`


### Output PDF Preview

Actual rendered preview of the PDF received directly via `curl -X POST http://localhost:3000/api/export-pdf -o sales_report.pdf`:

![PDF output preview with data table and clustered column chart](../assets/images/sales_report_pdf_preview.png)


## 解説

* **外部依存のない高速 PDF レンダリング:** osbxl 内蔵のレンダリングエンジンによって、Excel の書式・罫線・フォント・数値表示形式を忠実に再現した PDF を超高速に出力します。
* **`includeLayout` の使い分け:** 帳票や納品書など通常の印刷用途では `false` を指定します。`true` にすると、Excel の行番号・列番号（1, 2, 3... や A, B, C...）とグリッド枠線が付与されるため、レビュー資料や校正用途に最適です。
* **メモリバッファ出力（`saveAsPdfToBuffer`）:** 一時ファイルをディスクに作成することなくバイナリを取得できるため、サーバーレス環境（AWS Lambda や Cloud Functions）や Web API から直接ダウンロード配信する場合に最適です。
* **余白と拡大率の調整:** 標準では Excel の標準余白が適用されます。余白を `0.25`〜`0.3` インチに設定し、`ps.setCenterHorizontally(true)` を指定することで、用紙いっぱいに美しくレイアウトされた印刷面が得られます。


## 注意点

* **PDF 出力前のブック保存:** `saveAsPdf()` を呼び出す前に、ワークシートへの変更内容を `wb.save()` して最新状態を確定させておくことが推奨されます。
* **出力対象シートの指定:** `sheetNames` を省略または空配列 `[]` を渡すと、ブック内のすべての「表示（visible）」シートが出力されます。非表示（hidden）シートは出力対象外となります。


## See also

* [Index](index.md)
* Previous: [Pivot tables](general-11-pivottables.md)
