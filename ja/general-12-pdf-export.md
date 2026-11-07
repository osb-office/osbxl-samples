# PDF エクスポート — ページ設定・余白・印刷範囲・用紙サイズ

level: advanced ／ verified: 2026-10-31 ／ 5言語タブ対応 ／ [English](../en/general-12-pdf-export.md)


## やりたいこと


Excel ワークブックやワークシートから**高品質な PDF を直接エクスポート**する。  

外部ツール（LibreOffice や headless Chrome など）や仮想プリンターに依存せず、osbxl 単体で高速に PDF ファイルやメモリバッファへ変換出力できます。


`PageSetupObject` を組み合わせることで、**用紙サイズ（A4/A3/Letter 等）・印刷の向き（縦/横）・余白（上下左右）・拡大縮小（ズーム）・中央配置**をきめ細かく制御できます。


| 出力方式 | メソッド | 用途 |
| --- | --- | --- |
| ファイルへ直接保存 | `wb.saveAsPdf(filePath, includeLayout, sheetNames?, includeHeaderFooter?, options?)` | 帳票ファイルやレポートの保存 |
| メモリバッファへ出力 | `wb.saveAsPdfToBuffer(includeLayout, sheetNames?, includeHeaderFooter?, options?)` | Web API のレスポンスやメール添付用ストリーム |
| 単一シートの保存 | `ws.saveAsPdf(filePath, includeLayout, includeHeaderFooter?, options?)` | 特定シートのみのピンポイント出力 |


## 主な引数とオプション


| 引数 | 型 | 説明 |
| --- | --- | --- |
| `savePath` | string | 保存先の PDF ファイルパス。 |
| `includeLayout` | boolean | `true` の場合、シートの行番号・列番号の見出しとセル枠線を含めて出力します。 |
| `sheetNames` | string[] | 出力対象のシート名の配列。空配列 `[]` の場合はすべての可視シートを出力します。 |
| `includeHeaderFooter` | boolean | ヘッダー・フッターを含めるかどうか（既定: `false`）。 |
| `options` | `dto.PdfExportObject` | エクスポート詳細設定（フォント置換や出力範囲の指定等）。 |


## PageSetupObject による印刷設定


ワークシートの `ws.getPageSetupObject()` を取得して値を設定し、`ws.setPageSetupObject(ps)` で反映します。


| 設定メソッド | 設定値・型 | 説明 |
| --- | --- | --- |
| `setPaperSize(size)` | `enums.XlPaperSize` | 用紙サイズを指定（`PaperA4`, `PaperA3`, `PaperLetter` など）。 |
| `setOrientation(orient)` | `enums.XlPageOrientation` | 印刷の向き（`Portrait` = 縦向き / `Landscape` = 横向き）。 |
| `setLeftMargin(val)` / `setRightMargin(val)` | 数値（インチ単位） | 左右の余白（Excel 標準は約 0.7 インチ、狭小は約 0.25 インチ）。 |
| `setTopMargin(val)` / `setBottomMargin(val)` | 数値（インチ単位） | 上下の余白（標準は約 0.75 インチ、狭小は約 0.3 インチ）。 |
| `setCenterHorizontally(true)` | boolean | 用紙の水平中央に印刷内容をセンタリングします。 |
| `setZoom(percent)` | 数値（パーセント） | 印刷時の拡大率を指定（例: `140` で 140% 拡大）。 |
| `setPrintArea(a1c1)` | string | 印刷対象セル範囲を指定（例: `"A1:G15"`）。 |


## コード


売上レポートのワークブックを作成し、A4 横向き・余白調整のうえ PDF ファイルおよびメモリバッファへ出力します。


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
console.log("PDF 出力完了: 総ページ数 =", pdfResult.getTotalPages());

// 5. メモリバッファへの出力（Web API レスポンス等に最適）
const pdfBuffer = wb.saveAsPdfToBuffer(false, ["Sheet1"]);
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
print(f"PDF 出力完了: 総ページ数 = {pdf_result.getTotalPages()}")

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

public class PdfExportExample {
    public static void main(String[] args) {
        App app = new App();
        WorkBook wb = app.createWorkBook("report.xlsx");
        WorkSheet ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

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
        PageSetupObject ps = ws.getPageSetupObject();
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
        PdfExportResult result = wb.saveAsPdf("report.pdf", false);
        System.out.println("PDF 出力完了: 総ページ数 = " + result.getTotalPages());

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
echo "PDF 出力完了: 総ページ数 = " . $pdfResult->getTotalPages() . "\n";

// 5. メモリバッファへの出力
$pdfBuffer = $wb->saveAsPdfToBuffer(false, ["Sheet1"]);
echo "PDF バッファ生成完了: バイト長 = " . strlen($pdfBuffer) . "\n";

$wb->close();
```


#### [WebAssembly]

```javascript
// pdf-export-wasm.js — WebAssembly 版での PDF エクスポート
const { osbxl, enums } = require("wasmosbxl");

const app = new osbxl.App();
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
    ws.getCells(3 + r, 5).setFormula(`SUM(B${4 + r}:E${4 + r})`);
});

// 2. ページ設定
const ps = ws.getPageSetupObject();
ps.setPaperSize(enums.XlPaperSize.PaperA4);
ps.setOrientation(enums.XlPageOrientation.Landscape);
ps.setLeftMargin(0.5);
ps.setRightMargin(0.5);
ps.setTopMargin(0.6);
ps.setBottomMargin(0.6);
ps.setCenterHorizontally(true);
ws.setPageSetupObject(ps);

// 3. ブックの保存
wb.save();

// 4. PDF 出力
const pdfResult = wb.saveAsPdf("report.pdf", false);
console.log("PDF 出力完了: 総ページ数 =", pdfResult.getTotalPages());

// 5. メモリバッファへの出力 (Uint8Array)
const pdfBuffer = wb.saveAsPdfToBuffer(false, ["Sheet1"]);
console.log("PDF バッファ生成完了: バイト長 =", pdfBuffer.byteLength);

wb.close();
```


## 実践例: データテーブルとグラフを並べた PDF 直接ストリーミング (Express)


データテーブル（構造化テーブル + 合計行）と集合縦棒グラフを作成し、A4 横向き・ヘッダー・フッター・パスワード暗号化（AES-256）を適用して、Express サーバーからディスク一時ファイル保存ゼロでクライアントへ直接ストリーミング返却する実用コードです。


`const express = require("express");
const { nodeosbxl, dto, enums } = require("nodeosbxl");

const app = express();

app.post("/api/export-pdf", (req, res) => {
    const xlApp = new nodeosbxl.App();
    const wb = xlApp.createWorkBook("sales_report.xlsx");
    const ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

    // 1. タイトルとテーブルヘッダー
    ws.getRange("A1").setValue("2026年度 四半期別 売上実績レポート");
    const headers = ["四半期", "商品A (万円)", "商品B (万円)", "合計 (万円)"];
    headers.forEach((h, c) => ws.getCells(2, c).setValue(h));

    // 2. 明細データと合計計算式
    const sales = [
        ["第1四半期", 120,  95],
        ["第2四半期", 145, 110],
        ["第3四半期", 160, 125],
        ["第4四半期", 190, 140],
    ];
    sales.forEach((row, idx) => {
        const r = 3 + idx;
        ws.getCells(r, 0).setValue(row[0]);
        ws.getCells(r, 1).setNumberValue(row[1]);
        ws.getCells(r, 2).setNumberValue(row[2]);
        const excelRow = r + 1; // 1-based 表示行
        ws.getCells(r, 3).setFormula(`SUM(B${excelRow}:C${excelRow})`);
    });

    // 3. 構造化テーブル (ListObjects) と集計行 (SUM) の作成
    const lo = ws.getListObjects();
    const table = lo.addList("SalesTable", "A3:D7", true, true);
    table.setTotalRowLabel(0, "通期合計");
    table.setTotalRowFunction(1, enums.XlTotalsCalculation.TotalsCalculationSum);
    table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
    table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);

    // 4. 集合縦棒グラフの配置 (left: 380pt, top: 40pt)
    const co = ws.addChartObject("SalesChart", 380, 40, 420, 240);
    const chart = co.getChart();
    chart.chartWizard(
        "Sheet1!A3:C7",
        enums.XlChartType.ChartTypeColumnClustered,
        enums.XlRowCol.Columns,
        1, 1, 1,
        enums.XlChartColorPalette.XlPaletteColorful1,
        true,
        "四半期売上推移 (商品別)"
    );

    // 5. PageSetupObject による印刷・PDF ページレイアウト設定
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

    // 6. PdfExportObject によるセキュリティ・出力オプション
    const pdfOpt = new dto.PdfExportObject();
    pdfOpt.setOpenPassword("customer2026");       // 閲覧パスワード (AES-256)
    pdfOpt.setPermissionsPassword("adminSecret"); // 権限マスターパスワード
    pdfOpt.setAllowPrint(true);
    pdfOpt.setAllowCopy(false);
    pdfOpt.setColorMode(enums.XlPdfColorMode.Color);
    pdfOpt.setFitMode(enums.XlPdfFitMode.FitAllColumnsOnOnePage);

    // 7. メモリ直接 PDF 生成 (ディスク一時保存なし)
    const pdfBuffer = wb.saveAsPdfToBuffer(false, ["Sheet1"], true, pdfOpt);
    wb.close();

    // 8. クライアントへ直接ストリーミング返却
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'attachment; filename="sales_report.pdf"');
    res.setHeader("Content-Length", pdfBuffer.length);
    res.send(pdfBuffer);
});

app.listen(3000);`


### 出力 PDF プレビュー

上記 Express サーバーへ `curl -X POST http://localhost:3000/api/export-pdf -o sales_report.pdf` を実行し、取得した PDF を描画したプレビューです：

![売上実績レポート PDF 出力プレビュー](../assets/images/sales_report_pdf_preview.png)


## 解説

* **外部依存のない高速 PDF レンダリング:** osbxl 内蔵のレンダリングエンジンによって、Excel の書式・罫線・フォント・数値表示形式を忠実に再現した PDF を超高速に出力します。
* **`includeLayout` の使い分け:** 帳票や納品書など通常の印刷用途では `false` を指定します。`true` にすると、Excel の行番号・列番号（1, 2, 3... や A, B, C...）とグリッド枠線が付与されるため、レビュー資料や校正用途に最適です。
* **メモリバッファ出力（`saveAsPdfToBuffer`）:** 一時ファイルをディスクに作成することなくバイナリを取得できるため、サーバーレス環境（AWS Lambda や Cloud Functions）や Web API から直接ダウンロード配信する場合に最適です。
* **余白と拡大率の調整:** 標準では Excel の標準余白が適用されます。余白を `0.25`〜`0.3` インチに設定し、`ps.setCenterHorizontally(true)` を指定することで、用紙いっぱいに美しくレイアウトされた印刷面が得られます。


## 注意点

* **PDF 出力前のブック保存:** `saveAsPdf()` を呼び出す前に、ワークシートへの変更内容を `wb.save()` して最新状態を確定させておくことが推奨されます。
* **出力対象シートの指定:** `sheetNames` を省略または空配列 `[]` を渡すと、ブック内のすべての「表示（visible）」シートが出力されます。非表示（hidden）シートは出力対象外となります。


## 関連

* [目次](index.md)
* 前: [ピボットテーブル](general-11-pivottables.md)
