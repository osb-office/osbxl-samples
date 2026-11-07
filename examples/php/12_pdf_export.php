<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
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