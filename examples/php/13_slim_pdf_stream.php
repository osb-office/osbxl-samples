<?php
// 13_slim_pdf_stream.php — Slim 4 によるインメモリ PDF 直接ストリーミング (phposbxl)
use Psr\Http\Message\ResponseInterface as Response;
use Psr\Http\Message\ServerRequestInterface as Request;
use Slim\Factory\AppFactory;
use phposbxl\App;
use phposbxl\enums\XlChartType;
use phposbxl\enums\XlRowCol;
use phposbxl\enums\XlChartColorPalette;
use phposbxl\enums\XlTotalsCalculation;
use phposbxl\enums\XlPaperSize;
use phposbxl\enums\XlPageOrientation;
use phposbxl\dto\PdfExportObject;
use phposbxl\dto\HeaderObject;
use phposbxl\dto\FooterObject;
use phposbxl\enums\XlPdfColorMode;
use phposbxl\enums\XlPdfFitMode;

if (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}

$app = AppFactory::create();

$app->post('/api/export-pdf', function (Request $request, Response $response) {
    // 1. ワークブックとシートの準備 (zeroBase: 0-based)
    $xlApp = new App();
    $wb = $xlApp->createWorkBook("sales_report.xlsx");
    $ws = $wb->openWorkSheet(0);

    // 2. 表タイトルとヘッダー作成
    $ws->getRange("A1")->setValue("2026年度 四半期別 売上実績レポート");
    $headers = ["四半期", "商品A (万円)", "商品B (万円)", "合計 (万円)"];
    foreach ($headers as $c => $h) {
        $ws->getCells(2, $c)->setValue($h);
    }

    // 3. 明細データと合計計算式 (Excel 表示行は 1-based のため r + 1)
    $sales = [
        ["第1四半期", 120,  95],
        ["第2四半期", 145, 110],
        ["第3四半期", 160, 125],
        ["第4四半期", 190, 140],
    ];
    foreach ($sales as $idx => $row) {
        $r = 3 + $idx;
        $ws->getCells($r, 0)->setValue($row[0]);
        $ws->getCells($r, 1)->setNumberValue($row[1]);
        $ws->getCells($r, 2)->setNumberValue($row[2]);
        $rowNum = $r + 1;
        $ws->getCells($r, 3)->setFormula("SUM(B{$rowNum}:C{$rowNum})");
    }

    // 4. 構造化テーブル (ListObjects) と集計行 (SUM) の作成
    $lo = $ws->getListObjects();
    $table = $lo->addList("SalesTable", "A3:D7", true, true);
    $table->setTotalRowLabel(0, "通期合計");
    $table->setTotalRowFunction(1, XlTotalsCalculation::TotalsCalculationSum);
    $table->setTotalRowFunction(2, XlTotalsCalculation::TotalsCalculationSum);
    $table->setTotalRowFunction(3, XlTotalsCalculation::TotalsCalculationSum);

    // 5. 集合縦棒グラフの配置 (left: 380pt, top: 40pt)
    $co = $ws->addChartObject("SalesChart", 380, 40, 420, 240);
    $chart = $co->getChart();
    $chart->chartWizard(
        "Sheet1!A3:C7",
        XlChartType::ChartTypeColumnClustered,
        XlRowCol::Columns,
        1, 1, 1,
        XlChartColorPalette::XlPaletteColorful1,
        true,
        "四半期売上推移 (商品別)"
    );

    // 6. PageSetupObject による印刷・PDF ページレイアウト設定
    $ps = $ws->getPageSetupObject();
    $ps->setPaperSize(XlPaperSize::PaperA4);
    $ps->setOrientation(XlPageOrientation::Landscape);
    $ps->setCenterHorizontally(true);
    $ps->setCenterVertically(true);
    $ps->setFitToPages(1, 1);

    $header = new HeaderObject();
    $header->setLeftHeader("株式会社サンプル");
    $header->setRightHeader("出力日: 2026/11/07");
    $ps->setHeader($header);

    $footer = new FooterObject();
    $footer->setCenterFooter("ページ 1 / 1");
    $ps->setFooter($footer);
    $ws->setPageSetupObject($ps);

    // 7. PdfExportObject によるセキュリティ・出力オプション設定
    $pdfOpt = new PdfExportObject();
    $pdfOpt->setOpenPassword("customer2026");       // 開封パスワード (AES-256)
    $pdfOpt->setPermissionsPassword("adminSecret"); // 権限マスターパスワード
    $pdfOpt->setAllowPrint(true);
    $pdfOpt->setAllowCopy(false);
    $pdfOpt->setAllowModify(false);
    $pdfOpt->setColorMode(XlPdfColorMode::Color);
    $pdfOpt->setFitMode(XlPdfFitMode::FitAllColumnsOnOnePage);

    // 8. ディスク保存なしでメモリから直接 PDF バイナリを取得
    $pdfData = $wb->saveAsPdfToBuffer(false, ["Sheet1"], true, $pdfOpt);
    $wb->close();

    // 9. PSR-7 レスポンスへ書き込んで返却
    $response->getBody()->write($pdfData);
    return $response
        ->withHeader('Content-Type', 'application/pdf')
        ->withHeader('Content-Disposition', 'attachment; filename="sales_report_slim.pdf"')
        ->withHeader('Content-Length', (string)strlen($pdfData));
});

$app->run();
