<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// charts.php — チャート（グラフ）の基本: node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)

// 出力ヘルパー: node 版 stdout と同じ表記にする
$bol = fn($v) => $v ? "true" : "false";                                   // bool → "true"/"false"
$jstr = fn($s) => json_encode($s, JSON_UNESCAPED_UNICODE);                 // JSON.stringify 相当（文字列）
$jarr = fn($a) => json_encode(array_values($a), JSON_UNESCAPED_UNICODE);   // JSON.stringify 相当（配列）
// 補助: チャートタイプ数値 → enum 名
$chartTypeName = fn($code) => \phposbxl\enums\XlChartType::tryFrom($code)?->name ?? (string)$code;
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
$countCharts = function () use (&$ws) {
    $n = 0;
    for (;;) {
        try { $ws->getChartObject($n); $n++; } catch (\Throwable $e) { break; }
    }
    return $n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
$countSeries = function ($chart) {
    $n = 0;
    for (;;) {
        try { $chart->getSeriesCollection($n + 1); $n++; } catch (\Throwable $e) { break; }
    }
    return $n;
};

$app = new \phposbxl\App();
$wb = $app->createWorkBook("charts-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

// ---------- 1. データ表（2商品の月次売上）を作る ----------
foreach (["月", "商品A", "商品B"] as $c => $h) {
    $ws->getCells(1, $c + 1)->setValue($h);
}
$sales = [
    ["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
    ["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
$totalA = 0;
foreach ($sales as $i => $row) {
    $ws->getCells($i + 2, 1)->setValue($row[0]);
    $ws->getCells($i + 2, 2)->setNumberValue($row[1]);
    $ws->getCells($i + 2, 3)->setNumberValue($row[2]);
    $totalA += $row[1];
}
echo "[1] データ表 A1:C7 を作成 / 商品A合計 = {$totalA}\n";

// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
// addChartObject(名前, left, top, width, height) — 位置・サイズはポイント(pt)単位
$co1 = $ws->addChartObject("SalesColumn", 320, 10, 360, 220);
$chart1 = $co1->getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
//             styleNo, colorNo, hasLegend, title)
$chart1->chartWizard(
    "Sheet1!A1:C7",
    \phposbxl\enums\XlChartType::ChartTypeColumnClustered,
    \phposbxl\enums\XlRowCol::Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                                   // seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                                   // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                                   // chartStyleNo（0 始まり (zeroBase)）
    \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1,
    true,                                // hasLegend: 凡例を表示
    "月次売上"                           // title: 設定するとタイトルが表示される
);
echo "[2] 型 = " . $chartTypeName($chart1->getChartType())
    . " / hasTitle = " . $bol($chart1->hasTitle())
    . " / タイトル = " . $jstr($chart1->getChartTitle()->getText()) . "\n";
echo "    データ源 = {$chart1->getDataSource()}"
    . " / plotBy = {$chart1->getPlotBy()}"
    . " / hasLegend = " . $bol($chart1->hasLegend()) . "\n";

// ---------- 3. 系列・軸の読み戻し ----------
// 系列名を直接返す API は無い。系列名は chartWizard のヘッダ行なので、
// 個数は getSeriesCollection で数え、名前はヘッダセルから対応付けて確認する。
$n1 = $countSeries($chart1);
echo "[3] 系列数 = {$n1}\n";
for ($i = 1; $i <= $n1; $i++) {
    $s = $chart1->getSeriesCollection($i);
    $hv = $ws->getCells(1, $i + 1)->getValue();   // 系列 i → 列 i+1 のヘッダ
    $name = array_values($hv)[0];
    echo "    系列{$i}「{$name}」値 = " . $jarr($s->getValues())
        . " / ポイント数 = {$s->getPointSize()}\n";
}
$catAxis = $chart1->getAxes(\phposbxl\enums\XlAxisType::Category);
echo "    カテゴリ名 = " . $jarr($catAxis->getCategoryNames()) . "\n";

// ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
$co2 = $ws->addChartObject("SalesLine", 320, 250, 360, 220);
$chart2 = $co2->getChart();
$chart2->chartWizard(
    "Sheet1!A1:C7",
    \phposbxl\enums\XlChartType::ChartTypeLineMarkers,
    \phposbxl\enums\XlRowCol::Columns, 1, 1,
    1, \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
// タイトルは setTitle(true) → getChartTitle()->setText() でも設定できる
$chart2->setTitle(true);
$chart2->getChartTitle()->setText("月次売上推移");
// 値軸: 最大/最小/主単位を固定（読み戻しは getMajorUnit と ...IsAuto 系）
$valAxis = $chart2->getAxes(\phposbxl\enums\XlAxisType::Value);
$valAxis->setMaximumScale(200);
$valAxis->setMinimumScale(0);
$valAxis->setMajorUnit(50);
$valAxis->setTitle(true);
$valAxis->getTitle()->setText("売上(万円)");
// 凡例を右へ
$chart2->getLegend()->setPosition(\phposbxl\enums\XlLegendPosition::LegendPositionRight);

$right = \phposbxl\enums\XlLegendPosition::LegendPositionRight->value;
echo "[4] 型 = " . $chartTypeName($chart2->getChartType())
    . " / タイトル = " . $jstr($chart2->getChartTitle()->getText()) . "\n";
echo "    値軸 主単位 = {$valAxis->getMajorUnit()}"
    . " / 最大自動 = " . $bol($valAxis->getMaximumScaleIsAuto())
    . " / 最小自動 = " . $bol($valAxis->getMinimumScaleIsAuto()) . "\n";
echo "    値軸タイトル = " . $jstr($valAxis->getTitle()->getText())
    . " / 凡例位置 = {$chart2->getLegend()->getPosition()} (Right={$right})\n";

// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
echo "[5] チャート個数 = " . $countCharts()
    . " / SalesColumn(left,top,width,height) = "
    . implode(",", [$co1->getLeft(), $co1->getTop(), $co1->getWidth(), $co1->getHeight()]) . "\n";

// ---------- 6. チャートの削除 ----------
$ws->deleteChartObject("SalesLine");   // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
echo "[完了] 削除後のチャート個数 = " . $countCharts()
    . " / 残 = {$ws->getChartObject(0)->getName()}\n";

$wb->save();
$wb->close();