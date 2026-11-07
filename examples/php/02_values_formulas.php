<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// values-formulas.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sales-php.xlsx");   // 新規ブック作成
$ws = $wb->openWorkSheet("Sheet1");             // 既定のシートを開く

// --- 見出し (1行目) ---
// $app->convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
$headers = ["商品", "単価", "個数", "金額"];
foreach ($headers as $i => $text) {
    $cell = $app->convertFromRowColNumber(1, $i + 1);
    $ws->getRange("{$cell}:{$cell}")->setValue($text);
}

// --- 明細 (2〜4行目) ---
$data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
foreach ($data as $i => [$name, $price, $qty]) {
    $row = $i + 2;
    $ws->getRange("A{$row}:A{$row}")->setValue($name);
    $ws->getRange("B{$row}:B{$row}")->setNumberValue($price);
    $ws->getRange("C{$row}:C{$row}")->setNumberValue($qty);
}

// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
// 数式文字列に先頭の "=" は付けない。
$formulas = [];
for ($i = 0; $i < count($data); $i++) {
    $row = $i + 2;
    $formulas[] = ["B{$row}*C{$row}"];
}
$ws->setRangeFormulas("D2", $formulas);

// --- 合計 (D5) ---
$ws->getRange("C5:C5")->setValue("合計");
$ws->getRange("D5:D5")->setFormula("SUM(D2:D4)");

// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
$values = $ws->getRange("A1:D5")->getValue(true);
echo "金額 D2:D4 = " . implode(", ", [$values["D2"], $values["D3"], $values["D4"]]) . "\n";
echo "合計: " . $values["D5"] . "\n";

// --- 保存して閉じる ---
$wb->save();
$wb->close();