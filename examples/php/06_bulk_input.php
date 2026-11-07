<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// bulk-input.php — 2次元配列による値と数式の一括高速投入 (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("bulk-input.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

$now = fn(): int => (int)(hrtime(true) / 1000000);   // ミリ秒

// ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
$data = [
    ["商品名", "単価", "数量", "金額"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
];
$ws->setRangeData("A1", $data);

echo "[1] setRangeData 投入後:\n";
foreach (["A1", "B2", "C2", "D1"] as $a1) {
    echo "    {$a1} = ", json_encode($ws->getRange($a1)->getValue()[$a1], JSON_UNESCAPED_UNICODE), "\n";
}

// ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
$formulas = [
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
$ws->setRangeFormulas("D2", $formulas);

echo "[2] setRangeFormulas 投入後:\n";
echo "    数式 D2:D4 = ", json_encode($ws->getRange("D2:D4")->getFormula(), JSON_UNESCAPED_UNICODE), "\n";
echo "    計算値 D2:D4 = ", json_encode($ws->getRange("D2:D4")->getValue(true), JSON_UNESCAPED_UNICODE), "\n";

// ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
$ROWS = 2000;
$COLS = 5;

// セルごとに getCells() して投入
$wsCell = $wb->addWorkSheet("CellByCell");
$t = $now();
for ($r = 1; $r <= $ROWS; $r++) {
    for ($c = 1; $c <= $COLS; $c++) $wsCell->getCells($r, $c)->setNumberValue($r * 100 + $c);
}
$msCell = $now() - $t;

// 2次元配列で setRangeData() 投入 (最高速)
$wsBulk = $wb->addWorkSheet("Bulk");
$t = $now();
$matrix = [];
for ($r = 1; $r <= $ROWS; $r++) {
    $row = [];
    for ($c = 1; $c <= $COLS; $c++) $row[] = (string)($r * 100 + $c);
    $matrix[] = $row;
}
$msBuild = $now() - $t;
$t = $now();
$wsBulk->setRangeData("A1", $matrix);
$msApply = $now() - $t;

echo "[3] ", $ROWS * $COLS, " セルの投入速度比較:\n";
echo "    セルごと getCells().setNumberValue() : {$msCell} ms\n";
echo "    setRangeData (構築{$msBuild} + 投入{$msApply}) : ", $msBuild + $msApply, " ms (最速)\n";

$a = $wsCell->getRange("E2000:E2000")->getValue(true)["E2000"];
$b = $wsBulk->getRange("E2000:E2000")->getValue(true)["E2000"];
echo "[検証] 両方式とも E2000 = ", $a === (int)$b ? $b : "不一致 {$a} / {$b}", "\n";

$wb->save();
$wb->close();