<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// basic-flow.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();                       // [1] App を生成

// ---------- 新規ブックを作って編集し、保存する ----------
$wb = $app->createWorkBook("flow-php.xlsx");      // [2] 新規作成
echo "[2] createWorkBook -> getBookPath = " . $wb->getBookPath() . "\n";

$ws = $wb->openWorkSheet("Sheet1");               // [3] ワークシートを開く
echo "[3] openWorkSheet  -> シート名 = " . $ws->getName() . "\n";

// [4] Range を取得して [5] 値を設定
//     数値(行, 列)で取る $ws->getCells(0, 0) も getRange("A1:A1") と同じ Range を返します (0-based)
$ws->getRange("A1:A1")->setValue("新規作成");
$ws->getRange("B1:B1")->setFormula("100*3");      // [5] 数式を設定（先頭の "=" は付けない）
$v = $ws->getRange("B1:B1")->getValue(true);
echo "[6] getValue(true) -> B1 = " . $v["B1"] . "\n";

$wb->save();                                      // [7] 保存
echo "[7] save() 完了\n";
$wb->close();                                     // [8] 閉じる（WorkSheet の close は不要）
echo "[8] close() 完了\n";

// ---------- 既存ブックを開いて、別名で保存する ----------
$wb = $app->openWorkBook("flow-php.xlsx");        // [2'] 既存のブックを開く
$ws = $wb->openWorkSheet("Sheet1");               // [3']
$read = $ws->getRange("A1:B1")->getValue(true);
echo "[2'] openWorkBook -> A1 = " . $read["A1"] . " / B1 = " . $read["B1"] . "\n";

$ws->getRange("A2:A2")->setValue("開き直して追記");
$wb->saveAs("flow-copy-php.xlsx");                // [7'] 別名で保存
echo "[7'] saveAs() 完了 -> getBookPath = " . $wb->getBookPath() . "\n";
$wb->close();                                     // [8']
echo "[8] close() 完了\n";