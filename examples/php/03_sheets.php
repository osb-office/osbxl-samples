<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// sheets.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sheets-php.xlsx");

// PHP の bool を node の true/false 表記に合わせる
function b($x) {
    return $x ? "true" : "false";
}

// インデックスは 0 始まり (zeroBase)。0..getSheetCount()-1 を順に開いて名前を並べる
function show($label) {
    global $wb;
    $names = [];
    for ($i = 0; $i < $wb->getSheetCount(); $i++) {
        $names[] = $wb->openWorkSheet($i)->getName();
    }
    echo $label . " count=" . $wb->getSheetCount() . " [" . implode(", ", $names) . "]\n";
}

show("[1] 新規作成直後        :");

$wb->addWorkSheet("集計");                       // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :");

$wb->addWorkSheet("受注", 0);                    // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,0):");

$wb->moveSheet("受注", 2);                       // 3 番目へ移動
show("[4] moveSheet(受注,2)   :");

$ws = $wb->openWorkSheet("受注");
echo "[5] getName() = " . $ws->getName() . " / getSheetNumber(受注) = " . $wb->getSheetNumber("受注") . "\n";
$ws->getRange("A1:A1")->setValue("受注データ");

$wb->openWorkSheet("Sheet1")->setName("明細");   // リネーム
show("[6] setName(明細)       :");

// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
echo "[7] isActive: 集計 = " . b($wb->openWorkSheet("集計")->isActive())
   . " / 明細 = " . b($wb->openWorkSheet("明細")->isActive()) . "\n";
$wb->openWorkSheet("集計")->activate(true);      // 他のシートのアクティブは自動的に解除されます
echo "    activate後: 集計 = " . b($wb->openWorkSheet("集計")->isActive())
   . " / 明細 = " . b($wb->openWorkSheet("明細")->isActive()) . "\n";

// --- 削除 ---
$wb->addWorkSheet("一時");
$wb->deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");

$wb->save();
$wb->close();

// --- 開き直して確認 ---
$wb2 = $app->openWorkBook("sheets-php.xlsx");
$names = [];
for ($i = 0; $i < $wb2->getSheetCount(); $i++) {
    $names[] = $wb2->openWorkSheet($i)->getName();
}
echo "[終] シート構成: " . implode(", ", $names) . "\n";
$v = $wb2->openWorkSheet("受注")->getRange("A1:A1")->getValue();
echo "     受注!A1 = " . $v["A1"] . "\n";
$wb2->close();