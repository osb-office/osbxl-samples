<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// rows-cols.php — 行・列・セル範囲の挿入/削除/クリア/コピー (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("rows-cols-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
$dump = function (string $label, string $range) use ($ws, $app): void {
    $v = $ws->getRange($range)->getValue();
    $keys = array_keys($v);
    usort($keys, function (string $a, string $b) use ($app): int {
        preg_match('/^([A-Z]+)(\d+)$/', $a, $ma);
        preg_match('/^([A-Z]+)(\d+)$/', $b, $mb);
        return ((int)$ma[2] <=> (int)$mb[2])
            ?: ($app->convertToColumnNumber($ma[1]) <=> $app->convertToColumnNumber($mb[1]));
    });
    $parts = [];
    foreach ($keys as $k) {
        $val = (string)$v[$k];
        $parts[] = $k . "=" . ($val === "" ? "·" : $val);
    }
    echo $label, " ", implode(" ", $parts), "\n";
};
// A1:C5 を r{行}c{列} で埋める
$fill = function () use ($ws): void {
    for ($r = 1; $r <= 5; $r++) {
        for ($c = 1; $c <= 3; $c++) {
            $ws->getCells($r, $c)->setValue("r{$r}c{$c}");
        }
    }
};

$fill();
$dump("[初期]       A列:", "A1:A5");

// ---------- 行 ----------
$ws->insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
$dump("[insertRow]  A列:", "A1:A7");
$ws->deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
$dump("[deleteRow]  A列:", "A1:A5");
$ws->clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
$dump("[clearRow]   A列:", "A1:A5");
$ws->copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
$dump("[copyRow]    A列:", "A1:A5");

$fill();
// ---------- 列 ----------
$ws->insertCol(2, 1);                    // B 列の前に 1 列挿入
$dump("[insertCol] 1行目:", "A1:D1");
$ws->deleteCol(2, 1);                    // 挿入した B 列を削除
$dump("[deleteCol] 1行目:", "A1:C1");
$ws->clearCol(3, 1);                     // C 列をクリア
$dump("[clearCol]  1行目:", "A1:C1");
$fill();
$ws->copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
$dump("[copyCol]   1行目:", "A1:F1");

// ---------- セル範囲 ----------
$ws->insertCell("B2:B3", \phposbxl\enums\XlInsertDirection::InsertDirectionDown);
$dump("[insertCell] B列:", "B1:B5");
$ws->deleteCell("B2:B3", \phposbxl\enums\XlInsertDirection::DeleteDirectionUp);
$dump("[deleteCell] B列:", "B1:B5");
$ws->clearCell("A1:B2");
$dump("[clearCell] A1:B2:", "A1:B2");
$ws->copyCell("Sheet1", "A3:B4", "F8");
$dump("[copyCell]  F8:G9:", "F8:G9");

$wb->save();
$wb->close();