<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// value-types.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("value-types-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

// JS の JSON.stringify 相当の引用
function jsonq($s) {
    return json_encode($s, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
}

// 整数値の float は int 表記にする（JS の出力と揃える）
function num($x) {
    return $x == floor($x) ? (string)(int)$x : (string)$x;
}

// PHP の bool を JS 風の true/false で表示する
function boolstr($x) {
    return $x ? "true" : "false";
}

function show($a1) {
    global $ws;
    $r = $ws->getRange("{$a1}:{$a1}");
    echo "  {$a1}: 表示値=" . jsonq($r->getValue()[$a1])
       . " 生値=" . jsonq($r->getValue(true)[$a1])
       . " 書式=" . jsonq($r->getNumberFormat()[$a1]) . "\n";
}

// --- 日付・時刻 ---
// DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
$date = new \phposbxl\dto\DateTimeObject();
$date->setYMD(2024, 1, 15);

$datetime = new \phposbxl\dto\DateTimeObject();
$datetime->setYMD(2024, 1, 15);
$datetime->setHMS(13, 45, 30);

$time = new \phposbxl\dto\DateTimeObject();
$time->setHMS(9, 5, 0);

echo "[日付・時刻]\n";
$ws->getRange("A1:A1")->setDateValue($date);                       // 日付のみ
$ws->getRange("A2:A2")->setDateValue($datetime);                   // 日付 + 時刻
$ws->getRange("A3:A3")->setDateValue($time);                       // 時刻のみ
$ws->getRange("A4:A4")->setDateStringValue("2024-01-15");          // 文字列から
$ws->getRange("A5:A5")->setDateStringValue("令和6年1月15日");        // 和暦も認識されます
$ws->getRange("A6:A6")->setDateValue($date, false, "yyyy-mm-dd");   // 表示書式を同時に指定
$ws->getRange("A7:A7")->setDateValue($date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
foreach (["A1", "A2", "A3", "A4", "A5", "A6", "A7"] as $a) show($a);

// --- 数値・文字列 ---
echo "[数値・文字列]\n";
$ws->getRange("B1:B1")->setNumberValue(1234567.891);
$ws->getRange("B2:B2")->setNumberValue(1234567.891, false, "#,##0.00");
$ws->getRange("B3:B3")->setNumberValue(0.25, false, "0.0%");
$ws->getRange("B4:B4")->setValue("00123");       // 先頭 0 は文字列として保持される
$ws->getRange("B5:B5")->setValue("2024/1/15");   // 汎用 setter は日付として自動認識
$ws->getRange("B6:B6")->setValue("12345", true); // forceString=true で強制的に文字列扱い
$ws->getRange("B7:B7")->setNumberValue(1234.5, false, '"円"#,##0');      // 引用符リテラル
$ws->getRange("B8:B8")->setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
foreach (["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"] as $bx) show($bx);

// --- 真偽値 ---
echo "[真偽値]\n";
$ws->getRange("C1:C1")->setBooleanValue(true);
$ws->getRange("C2:C2")->setBooleanValue(false);
$ws->getRange("C3:C3")->setFormula("TRUE");      // 数式でも真偽値になる
foreach (["C1", "C2", "C3"] as $c) show($c);

// --- 日付のシリアル値 ---
echo "[シリアル値]\n";
echo "  2024/1/15 のシリアル値 = " . num($app->getNumericValue($date)) . "\n";
echo "  1904 年形式の場合      = " . num($app->getNumericValue($date, true)) . "\n";
echo "  wb.isDate1904()        = " . boolstr($wb->isDate1904()) . "\n";

$wb->save();
$wb->close();