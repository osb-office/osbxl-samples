<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// pivot-sample.php — ピボットテーブル: node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("pivot-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
$headers = ["地域", "商品", "数量"];
$srcData = [
  ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
  ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
  ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
  ["北", "りんご", 20],
];
foreach ($headers as $i => $h) {
  $col = chr(65 + $i);
  $ws->getRange("{$col}1:{$col}1")->setValue($h);
}
$srcTotal = 0;
foreach ($srcData as $ri => $row) {
  foreach ($row as $ci => $v) {
    $col = chr(65 + $ci);
    $r = $ri + 2;
    if (is_int($v)) $ws->getRange("{$col}{$r}:{$col}{$r}")->setNumberValue($v);
    else $ws->getRange("{$col}{$r}:{$col}{$r}")->setValue($v);
  }
  $srcTotal += $row[2];
}
echo "[1] ソースデータ: " . count($srcData) . " 行, 数量の合計 = {$srcTotal}\n";

// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
$pt = $ws->getPivotTables()->addPivotTable("売上集計", "A1:C11", "E1");

$rf = $pt->getField("地域");
$rf->setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeRow);
$df = $pt->getField("数量");
$df->setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeData);
$df->setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeSum);

$pt->setFields($pt->getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
               [$rf], [], [], [$df]);
$pt->refreshData();
$pt->refresh();
echo "[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)\n";

// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
// getValue() は連想配列（キー=セルアドレス、値=表示値の文字列）を返す
$pivotRows = [];
for ($r = 2; $r <= 10; $r++) {
  $label = $ws->getRange("E{$r}:E{$r}")->getValue()["E{$r}"] ?? null;
  $val   = $ws->getRange("F{$r}:F{$r}")->getValue()["F{$r}"] ?? null;
  if ($label !== null && $label !== "") {
    $pivotRows[$label] = (float)$val;
    echo "    {$label} = {$val}\n";
  }
}
$grandTotal = $pivotRows["総計"];
// node の Number 表示（205）に合わせるため、整数なら int で表示する
$gt = ($grandTotal == (int)$grandTotal) ? (string)(int)$grandTotal : (string)$grandTotal;
echo "[3] 検算: 総計 {$gt} === ソース合計 {$srcTotal} → " . ($grandTotal == $srcTotal ? "OK" : "NG") . "\n";

// ---------- 4. 集計方法を平均に変更 ----------
$df->setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeAverage);
$pt->setFields($pt->getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
               [$rf], [], [], [$df]);
$pt->refreshData();
$pt->refresh();
echo "[4] 集計方法を合計→平均に変更\n";
for ($r = 2; $r <= 10; $r++) {
  $label = $ws->getRange("E{$r}:E{$r}")->getValue()["E{$r}"] ?? null;
  $val   = $ws->getRange("F{$r}:F{$r}")->getValue()["F{$r}"] ?? null;
  if ($label !== null && $label !== "") {
    echo "    {$label} = {$val}\n";
  }
}

// ---------- 5. フィールド情報の取得 ----------
$fields = $pt->getFields();
$caps = array_map(fn($f) => $f->getCaption(), $fields);
echo "[5] getFields: " . count($fields) . " 個 → " . implode(", ", $caps) . "\n";
$tabular = \phposbxl\enums\XlPivotTableLayoutType::TabularForm->value;
echo "    getLayoutType = {$pt->getLayoutType()} (TabularForm={$tabular})\n";

// ---------- 6. ピボットテーブルの削除 ----------
$ws->getPivotTables()->removePivotTable("売上集計");
try {
  $ws->getPivotTables()->getPivotTable("売上集計");
  echo "[6] 削除失敗\n";
} catch (\Throwable $e) {
  echo "[6] removePivotTable → 削除確認OK\n";
}

$wb->save();
$wb->close();