<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// af-sort.php — オートフィルターと並べ替え (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("af-sort-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

// ---------- 1. データ準備（ヘッダ + 10 行）----------
$DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
];
foreach (["商品名", "地域", "数量", "単価"] as $i => $h) {
    $col = "ABCD"[$i];
    $ws->getRange("{$col}1:{$col}1")->setValue($h);
}
foreach ($DATA as $i => $row) {
    $r = $i + 2;
    $ws->getRange("A{$r}:A{$r}")->setValue($row[0]);
    $ws->getRange("B{$r}:B{$r}")->setValue($row[1]);
    $ws->getRange("C{$r}:C{$r}")->setNumberValue($row[2]);
    $ws->getRange("D{$r}:D{$r}")->setNumberValue($row[3]);
}
echo "[1] データ準備: A1:D11（ヘッダ + 10 行）\n";

// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
$text = fn(int $r, string $col): string => $ws->getRange("{$col}{$r}:{$col}{$r}")->getValue()["{$col}{$r}"];
$label = fn(int $r, string $col): string => $text($r, "A") . "(" . $text($r, $col) . ")";
$visible = function (string $col) use ($ws, $label): string {
    $out = [];
    for ($r = 2; $r <= 11; $r++) {
        if (!$ws->getRow($r)->getHidden()) {
            $out[] = $label($r, $col);
        }
    }
    return implode(" ", $out);
};
$hiddenList = function () use ($ws): string {
    $out = [];
    for ($r = 2; $r <= 11; $r++) {
        if ($ws->getRow($r)->getHidden()) {
            $out[] = (string)$r;
        }
    }
    return count($out) > 0 ? implode(",", $out) : "なし";
};

// ---------- 2. オートフィルター設定 ----------
$af = $ws->getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
echo "[2] getAutoFilter → 範囲 = ", $af->getAddress(), " / 非表示行 = ", $hiddenList(), "\n";

// ---------- 3. 条件フィルター: 数量 > 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
$af->setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::GreaterThan, "100", 3);
echo "[3] 数量>100 の可視行: ", $visible("C"), "\n";
echo "    非表示行 = ", $hiddenList(),
    " / getValue(\"A2:A11\") の件数 = ", count($ws->getRange("A2:A11")->getValue()), "\n";
$af->resetAllFilter();                      // 条件だけ解除（フィルター自体は残る）

// ---------- 4. Top-N / 値リストのフィルター ----------
$af->setTop10ValueFilter(3, 3);             // 数量の上位 3 件
echo "[4] 数量上位3件の可視行: ", $visible("C"), "\n";
$af->resetAllFilter();
$af->setUniqueValuesNumberFilter([120, 40], 3);   // 数量が 120 または 40
echo "[5] 数量∈{120,40} の可視行: ", $visible("C"), "\n";
$af->resetAllFilter();
$af->setUniqueValuesStringFilter(["東京", "福岡"], 2);   // 地域が 東京 または 福岡（文字列リスト）
echo "[6] 地域∈{東京,福岡} の可視行: ", $visible("B"), "\n";
$af->resetAllFilter();
$af->setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
echo "[7] 地域=\"大阪\" の可視行: ", $visible("B"), "\n";
$af->resetAllFilter();

// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
$ws->getSort("A2:D11")->executeSortAscending(2);   // 数量（範囲内3列目 (0-based: 2)）で昇順
echo "[8] 数量の昇順: ", $visible("C"), "\n";

// オートフィルター経由の Sort はヘッダ行が自動的に保護される
$sort = $af->getSort();
$sort->executeSortDescending(3);                  // 単価（4列目 (0-based: 3)）で降順
echo "[9] 単価の降順: ", $visible("D"), "\n";

// ---------- 6. 並び替え条件の読み戻し ----------
$cond = $sort->getSortConditions()[0];
// PHP の bool は echo すると 1/空文字になるため JS 風の true/false に変換する
echo "[10] 条件の読み戻し: 対象列 = ", $cond->getTarget(),
    " / 昇順 = ", $cond->getSortAscending() ? "true" : "false", "\n";

// ---------- 7. オートフィルター自体の削除 ----------
$af->removeFilter();
echo "[11] removeFilter 後: 非表示行 = ", $hiddenList(), "\n";

$wb->save();
$wb->close();