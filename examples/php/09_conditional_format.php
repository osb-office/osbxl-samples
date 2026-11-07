<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// conditional-format.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)

// 出力ヘルパー: node 版 stdout と同じ表記にする (bool → "true"/"false")
$bol = fn($v) => $v ? "true" : "false";
// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
$color = function (string $hex): \phposbxl\dto\ColorObject {
    $c = new \phposbxl\dto\ColorObject();
    $c->setHexColor($hex);
    return $c;
};

$app = new \phposbxl\App();
$wb = $app->createWorkBook("condfmt-php.xlsx"); // 出力文言は node 版に合わせ condfmt.xlsx と表示
$ws = $wb->openWorkSheet("Sheet1");

// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
foreach (["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"] as $i => $h) {
    $col = "ABCDE"[$i];
    $ws->getRange("{$col}1:{$col}1")->setValue($h);
}
$sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
foreach ($sales as $i => $v) {
    $r = $i + 2;
    $ws->getRange("A{$r}:A{$r}")->setNumberValue($i + 1);
    foreach (["B", "C", "D", "E"] as $col) {
        $ws->getRange("{$col}{$r}:{$col}{$r}")->setNumberValue($v);
    }
}

$fc = $ws->getFormatConditions();

// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
$cellVal = new \phposbxl\dto\CellValueObject();
$cellVal->setValueCondition(60, \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater);
$cellVal->setA1C1("B2:B11");
$redFont = new \phposbxl\dto\FontObject();
$redFont->setBold(true);
$redFont->setColorObject($color("9C0006"));
$cellVal->setFont($redFont);
$pinkFill = new \phposbxl\dto\PatternFillObject();
$pinkFill->setPattern(\phposbxl\enums\XlPatternType::PatternSolid);
$pinkFill->setCellColor($color("FFC7CE"));
$cellVal->setPatternFill($pinkFill);
$fc->setCellValueCondition(1, $cellVal);

// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
$colorScale = new \phposbxl\dto\ColorScaleObject();
$colorScale->setA1C1("C2:C11");
$colorScale->setMinimumCondition($color("F8696B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$colorScale->setMediumCondition($color("FFEB84"), \phposbxl\enums\XlConditionValueType::ConditionValueTypePercentile, 50);
$colorScale->setMaximumCondition($color("63BE7B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc->setColorScaleFormatCondition(2, $colorScale);

// ---------- (3) データバー: D列 ----------
$dataBar = new \phposbxl\dto\DataBarObject();
$dataBar->setA1C1("D2:D11");
$dataBar->setDataBarColor($color("638EC6"));
$dataBar->setMinimumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$dataBar->setMaximumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc->setDataBarFormatCondition(3, $dataBar);

// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
$top3 = new \phposbxl\dto\Top10Object();
$top3->setRank(3, true);                      // true = 上位（false なら下位）
$top3->setA1C1("E2:E11");
$blueFont = new \phposbxl\dto\FontObject();
$blueFont->setBold(true);
$blueFont->setColorObject($color("0070C0"));
$top3->setFont($blueFont);
$fc->setTop10FormatCondition(4, $top3);

$aboveAvg = new \phposbxl\dto\AboveAverageObject();
$aboveAvg->setAboveBelow(\phposbxl\enums\XlAboveBelow::XlAboveAverage);
$aboveAvg->setA1C1("E2:E11");
$greenFont = new \phposbxl\dto\FontObject();
$greenFont->setBold(true);
$greenFont->setColorObject($color("006100"));
$aboveAvg->setFont($greenFont);
$fc->setAboveAverageCondition(5, $aboveAvg);

$fc->apply();                                 // 全条件の設定後に apply() が必須
echo "[1] 5 件のルールを登録し apply() しました\n";

// ---------- ルール一覧の読み戻し ----------
$types = $fc->getConditionTypes();            // array<int, array<string, int>>
ksort($types);                                // node の Object.entries と同じく priority 昇順で走査する
echo "[2] ルール数 = " . count($types) . "\n";
foreach ($types as $prio => $ranges) {
    foreach ($ranges as $range => $type) {
        $name = \phposbxl\enums\XlFormatConditionType::from($type)->name;
        echo "    priority {$prio}: {$range} → {$name} (type={$type})\n";
    }
}

// ---------- 各ルールの内容の読み戻し ----------
echo "[3] ルール内容の読み戻し\n";
$g1 = $fc->getCellValueCondition(1);
$greater = \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater->value;
echo "    セル値  : {$g1->getA1C1()} 値={$g1->getConditionValue()}"
    . " 演算子={$g1->getConditionOperator()} (Greater={$greater})"
    . " 文字色={$g1->getFont()->getColorObject()->getHexColor()} 太字=" . $bol($g1->getFont()->isBold())
    . " 背景={$g1->getPatternFill()->getCellColor()->getHexColor()}\n";
$g2 = $fc->getColorScaleFormatCondition(2);
echo "    カラースケール: min={$g2->getMinimumColor()->getHexColor()} (type={$g2->getMinimumValueType()})"
    . " mid={$g2->getMediumColor()->getHexColor()} (type={$g2->getMediumValueType()}, 閾値={$g2->getMediumValue()})"
    . " max={$g2->getMaximumColor()->getHexColor()} (type={$g2->getMaximumValueType()})\n";
$g3 = $fc->getDataBarFormatCondition(3);
echo "    データバー: 色={$g3->getDataBarColor()->getHexColor()}"
    . " minType={$g3->getMinimumCondition()} maxType={$g3->getMaximumCondition()}\n";
$g4 = $fc->getTop10FormatCondition(4);
echo "    上位 N 件: rank={$g4->getRank()} isTop=" . $bol($g4->getTopBottom()) . "\n";
$g5 = $fc->getAboveAverageCondition(5);
$xlAbove = \phposbxl\enums\XlAboveBelow::XlAboveAverage->value;
echo "    平均比較: aboveBelow={$g5->getAboveBelow()} (XlAboveAverage={$xlAbove})\n";

// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
$byRow = fn($a, $b) => (int)substr($a, 1) - (int)substr($b, 1);
echo "[4] 書式が適用されたセル\n";
$appliedB = array_keys($fc->getAppliedFontObject("B2:B11"));
usort($appliedB, $byRow);
echo "    B列 (>60)     : " . implode(" ", $appliedB) . "\n";
$appliedE = array_keys($fc->getAppliedFontObject("E2:E11"));
usort($appliedE, $byRow);
echo "    E列 (上位/平均): " . implode(" ", $appliedE) . "\n";
$appliedBar = $fc->getAppliedDataBar("D2:D11");
echo "    D列 データバー : " . count(array_filter($appliedBar)) . " セル\n";

// ---------- 削除 ----------
$fc->removeCondition(4);                      // 上位 3 件のルールだけ削除
echo "[5] removeCondition(4) 後のルール数 = " . count($fc->getConditionTypes()) . "\n";
$fc->removeAllCondition();                    // 全ルールを削除
echo "[6] removeAllCondition() 後のルール数 = " . count($fc->getConditionTypes()) . "\n";

// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
$fc->setCellValueCondition(1, $cellVal);
$fc->apply();

$wb->save();
$wb->close();
echo "[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）\n";