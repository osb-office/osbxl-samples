<?php
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
} elseif (file_exists(__DIR__ . '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    require_once 'vendor/autoload.php';
}
// tables-comments.php — テーブル(ListObject)作成・集計行・構造化参照・セルコメント (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("tables-comments-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

// JSON.stringify 相当の出力（非 ASCII はそのまま）
$jq = fn($x): string => json_encode($x, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);

// ---------- データ投入（ヘッダ1行 + データ3行）----------
$rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", "120", "3"],
    ["ミカン", "80", "5"],
    ["ブドウ", "300", "2"],
];
$ws->setRangeData("A1", $rows);

// ---------- 1. テーブルを作成 ----------
$lo = $ws->getListObjects();
$table = $lo->addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
echo "[1] テーブル作成\n";
echo "    名前      = ", $table->getName(), "\n";
echo "    列名      = ", $jq($table->getColumnName()), "\n";
echo "    スタイル  = ", $table->getTableStyleName(), "\n";
echo "    全範囲    = ", $table->getAllRange()->getAddress(), "\n";
echo "    ヘッダ    = ", $table->getHeaderRowRange()->getAddress(), "\n";
echo "    データ部  = ", $table->getDataBodyRange()->getAddress(), "\n";
echo "    集計行    = ", $table->getTotalRowRange()->getAddress(), "\n";

// ---------- 2. 集計行 ----------
$table->setTotalRowLabel(0, "合計");
$table->setTotalRowFunction(1, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
$table->setTotalRowFunction(2, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
echo "[2] 集計行\n";
echo "    ラベル(1列目 (0-based: 0)) = ", $jq($table->getTotalRowLabel(1)), "\n";
echo "    数式          = ", $jq($ws->getRange("B5:C5")->getFormula()), "\n";
echo "    値            = ", $jq($ws->getRange("B5:C5")->getValue(true)), "\n";

// ---------- 3. テーブルスタイルを変更 ----------
$table->setBuiltinStyleName(\phposbxl\enums\XlDefaultTableStyle::TableStyleMedium9, true);
echo "[3] スタイル変更後 = ", $table->getTableStyleName(), "\n";

// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
$ws->getRange("E1:E1")->setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
echo "[4] 構造化参照\n";
echo "    E1 数式 = ", $jq($ws->getRange("E1:E1")->getFormula()), "\n";
echo "    E1 値   = ", $jq($ws->getRange("E1:E1")->getValue(true)), "\n";

// ---------- 5. コメント（新式）とメモ（旧式）----------
$cm = $ws->getComments();
$mkComment = function (string $author, string $content): \phposbxl\dto\CommentObject {
    $c = new \phposbxl\dto\CommentObject();
    $c->setAuthor($author);
    $c->setContent($content);
    return $c;
};

// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
$cm->setCommentThread("B2", [$mkComment("alice", "この数値を確認"), $mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
$cm->setComment("D2", $mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
$cm->setMemo("A2", "Carol", "最優先商品です", new \phposbxl\dto\FontObject());

$thread = $cm->getComment("B2");
echo "[5] コメント\n";
echo "    B2 スレッド件数 = ", count($thread), "\n";
foreach ($thread as $i => $c) {
    echo "      [{$i}] ", $c->getAuthor(), ": ", $c->getContent(), $c->getParentId() ? " (返信)" : "", "\n";
}
echo "    D2 コメント件数 = ", count($cm->getComment("D2")), "\n";
echo "    A2 メモ = ", $jq($cm->getMemoText("A2")), " / 作者 ", $jq($cm->getMemoAuthor("A2")), "\n";

// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
$cm->removeComment("D2");
$cm->removeMemo("A2");
$tryRead = function (string $label, callable $fn): void {
    try {
        $fn();
        echo "    削除後 {$label}: 例外なし\n";
    } catch (\Exception $e) {
        echo "    削除後 {$label}: 例外 ", $e->getMessage(), "\n";
    }
};
$tryRead("D2コメント", fn() => $cm->getComment("D2"));
$tryRead("A2メモ", fn() => $cm->getMemoText("A2"));
echo "    B2 は残すので save 後も ", count($cm->getComment("B2")), "件\n";

$total = $ws->getRange("B5:B5")->getValue(true)["B5"];
$e1 = $ws->getRange("E1:E1")->getValue(true)["E1"];
echo "[確認] テーブル=", $table->getName(), " 集計(単価)=", $total, " 構造化参照(E1)=", $e1, "\n";

$wb->save();
$wb->close();