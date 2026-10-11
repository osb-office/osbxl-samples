# テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ

level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-08-tables-comments.md)

## やりたいこと

データ範囲を Excel の「テーブル」（`ListObject`）にして、名前・列名・スタイル・集計行を扱う。 テーブル列を数式から参照する「構造化参照」を使い、あわせてセルへのコメント（新式スレッド）と メモ（旧式吹き出し）の追加・読み取り・削除を行う。

## 対象クラス

テーブルは `ws.getListObjects()` から、コメントは `ws.getComments()` から取得します。

### `ListObjects`（テーブルのコレクション）

| メソッド | 説明 |
| --- | --- |
| `addList(name, A1C1, useFirstRowAsHeader, insertTotals)` | 既存のセル範囲からテーブルを作成し、`Table` を返す。 |
| `addListFromRange(name, topA1, copyFromA1C1, useFirstRowAsHeader, insertTotals)` | データを `topA1` へ**コピーして**テーブルを作成。`copyFromA1C1` は `Sheet2!A1:C1` のように他シートも可。 |
| `getList(name)` | テーブルを取得。無い場合は `table is not found` 例外。 |
| `removeList(name, deleteData?)` | テーブルを削除。`deleteData` 既定は `true`（**データも消える**）。残すなら `false`。 |

### `Table`（テーブル 1 つ）

| 分類 | メソッド |
| --- | --- |
| 範囲 | `getAllRange()` / `getHeaderRowRange()` / `getDataBodyRange()` / `getTotalRowRange()`（いずれも `Range` を返す。集計行なしで `getTotalRowRange()` は `totalRow is not set` 例外） |
| 名前 | `getName()` / `changeTableName(name)` |
| 列名 | `getColumnName()` / `setColumnName([...])`（ヘッダセルの表示も変わる） |
| スタイル | `getTableStyleName()` / `setBuiltinStyleName(enums.XlDefaultTableStyle.Xxx, clearFormat)` / `setCustomTableStyleName(name, clearFormat)` |
| 集計行 | `isShowTotals()` / `setShowTotals(bool)` / `getTotalRowFunction(col)` / `setTotalRowFunction(col, enums.XlTotalsCalculation.Xxx)` / `setCustomRowFunction(col, formula, isArray?)` / `getTotalRowLabel(col)` / `setTotalRowLabel(col, text)` |
| フィルター・表示 | `getAutoFilter()` / `removeAutoFilter()` / `isShowAutoFilter()` / `isShowHeaders()` / `setShowHeaders(bool)` / `isShowTableStyleRowStripes()` ほか |

### `Comments`（セルコメント）

Excel にはコメントが 2 系統あります。**新式**（スレッド・返信・完了フラグ）と**旧式メモ**（吹き出し）です。

| メソッド | 系統 | 説明 |
| --- | --- | --- |
| `setComment(A1, commentObject)` | 新式 | 1 件追加。既にコメントがあるセルでは**追加（自動で返信化）**。 |
| `setCommentThread(A1, [commentObject...])` | 新式 | スレッドを**一括設定（上書き）**。先頭が本体、以降が返信。 |
| `getComment(A1)` | 新式 | `Array<dto.CommentObject>` を返す。無ければ例外 `comment not found`。 |
| `removeComment(A1)` | 新式 | 新式コメントを削除。 |
| `setMemo(A1, author, text, fontObject, visible?, rowColumnsNum?, colColumnsNum?)` | 旧式 | 吹き出しメモを設定。`fontObject` は `new dto.FontObject()` で可。 |
| `getMemoText(A1)` / `getMemoAuthor(A1)` | 旧式 | メモ本文 / 作者を取得。無ければ例外 `memo not found`。 |
| `removeMemo(A1)` | 旧式 | メモを削除。 |

`dto.CommentObject` は `setAuthor` / `setContent` / `setParentId`（返信の親ID）/ `setDone`（完了フラグ）と それぞれの getter を持ちます。ID は設定すると `{...}` 形式の UUID が自動採番されます。

### 集計方法 `enums.XlTotalsCalculation`

| メンバ | 数値 | 意味 |
| --- | --- | --- |
| `TotalsCalculationNone` | 0 | 集計なし |
| `TotalsCalculationSum` | 1 | 合計 |
| `TotalsCalculationAverage` | 2 | 平均 |
| `TotalsCalculationCount` | 3 | 件数 |
| `TotalsCalculationCountNums` | 4 | 数値の個数 |
| `TotalsCalculationMin` / `TotalsCalculationMax` | 5 / 6 | 最小 / 最大 |
| `TotalsCalculationStdDev` / `TotalsCalculationVar` | 7 / 8 | 標準偏差 / 分散 |
| `TotalsCalculationCustom` | 9 | カスタム数式（`setCustomRowFunction` で設定） |

## コード

実行結果:

#### [Node.js]

```javascript
// tables-comments.js — テーブル(ListObject)作成・集計行・構造化参照・セルコメント
const { nodeosbxl, enums, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("tables-comments.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- データ投入（ヘッダ1行 + データ3行）----------
const rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", "120", "3"],
    ["ミカン", "80", "5"],
    ["ブドウ", "300", "2"],
];
ws.setRangeData("A1", rows);
// ---------- 1. テーブルを作成 ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] テーブル作成");
console.log(` 名前 = ${table.getName()}`);
console.log(` 列名 = ${JSON.stringify(table.getColumnName())}`);
console.log(` スタイル = ${table.getTableStyleName()}`);
console.log(` 全範囲 = ${table.getAllRange().getAddress()}`);
console.log(` ヘッダ = ${table.getHeaderRowRange().getAddress()}`);
console.log(` データ部 = ${table.getDataBodyRange().getAddress()}`);
console.log(` 集計行 = ${table.getTotalRowRange().getAddress()}`);
// ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] 集計行");
console.log(` ラベル(1列目) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(` 数式 = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(` 値 = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);
// ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] スタイル変更後 = ${table.getTableStyleName()}`);
// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
console.log("[4] 構造化参照");
console.log(` E1 数式 = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(` E1 値 = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);
// ---------- 5. コメント（新式）とメモ（旧式）----------
const cm = ws.getComments();
const mkComment = (author, content) => {
const c = new dto.CommentObject();
c.setAuthor(author);
c.setContent(content);
return c;
};
// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", new dto.FontObject());
const thread = cm.getComment("B2");
console.log("[5] コメント");
console.log(` B2 スレッド件数 = ${thread.length}`);
thread.forEach((c, i) => console.log(` [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (返信)" : ""}`));
console.log(` D2 コメント件数 = ${cm.getComment("D2").length}`);
console.log(` A2 メモ = ${JSON.stringify(cm.getMemoText("A2"))} / 作者 ${JSON.stringify(cm.getMemoAuthor("A2"))}`);
// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) => {
try { fn(); console.log(` 削除後 ${label}: 例外なし`); }
catch (e) { console.log(` 削除後 ${label}: 例外 ${e.message}`); }
};
tryRead("D2コメント", () => cm.getComment("D2"));
tryRead("A2メモ", () => cm.getMemoText("A2"));
console.log(` B2 は残すので save 後も ${cm.getComment("B2").length}件`);
const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[確認] テーブル=${table.getName()} 集計(単価)=${total} 構造化参照(E1)=${e1}`);
wb.save();
wb.close();
```

#### [Python]

```python
# tables-comments.py — node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl
def jdump(obj): # JSON.stringify 相当の出力（区切り空白なし・非 ASCII そのまま）
return json.dumps(obj, ensure\_ascii=False, separators=(",", ":"))
app = pyosbxl.App()
wb = app.createWorkBook("tables-comments-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- データ投入（ヘッダ1行 + データ3行）----------
rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", "120", "3"],
    ["ミカン", "80", "5"],
    ["ブドウ", "300", "2"],
]
ws.setRangeData("A1", rows)
# ---------- 1. テーブルを作成 ----------
lo = ws.getListObjects()
table = lo.addList("Sales", "A1:C4", True, True) # useFirstRowAsHeader, insertTotals
print("[1] テーブル作成")
print(f" 名前 = {table.getName()}")
print(f" 列名 = {jdump(table.getColumnName())}")
print(f" スタイル = {table.getTableStyleName()}")
print(f" 全範囲 = {table.getAllRange().getAddress()}")
print(f" ヘッダ = {table.getHeaderRowRange().getAddress()}")
print(f" データ部 = {table.getDataBodyRange().getAddress()}")
print(f" 集計行 = {table.getTotalRowRange().getAddress()}")
# ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計")
table.setTotalRowFunction(2, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
table.setTotalRowFunction(3, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
print("[2] 集計行")
print(f" ラベル(1列目) = {jdump(table.getTotalRowLabel(1))}")
print(f" 数式 = {jdump(ws.getRange('B5:C5').getFormula())}")
print(f" 値 = {jdump(ws.getRange('B5:C5').getValue(True))}")
# ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(pyosbxl.enums.XlDefaultTableStyle.TableStyleMedium9, True)
print(f"[3] スタイル変更後 = {table.getTableStyleName()}")
# ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])") # 先頭に = は付けない
print("[4] 構造化参照")
print(f" E1 数式 = {jdump(ws.getRange('E1:E1').getFormula())}")
print(f" E1 値 = {jdump(ws.getRange('E1:E1').getValue(True))}")
# ---------- 5. コメント（新式）とメモ（旧式）----------
cm = ws.getComments()
def mkComment(author, content):
c = pyosbxl.dto.CommentObject()
c.setAuthor(author)
c.setContent(content)
return c
# 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")])
# 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"))
# 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", pyosbxl.dto.FontObject())
thread = cm.getComment("B2")
print("[5] コメント")
print(f" B2 スレッド件数 = {len(thread)}")
for i, c in enumerate(thread):
print(f" [{i}] {c.getAuthor()}: {c.getContent()}{' (返信)' if c.getParentId() else ''}")
print(f" D2 コメント件数 = {len(cm.getComment('D2'))}")
print(f" A2 メモ = {jdump(cm.getMemoText('A2'))} / 作者 {jdump(cm.getMemoAuthor('A2'))}")
# 削除: 新式は removeComment / 旧式は removeMemo
cm.removeComment("D2")
cm.removeMemo("A2")
# 削除後の読み取りは node 版と同様に "comment not found" / "memo not found" 例外になる
def tryRead(label, fn):
try:
fn()
print(f" 削除後 {label}: 例外なし")
except Exception as e:
print(f" 削除後 {label}: 例外 {e}")
tryRead("D2コメント", lambda: cm.getComment("D2"))
tryRead("A2メモ", lambda: cm.getMemoText("A2"))
print(f" B2 は残すので save 後も {len(cm.getComment('B2'))}件")
total = ws.getRange("B5:B5").getValue(True)["B5"]
e1 = ws.getRange("E1:E1").getValue(True)["E1"]
print(f"[確認] テーブル={table.getName()} 集計(単価)={total} 構造化参照(E1)={e1}")
wb.save()
wb.close()
```

#### [Java]

```java
// tables-comments.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
public class Main {
// JSON.stringify 相当の出力ヘルパー（Map は挿入順のまま）
static String jsonMap(Map<String, String> m) {
StringBuilder sb = new StringBuilder("{");
boolean first = true;
for (Map.Entry<String, String> e : m.entrySet()) {
if (!first) sb.append(",");
first = false;
sb.append("\"").append(e.getKey()).append("\":\"").append(e.getValue()).append("\"");
}
return sb.append("}").toString();
}
static String jsonList(List<String> l) {
StringBuilder sb = new StringBuilder("[");
for (int i = 0; i < l.size(); i++) {
if (i > 0) sb.append(",");
sb.append("\"").append(l.get(i)).append("\"");
}
return sb.append("]").toString();
}
static CommentObjectWrapper mkComment(String author, String content) {
CommentObjectWrapper c = new CommentObjectWrapper();
c.setAuthor(author);
c.setContent(content);
return c;
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("tables-comments-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
        // ---------- データ投入（ヘッダ1行 + データ3行）----------
        List<List<String>> rows = Arrays.asList(
            Arrays.asList("商品名", "単価", "数量"),
            Arrays.asList("リンゴ", "120", "3"),
            Arrays.asList("ミカン", "80", "5"),
            Arrays.asList("ブドウ", "300", "2")
        );
        ws.setRangeData("A1", rows);
        // ---------- 1. テーブルを作成 ----------
ListObjectsWrapper lo = ws.getListObjects();
TableWrapper table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
System.out.println("[1] テーブル作成");
System.out.println(" 名前 = " + table.getName());
System.out.println(" 列名 = " + jsonList(table.getColumnName()));
System.out.println(" スタイル = " + table.getTableStyleName());
System.out.println(" 全範囲 = " + table.getAllRange().getAddress());
System.out.println(" ヘッダ = " + table.getHeaderRowRange().getAddress());
System.out.println(" データ部 = " + table.getDataBodyRange().getAddress());
System.out.println(" 集計行 = " + table.getTotalRowRange().getAddress());
// ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計");
table.setTotalRowFunction(2, XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, XlTotalsCalculation.TotalsCalculationSum);
System.out.println("[2] 集計行");
System.out.println(" ラベル(1列目) = \"" + table.getTotalRowLabel(1) + "\"");
System.out.println(" 数式 = " + jsonMap(ws.getRange("B5:C5").getFormula()));
System.out.println(" 値 = " + jsonMap(ws.getRange("B5:C5").getValue(true)));
// ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(XlDefaultTableStyle.TableStyleMedium9, true);
System.out.println("[3] スタイル変更後 = " + table.getTableStyleName());
// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
System.out.println("[4] 構造化参照");
System.out.println(" E1 数式 = " + jsonMap(ws.getRange("E1:E1").getFormula()));
System.out.println(" E1 値 = " + jsonMap(ws.getRange("E1:E1").getValue(true)));
// ---------- 5. コメント（新式）とメモ（旧式）----------
CommentsWrapper cm = ws.getComments();
// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", Arrays.asList(
mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")));
// 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", new FontObjectWrapper());
List<CommentObjectWrapper> thread = cm.getComment("B2");
System.out.println("[5] コメント");
System.out.println(" B2 スレッド件数 = " + thread.size());
for (int i = 0; i < thread.size(); i++) {
CommentObjectWrapper c = thread.get(i);
String pid = c.getParentId();
boolean isReply = pid != null && !pid.isEmpty();
System.out.println(" [" + i + "] " + c.getAuthor() + ": " + c.getContent()
+ (isReply ? " (返信)" : ""));
}
System.out.println(" D2 コメント件数 = " + cm.getComment("D2").size());
System.out.println(" A2 メモ = \"" + cm.getMemoText("A2") + "\" / 作者 \"" + cm.getMemoAuthor("A2") + "\"");
// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
cm.removeComment("D2");
cm.removeMemo("A2");
try {
cm.getComment("D2");
System.out.println(" 削除後 D2コメント: 例外なし");
} catch (RuntimeException e) {
System.out.println(" 削除後 D2コメント: 例外 " + e.getMessage());
}
try {
cm.getMemoText("A2");
System.out.println(" 削除後 A2メモ: 例外なし");
} catch (RuntimeException e) {
System.out.println(" 削除後 A2メモ: 例外 " + e.getMessage());
}
System.out.println(" B2 は残すので save 後も " + cm.getComment("B2").size() + "件");
String total = ws.getRange("B5:B5").getValue(true).get("B5");
String e1 = ws.getRange("E1:E1").getValue(true).get("E1");
System.out.println("[確認] テーブル=" + table.getName() + " 集計(単価)=" + total + " 構造化参照(E1)=" + e1);
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// tables-comments.md — テーブル(ListObject)作成・集計行・構造化参照・セルコメント (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("tables-comments-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// JSON.stringify 相当の出力（非 ASCII はそのまま）
$jq = fn($x): string => json\_encode($x, JSON\_UNESCAPED\_UNICODE | JSON\_UNESCAPED\_SLASHES);
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
echo " 名前 = ", $table->getName(), "\n";
echo " 列名 = ", $jq($table->getColumnName()), "\n";
echo " スタイル = ", $table->getTableStyleName(), "\n";
echo " 全範囲 = ", $table->getAllRange()->getAddress(), "\n";
echo " ヘッダ = ", $table->getHeaderRowRange()->getAddress(), "\n";
echo " データ部 = ", $table->getDataBodyRange()->getAddress(), "\n";
echo " 集計行 = ", $table->getTotalRowRange()->getAddress(), "\n";
// ---------- 2. 集計行 ----------
$table->setTotalRowLabel(1, "合計");
$table->setTotalRowFunction(2, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
$table->setTotalRowFunction(3, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
echo "[2] 集計行\n";
echo " ラベル(1列目) = ", $jq($table->getTotalRowLabel(1)), "\n";
echo " 数式 = ", $jq($ws->getRange("B5:C5")->getFormula()), "\n";
echo " 値 = ", $jq($ws->getRange("B5:C5")->getValue(true)), "\n";
// ---------- 3. テーブルスタイルを変更 ----------
$table->setBuiltinStyleName(\phposbxl\enums\XlDefaultTableStyle::TableStyleMedium9, true);
echo "[3] スタイル変更後 = ", $table->getTableStyleName(), "\n";
// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
$ws->getRange("E1:E1")->setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
echo "[4] 構造化参照\n";
echo " E1 数式 = ", $jq($ws->getRange("E1:E1")->getFormula()), "\n";
echo " E1 値 = ", $jq($ws->getRange("E1:E1")->getValue(true)), "\n";
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
echo " B2 スレッド件数 = ", count($thread), "\n";
foreach ($thread as $i => $c) {
echo " [{$i}] ", $c->getAuthor(), ": ", $c->getContent(), $c->getParentId() ? " (返信)" : "", "\n";
}
echo " D2 コメント件数 = ", count($cm->getComment("D2")), "\n";
echo " A2 メモ = ", $jq($cm->getMemoText("A2")), " / 作者 ", $jq($cm->getMemoAuthor("A2")), "\n";
// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
$cm->removeComment("D2");
$cm->removeMemo("A2");
$tryRead = function (string $label, callable $fn): void {
try {
$fn();
echo " 削除後 {$label}: 例外なし\n";
} catch (\Exception $e) {
echo " 削除後 {$label}: 例外 ", $e->getMessage(), "\n";
}
};
$tryRead("D2コメント", fn() => $cm->getComment("D2"));
$tryRead("A2メモ", fn() => $cm->getMemoText("A2"));
echo " B2 は残すので save 後も ", count($cm->getComment("B2")), "件\n";
$total = $ws->getRange("B5:B5")->getValue(true)["B5"];
$e1 = $ws->getRange("E1:E1")->getValue(true)["E1"];
echo "[確認] テーブル=", $table->getName(), " 集計(単価)=", $total, " 構造化参照(E1)=", $e1, "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// tables-comments.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("tables-comments-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- データ投入（ヘッダ1行 + データ3行）----------
const rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", "120", "3"],
    ["ミカン", "80", "5"],
    ["ブドウ", "300", "2"],
];
ws.setRangeData("A1", rows);
// ---------- 1. テーブルを作成 ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] テーブル作成");
console.log(` 名前 = ${table.getName()}`);
console.log(` 列名 = ${JSON.stringify(table.getColumnName())}`);
console.log(` スタイル = ${table.getTableStyleName()}`);
console.log(` 全範囲 = ${table.getAllRange().getAddress()}`);
console.log(` ヘッダ = ${table.getHeaderRowRange().getAddress()}`);
console.log(` データ部 = ${table.getDataBodyRange().getAddress()}`);
console.log(` 集計行 = ${table.getTotalRowRange().getAddress()}`);
// ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] 集計行");
console.log(` ラベル(1列目) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(` 数式 = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(` 値 = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);
// ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] スタイル変更後 = ${table.getTableStyleName()}`);
// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
console.log("[4] 構造化参照");
console.log(` E1 数式 = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(` E1 値 = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);
// ---------- 5. コメント（新式）とメモ（旧式）----------
const cm = ws.getComments();
const mkComment = (author, content) => {
const c = new dto.CommentObject();
c.setAuthor(author);
c.setContent(content);
return c;
};
// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", new dto.FontObject());
const thread = cm.getComment("B2");
console.log("[5] コメント");
console.log(` B2 スレッド件数 = ${thread.length}`);
thread.forEach((c, i) => console.log(` [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (返信)" : ""}`));
console.log(` D2 コメント件数 = ${cm.getComment("D2").length}`);
console.log(` A2 メモ = ${JSON.stringify(cm.getMemoText("A2"))} / 作者 ${JSON.stringify(cm.getMemoAuthor("A2"))}`);
// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) => {
try { fn(); console.log(` 削除後 ${label}: 例外なし`); }
catch (e) {
// embind の例外メッセージは "Comments.getComment: " のような前置が付くため、
// node 版と同一の出力にするため前置を除去する
const msg = String(e.message || e).replace(/^[A-Za-z\_$][\w$]\*\.[\w$]+: /, "");
console.log(` 削除後 ${label}: 例外 ${msg}`);
}
};
tryRead("D2コメント", () => cm.getComment("D2"));
tryRead("A2メモ", () => cm.getMemoText("A2"));
console.log(` B2 は残すので save 後も ${cm.getComment("B2").length}件`);
const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[確認] テーブル=${table.getName()} 集計(単価)=${total} 構造化参照(E1)=${e1}`);
wb.save();
wb.close();
```

```plaintext
[1] テーブル作成
名前 = Sales
列名 = ["商品名","単価","数量"]
スタイル = TableStyleMedium2
全範囲 = A1:C5
ヘッダ = A1:C1
データ部 = A2:C4
集計行 = A5:C5
[2] 集計行
ラベル(1列目) = "合計"
数式 = {"B5":"SUBTOTAL(109,Sales[単価])","C5":"SUBTOTAL(109,Sales[数量])"}
値 = {"B5":"500","C5":"10"}
[3] スタイル変更後 = TableStyleMedium9
[4] 構造化参照
E1 数式 = {"E1":"SUM(Sales[単価])"}
E1 値 = {"E1":"500"}
[5] コメント
B2 スレッド件数 = 2
[0] alice: この数値を確認
[1] bob: 確認しました (返信)
D2 コメント件数 = 1
A2 メモ = "最優先商品です" / 作者 "Carol"
削除後 D2コメント: 例外 comment not found
削除後 A2メモ: 例外 memo not found
B2 は残すので save 後も 2件
[確認] テーブル=Sales 集計(単価)=500 構造化参照(E1)=500
```

## 解説

* **テーブルは「範囲」ではなく「名前」で扱う。** `addList()` の戻り値（`Table`）や `getList("Sales")` から、ヘッダ・データ部・集計行の範囲を `Range` として取り出せます。  
   集計行を付けたので `getAllRange()` は `A1:C5`（データ `A1:C4` + 集計行 5 行目）になります。
* **集計行は `SUBTOTAL` 数式として入る。** `setTotalRowFunction(2, ...Sum)` を設定すると、 B5 には `SUBTOTAL(109,Sales[単価])`（109 = 無視しない合計）が自動生成されます。  
   平均は `SUBTOTAL(101,...)`、件数は `SUBTOTAL(103,...)` に対応します（実測）。  
   任意の数式を入れたいときは `setCustomRowFunction(col, "MAX(C2:C4)")` を使い、 この場合 `getTotalRowFunction()` は `TotalsCalculationCustom`(9) を返します。
* **構造化参照は「テーブル名[列名]」で書く。** `SUM(Sales[単価])` は 500 になります。  
   列名だけ（`SUM(単価)`）では `#NAME?` エラーになります。  
  数式に先頭の `=` は付けません。
* **列名を変えるとテーブル内の構造化参照も追従する。** `setColumnName([...])` で列名を変更すると、 集計行や既存数式の中の `Sales[旧名]` が `Sales[新名]` に自動で置き換わります（実測）。
* **コメントは 2 系統。** スレッド/返信/完了フラグを持つ**新式**（`setComment`・`setCommentThread`・ `getComment`）と、Excel 旧来の吹き出し**メモ**（`setMemo`・`getMemoText`）があります。  
   用途に合わせて使い分けます。  
  どちらも `save()` すればファイルに永続化されます。

## 注意点

* **テーブル名にセル参照と紛らわしい名前は使えません。** `addList("A1", ...)` や `addList("S2", ...)` は `table name is invalid format` 例外になります。  
  同名の再作成は `table name is already exists` 例外です。
* **`addList(insertTotals=true)` の既定集計は全列「なし」（Excel 準拠、2026-10-01 修正）。** 先頭列にラベル `集計` だけが自動で入り、データ列には集計関数が入りません （旧バージョンは末尾列だけ `Count`（`SUBTOTAL(103)`）が既定で入りました）。  
   集計したい列は `setTotalRowFunction()` で明示します。
* **`removeList()` は既定でデータも削除します。** `removeList("Sales")` は `deleteData=true` 扱いで セルの値も消えます。  
  テーブル枠だけ外して値を残すには `removeList("Sales", false)` を使います。
* **新式コメントは `dto.CommentObject` インスタンスが必須。** ドキュメントの `@example` にある `setComment("A1", { text: "..." })` のような**プレーンなリテラルは `Invalid argument` 例外**になります（実測）。  
   `new dto.CommentObject()` → `setAuthor()` / `setContent()` で組み立てて渡します。
* **`setComment` は追加、`setCommentThread` は上書き。** 既にコメントのあるセルに `setComment` すると 件数が増え、新件は自動的に既存スレッドへの返信（`parentId` 付き）になります。  
   スレッドを丸ごと差し替えたいときは `setCommentThread` を使います（実測）。
* **無いコメント/メモを読むと例外になります。** `getComment()` は `comment not found`、 `getMemoText()` / `getMemoAuthor()` は `memo not found` を投げます（2026-10-01 修正）。  
   削除後の確認は try/catch で例外を捕捉して行います。
* **集計行の平均などは浮動小数点のまま返ります。** `getValue(true)` は表示丸め前の値を文字列で返すため、 平均 `SUBTOTAL(101,...)` は `3.3333333333333335` のような値になります（表示書式で丸めてください）。

## 関連

* [目次](index.md)
* 前: [オートフィルターと並べ替え](general-07-autofilter-sort.md) / 次: [条件付き書式](general-09-conditional-format.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
