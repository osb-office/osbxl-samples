# オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し

level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](/en/general-07-autofilter-sort.md)

## やりたいこと

データ表に**オートフィルター**を設定して条件に合う行だけを表示し、 あわせて **Sort** で並べ替え（昇順 / 降順）を行う。 フィルターは「条件に合わない行を**非表示にする**」仕組みで、 非表示かどうかは `Row.getHidden()` で読み取れる。

| やり方 | 使うもの |
| --- | --- |
| フィルター設定・条件 | `ws.getAutoFilter(range, firstRowAsHeader)` → `AutoFilter` の各 `set*Filter` |
| 条件の解除 | `af.resetAllFilter()`（条件のみ） / `af.removeFilter()`（フィルター自体を削除） |
| 並べ替え | `ws.getSort(range)` / `af.getSort()` → `executeSortAscending` / `executeSortDescending` |
| 非表示行の判定 | `ws.getRow(r).getHidden()` |

## 対象クラス

### `nodeosbxl.AutoFilter` — `ws.getAutoFilter(A1C1, firstRowAsHeader?)` で取得

| メソッド | 説明 |
| --- | --- |
| `getAddress()` | 対象セル範囲を返します。 |
| `setCustomFilter(op, criteria, colId?)` | 演算子 + 条件式のフィルター。`criteria` に `*` `?` のワイルドカードが使えます。 |
| `setAndCustomFilter(op1, c1, op2, c2, colId?)` / `setOrCustomFilter(...)` | 2 条件の AND / OR 結合。 |
| `setTop10ValueFilter(n, colId?)` / `setBottom10ValueFilter(n, colId?)` | 値の上位 / 下位 n 件。 |
| `setTop10PercentFilter(p, colId?)` / `setBottom10PercentFilter(p, colId?)` | 値の上位 / 下位 p パーセント。 |
| `setAverageFilter(aboveAverage, colId?)` | 平均以上（true）/ 平均以下（false）。 |
| `setUniqueValuesNumberFilter(vals, colId?)` | 数値リストのいずれかと一致する行だけ表示。 |
| `setUniqueValuesStringFilter` / `setUniqueValuesDateTimeFilter` | 文字列 / 日時版（[注意点](#注意点)参照）。 |
| `setDateTimeFilter` / `setDateTimeGroupingFilter` | 日付の期間・グルーピング条件。 |
| `setFontColorFilter` / `setCellColorFilter` / `setIconFilter` | 色・アイコンによる条件。 |
| `getSort()` | ヘッダ行を保護した `Sort` を返します。 |
| `resetFilter(colId)` / `resetAllFilter()` | 条件をクリアします（フィルター自体は残る）。 |
| `removeFilter()` | オートフィルター自体を削除し、非表示行も戻します。 |

`op` は `enums.XlAutoFilterOperator`:

| 値 | 数値 | 意味 |
| --- | --- | --- |
| `Equal` | 1 | 同一 |
| `LessThan` | 2 | 未満 |
| `LessThanOrEqual` | 3 | 以下 |
| `NotEqual` | 4 | 同一でない |
| `GreaterThanOrEqual` | 5 | 以上 |
| `GreaterThan` | 6 | より大きい |

### `nodeosbxl.Sort` — `ws.getSort(A1C1, firstRowAsHeader?)` または `af.getSort()` で取得

| メソッド | 説明 |
| --- | --- |
| `executeSortAscending(target, direction?, matchCase?)` | 昇順に並べ替えます。`target` は範囲内の列番号（先頭列 = 1）。 |
| `executeSortDescending(target, direction?, matchCase?)` | 降順に並べ替えます。 |
| `execute(sortFieldObject)` | `dto.SortFieldObject`（`setSortOnValues(target, ascending)` で条件設定）1 つで並べ替えます。 |
| `executeMultiple([...])` | 複数条件の並べ替え（2026-10-01 修正により正常動作。[注意点](#注意点)参照）。 |
| `getSortConditions()` | 設定済み条件を `dto.SortFieldObject` 配列で読み戻します。 |
| `resetSort(target)` / `resetAllSort()` | 条件をクリアします。**AutoFilter / Table 経由の Sort でのみ有効**。 |

`direction` は `enums.XlRowCol`（`Rows` = 1 / `Columns` = 2）。既定は `Rows`（行を並び替え）。

## コード

実行結果:

#### [Node.js]

```javascript
// af-sort.js — オートフィルターと並べ替え
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("af-sort.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. データ準備（ヘッダ + 10 行）----------
const DATA = [
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
["商品名", "地域", "数量", "単価"].forEach((h, i) =>
ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) => {
const r = i + 2;
ws.getRange(`A${r}:A${r}`).setValue(row[0]);
ws.getRange(`B${r}:B${r}`).setValue(row[1]);
ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] データ準備: A1:D11（ヘッダ + 10 行）");
// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
const text = (r, col) => ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) => `${text(r, "A")}(${text(r, col)})`;
const visible = (col) => {
const out = [];
for (let r = 2; r <= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
return out.join(" ");
};
const hiddenList = () => {
const out = [];
for (let r = 2; r <= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
return out.length ? out.join(",") : "なし";
};
// ---------- 2. オートフィルター設定 ----------
const af = ws.getAutoFilter("A1:D11", true); // 先頭行はヘッダ
console.log(`[2] getAutoFilter → 範囲 = ${af.getAddress()} / 非表示行 = ${hiddenList()}`);
// ---------- 3. 条件フィルター: 数量 > 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] 数量>100 の可視行: ${visible("C")}`);
console.log(` 非表示行 = ${hiddenList()} / getValue("A2:A11") の件数 = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter(); // 条件だけ解除（フィルター自体は残る）
// ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3); // 数量の上位 3 件
console.log(`[4] 数量上位3件の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3); // 数量が 120 または 40
console.log(`[5] 数量∈{120,40} の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["東京", "福岡"], 2); // 地域が 東京 または 福岡（文字列リスト）
console.log(`[6] 地域∈{東京,福岡} の可視行: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "大阪", 2); // 地域 = 大阪（文字列一致）
console.log(`[7] 地域="大阪" の可視行: ${visible("B")}`);
af.resetAllFilter();
// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
ws.getSort("A2:D11").executeSortAscending(3); // 数量（範囲内3列目）で昇順
console.log(`[8] 数量の昇順: ${visible("C")}`);
// オートフィルター経由の Sort はヘッダ行が自動的に保護される
const sort = af.getSort();
sort.executeSortDescending(4); // 単価（4列目）で降順
console.log(`[9] 単価の降順: ${visible("D")}`);
// ---------- 6. 並び替え条件の読み戻し ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] 条件の読み戻し: 対象列 = ${cond.getTarget()} / 昇順 = ${cond.getSortAscending()}`);
// ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter();
console.log(`[11] removeFilter 後: 非表示行 = ${hiddenList()}`);
wb.save();
wb.close();
```

#### [Python]

```python
# af-sort.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("af-sort-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- 1. データ準備（ヘッダ + 10 行）----------
DATA = [
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
]
for i, h in enumerate(["商品名", "地域", "数量", "単価"]):
col = "ABCD"[i]
ws.getRange(f"{col}1:{col}1").setValue(h)
for i, row in enumerate(DATA):
r = i + 2
ws.getRange(f"A{r}:A{r}").setValue(row[0])
ws.getRange(f"B{r}:B{r}").setValue(row[1])
ws.getRange(f"C{r}:C{r}").setNumberValue(row[2])
ws.getRange(f"D{r}:D{r}").setNumberValue(row[3])
print("[1] データ準備: A1:D11（ヘッダ + 10 行）")
# 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
def text(r, col):
return ws.getRange(f"{col}{r}:{col}{r}").getValue()[f"{col}{r}"]
def label(r, col):
return f"{text(r, 'A')}({text(r, col)})"
def visible(col):
out = []
for r in range(2, 12):
if not ws.getRow(r).getHidden():
out.append(label(r, col))
return " ".join(out)
def hidden\_list():
out = [str(r) for r in range(2, 12) if ws.getRow(r).getHidden()]
return ",".join(out) if out else "なし"
# ---------- 2. オートフィルター設定 ----------
af = ws.getAutoFilter("A1:D11", True) # 先頭行はヘッダ
print(f"[2] getAutoFilter → 範囲 = {af.getAddress()} / 非表示行 = {hidden\_list()}")
# ---------- 3. 条件フィルター: 数量 > 100 ----------
# colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.GreaterThan, "100", 3)
print(f"[3] 数量>100 の可視行: {visible('C')}")
print(f' 非表示行 = {hidden\_list()} / getValue("A2:A11") の件数 = {len(ws.getRange("A2:A11").getValue())}')
af.resetAllFilter() # 条件だけ解除（フィルター自体は残る）
# ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3) # 数量の上位 3 件
print(f"[4] 数量上位3件の可視行: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesNumberFilter([120, 40], 3) # 数量が 120 または 40
print(f"[5] 数量∈{{120,40}} の可視行: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesStringFilter(["東京", "福岡"], 2) # 地域が 東京 または 福岡（文字列リスト）
print(f"[6] 地域∈{{東京,福岡}} の可視行: {visible('B')}")
af.resetAllFilter()
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.Equal, "大阪", 2) # 地域 = 大阪（文字列一致）
print(f'[7] 地域="大阪" の可視行: {visible("B")}')
af.resetAllFilter()
# ---------- 5. 並べ替え（単一キー）----------
# ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = False）
ws.getSort("A2:D11").executeSortAscending(3) # 数量（範囲内3列目）で昇順
print(f"[8] 数量の昇順: {visible('C')}")
# オートフィルター経由の Sort はヘッダ行が自動的に保護される
sort = af.getSort()
sort.executeSortDescending(4) # 単価（4列目）で降順
print(f"[9] 単価の降順: {visible('D')}")
# ---------- 6. 並び替え条件の読み戻し ----------
cond = sort.getSortConditions()[0]
# Python の bool は True/False 表示なので lower() で JS 風に揃える
print(f"[10] 条件の読み戻し: 対象列 = {cond.getTarget()} / 昇順 = {str(cond.getSortAscending()).lower()}")
# ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter()
print(f"[11] removeFilter 後: 非表示行 = {hidden\_list()}")
wb.save()
wb.close()
```

#### [Java]

```java
// af-sort.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
public class Main {
static WorkSheetWrapper ws;
// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
static String text(int r, String col) {
return ws.getRange(col + r + ":" + col + r).getValue().get(col + r);
}
static String label(int r, String col) {
return text(r, "A") + "(" + text(r, col) + ")";
}
static String visible(String col) {
List<String> out = new ArrayList<>();
for (int r = 2; r <= 11; r++) {
if (!ws.getRow(r).getHidden()) out.add(label(r, col));
}
return String.join(" ", out);
}
static String hiddenList() {
List<String> out = new ArrayList<>();
for (int r = 2; r <= 11; r++) {
if (ws.getRow(r).getHidden()) out.add(String.valueOf(r));
}
return out.isEmpty() ? "なし" : String.join(",", out);
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("af-sort-java.xlsx");
ws = wb.openWorkSheet("Sheet1");
// ---------- 1. データ準備（ヘッダ + 10 行）----------
Object[][] DATA = {
{"りんご", "東京", 120.0, 100.0},
{"みかん", "大阪", 40.0, 80.0},
{"バナナ", "東京", 200.0, 150.0},
{"ぶどう", "福岡", 60.0, 300.0},
{"なし", "大阪", 90.0, 120.0},
{"もも", "福岡", 150.0, 400.0},
{"キウイ", "東京", 30.0, 90.0},
{"マンゴー", "大阪", 70.0, 500.0},
{"いちご", "福岡", 110.0, 350.0},
{"メロン", "東京", 25.0, 800.0},
};
String[] headers = {"商品名", "地域", "数量", "単価"};
for (int i = 0; i < headers.length; i++) {
String col = String.valueOf("ABCD".charAt(i));
ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
}
for (int i = 0; i < DATA.length; i++) {
int r = i + 2;
ws.getRange("A" + r + ":A" + r).setValue((String) DATA[i][0]);
ws.getRange("B" + r + ":B" + r).setValue((String) DATA[i][1]);
ws.getRange("C" + r + ":C" + r).setNumberValue((Double) DATA[i][2]);
ws.getRange("D" + r + ":D" + r).setNumberValue((Double) DATA[i][3]);
}
System.out.println("[1] データ準備: A1:D11（ヘッダ + 10 行）");
// ---------- 2. オートフィルター設定 ----------
AutoFilterWrapper af = ws.getAutoFilter("A1:D11", true); // 先頭行はヘッダ
System.out.println("[2] getAutoFilter → 範囲 = " + af.getAddress() + " / 非表示行 = " + hiddenList());
// ---------- 3. 条件フィルター: 数量 > 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(XlAutoFilterOperator.GreaterThan, "100", 3);
System.out.println("[3] 数量>100 の可視行: " + visible("C"));
System.out.println(" 非表示行 = " + hiddenList()
+ " / getValue(\"A2:A11\") の件数 = " + ws.getRange("A2:A11").getValue().size());
af.resetAllFilter(); // 条件だけ解除（フィルター自体は残る）
// ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3); // 数量の上位 3 件
System.out.println("[4] 数量上位3件の可視行: " + visible("C"));
af.resetAllFilter();
af.setUniqueValuesNumberFilter(Arrays.asList(120.0, 40.0), 3); // 数量が 120 または 40
System.out.println("[5] 数量∈{120,40} の可視行: " + visible("C"));
af.resetAllFilter();
af.setUniqueValuesStringFilter(Arrays.asList("東京", "福岡"), 2); // 地域が 東京 または 福岡（文字列リスト）
System.out.println("[6] 地域∈{東京,福岡} の可視行: " + visible("B"));
af.resetAllFilter();
af.setCustomFilter(XlAutoFilterOperator.Equal, "大阪", 2); // 地域 = 大阪（文字列一致）
System.out.println("[7] 地域=\"大阪\" の可視行: " + visible("B"));
af.resetAllFilter();
// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
ws.getSort("A2:D11").executeSortAscending(3); // 数量（範囲内3列目）で昇順
System.out.println("[8] 数量の昇順: " + visible("C"));
// オートフィルター経由の Sort はヘッダ行が自動的に保護される
SortWrapper sort = af.getSort();
sort.executeSortDescending(4); // 単価（4列目）で降順
System.out.println("[9] 単価の降順: " + visible("D"));
// ---------- 6. 並び替え条件の読み戻し ----------
SortFieldObjectWrapper cond = sort.getSortConditions().get(0);
System.out.println("[10] 条件の読み戻し: 対象列 = " + cond.getTarget()
+ " / 昇順 = " + cond.getSortAscending());
// ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter();
System.out.println("[11] removeFilter 後: 非表示行 = " + hiddenList());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// af-sort.md — オートフィルターと並べ替え (phposbxl)
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
$af = $ws->getAutoFilter("A1:D11", true); // 先頭行はヘッダ
echo "[2] getAutoFilter → 範囲 = ", $af->getAddress(), " / 非表示行 = ", $hiddenList(), "\n";
// ---------- 3. 条件フィルター: 数量 > 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
$af->setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::GreaterThan, "100", 3);
echo "[3] 数量>100 の可視行: ", $visible("C"), "\n";
echo " 非表示行 = ", $hiddenList(),
" / getValue(\"A2:A11\") の件数 = ", count($ws->getRange("A2:A11")->getValue()), "\n";
$af->resetAllFilter(); // 条件だけ解除（フィルター自体は残る）
// ---------- 4. Top-N / 値リストのフィルター ----------
$af->setTop10ValueFilter(3, 3); // 数量の上位 3 件
echo "[4] 数量上位3件の可視行: ", $visible("C"), "\n";
$af->resetAllFilter();
$af->setUniqueValuesNumberFilter([120, 40], 3); // 数量が 120 または 40
echo "[5] 数量∈{120,40} の可視行: ", $visible("C"), "\n";
$af->resetAllFilter();
$af->setUniqueValuesStringFilter(["東京", "福岡"], 2); // 地域が 東京 または 福岡（文字列リスト）
echo "[6] 地域∈{東京,福岡} の可視行: ", $visible("B"), "\n";
$af->resetAllFilter();
$af->setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::Equal, "大阪", 2); // 地域 = 大阪（文字列一致）
echo "[7] 地域=\"大阪\" の可視行: ", $visible("B"), "\n";
$af->resetAllFilter();
// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
$ws->getSort("A2:D11")->executeSortAscending(3); // 数量（範囲内3列目）で昇順
echo "[8] 数量の昇順: ", $visible("C"), "\n";
// オートフィルター経由の Sort はヘッダ行が自動的に保護される
$sort = $af->getSort();
$sort->executeSortDescending(4); // 単価（4列目）で降順
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
```

#### [WebAssembly]

```javascript
// af-sort.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("af-sort-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. データ準備（ヘッダ + 10 行）----------
const DATA = [
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
["商品名", "地域", "数量", "単価"].forEach((h, i) =>
ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) => {
const r = i + 2;
ws.getRange(`A${r}:A${r}`).setValue(row[0]);
ws.getRange(`B${r}:B${r}`).setValue(row[1]);
ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] データ準備: A1:D11（ヘッダ + 10 行）");
// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
const text = (r, col) => ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) => `${text(r, "A")}(${text(r, col)})`;
const visible = (col) => {
const out = [];
for (let r = 2; r <= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
return out.join(" ");
};
const hiddenList = () => {
const out = [];
for (let r = 2; r <= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
return out.length ? out.join(",") : "なし";
};
// ---------- 2. オートフィルター設定 ----------
const af = ws.getAutoFilter("A1:D11", true); // 先頭行はヘッダ
console.log(`[2] getAutoFilter → 範囲 = ${af.getAddress()} / 非表示行 = ${hiddenList()}`);
// ---------- 3. 条件フィルター: 数量 > 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] 数量>100 の可視行: ${visible("C")}`);
console.log(` 非表示行 = ${hiddenList()} / getValue("A2:A11") の件数 = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter(); // 条件だけ解除（フィルター自体は残る）
// ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3); // 数量の上位 3 件
console.log(`[4] 数量上位3件の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3); // 数量が 120 または 40
console.log(`[5] 数量∈{120,40} の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["東京", "福岡"], 2); // 地域が 東京 または 福岡（文字列リスト）
console.log(`[6] 地域∈{東京,福岡} の可視行: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "大阪", 2); // 地域 = 大阪（文字列一致）
console.log(`[7] 地域="大阪" の可視行: ${visible("B")}`);
af.resetAllFilter();
// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
ws.getSort("A2:D11").executeSortAscending(3); // 数量（範囲内3列目）で昇順
console.log(`[8] 数量の昇順: ${visible("C")}`);
// オートフィルター経由の Sort はヘッダ行が自動的に保護される
const sort = af.getSort();
sort.executeSortDescending(4); // 単価（4列目）で降順
console.log(`[9] 単価の降順: ${visible("D")}`);
// ---------- 6. 並び替え条件の読み戻し ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] 条件の読み戻し: 対象列 = ${cond.getTarget()} / 昇順 = ${cond.getSortAscending()}`);
// ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter();
console.log(`[11] removeFilter 後: 非表示行 = ${hiddenList()}`);
wb.save();
wb.close();
```

```plaintext
[1] データ準備: A1:D11（ヘッダ + 10 行）
[2] getAutoFilter → 範囲 = A1:D11 / 非表示行 = なし
[3] 数量>100 の可視行: りんご(120) バナナ(200) もも(150) いちご(110)
非表示行 = 3,5,6,8,9,11 / getValue("A2:A11") の件数 = 10
[4] 数量上位3件の可視行: りんご(120) バナナ(200) もも(150)
[5] 数量∈{120,40} の可視行: りんご(120) みかん(40)
[6] 地域∈{東京,福岡} の可視行: りんご(東京) バナナ(東京) ぶどう(福岡) もも(福岡) キウイ(東京) いちご(福岡) メロン(東京)
[7] 地域="大阪" の可視行: みかん(大阪) なし(大阪) マンゴー(大阪)
[8] 数量の昇順: メロン(25) キウイ(30) みかん(40) ぶどう(60) マンゴー(70) なし(90) いちご(110) りんご(120) もも(150) バナナ(200)
[9] 単価の降順: メロン(800) マンゴー(500) もも(400) いちご(350) ぶどう(300) バナナ(150) なし(120) りんご(100) キウイ(90) みかん(80)
[10] 条件の読み戻し: 対象列 = 4 / 昇順 = false
[11] removeFilter 後: 非表示行 = なし
```

## 解説

### フィルター = 行の非表示

オートフィルターは**条件に合わない行を非表示にする**だけで、データは削除されません。

* 非表示判定は `ws.getRow(r).getHidden()`（`Row` の使い方は[行・列の操作](general-05-rows-cols.md)参照）。
* **`getValue()` は非表示行の値もそのまま返します**（実行結果 `[3]`: 6 行が非表示でも件数は 10）。  
   「見えている行だけ」が欲しい場合は、`getHidden()` で絞り込んでから読んでください。
* `resetFilter(colId)` / `resetAllFilter()` は**条件だけ**をクリアし、フィルター（枠）は残します。
* `removeFilter()` はフィルター自体を削除し、非表示行もすべて表示に戻します（実行結果 `[11]`）。
* フィルター範囲・条件・非表示状態は **save → 再オープン後も復元**されます（実測）。

### 条件の書き方

* `setCustomFilter(op, criteria, colId)` の `colId` と、`Sort` の `target` は **どちらも「指定範囲内の列番号」（先頭列 = 1）** です。  
  シートの列番号そのものではないので、 範囲が `B` 列始まりのときはずれに注意してください。
* `criteria` は**値だけ**を書きます（`"100"`）。  
  演算子は `op`（`enums.XlAutoFilterOperator`）で指定します。
* 数値の大小比較（`GreaterThan` など）、Top-N、平均、数値リスト一致（`setUniqueValuesNumberFilter`）は 期待どおり動作します（実行結果 `[3]`〜`[5]`）。

### 並べ替え

* `ws.getSort("A2:D11")` のように**データ部だけを範囲にする**のが基本形です。  
   ヘッダ込みの範囲を渡すなら `ws.getSort("A1:D11", true)`（第2引数 `firstRowAsHeader`）か、 **`af.getSort()`（ヘッダが自動保護される）** を使います。
* 並べ替えは**行単位**で動きます。キー列以外の列（商品名・地域など）も行と一緒に移動します。
* 実行後は `getSortConditions()` で条件を読み戻せます（`getTarget()` / `getSortAscending()` / `getDirection()` / `getMatchCase()`）。  
  読み戻した条件は save → 再オープン後も残ります。
* `executeSortAscending(target, direction, matchCase)` の `direction` は `enums.XlRowCol.Rows`(1) / `Columns`(2)。  
  既定は `Rows` です。

## 注意点

* **`criteria` には値のみを渡します（演算子記号・ワイルドカードを付けない）。** 演算子は `enums.XlAutoFilterOperator` で指定する設計です。  
   `setCustomFilter(GreaterThan, ">100", 3)` のように記号を混ぜると条件が文字列扱いになり、 数値列では 1 件もマッチせず全行が非表示になります（`"100"` なら正常）。
* **文字列の一致系フィルターは正常に動作します**（2026-10-01 修正。  
  旧バージョン 1.5.0 では セッション内に新規作成した文字列セルが一致せず全行非表示になるバグがありました）。  
   `setCustomFilter(Equal, "東京", ...)` / `setUniqueValuesStringFilter(["東京","福岡"], ...)` / `setOrCustomFilter(...)` いずれも実行結果 `[6]` `[7]` のとおり動作します。
* **`executeMultiple()`（複数キーの並べ替え）は正常に動作します**（2026-10-01 修正）。  
   旧バージョン 1.5.0 では行の重複・欠落が起きていました。  
   `SortFieldObject.setSortOnValues(列, 昇順?)` で条件を作り、優先順に配列で渡します。  
   受け入れテスト: `test/sort_af_fc_ext.test.js`「Sort executeMultiple (複数キー・A-2 修正)」。
* **`executeSortAscending()` を複数回呼んでも複数キーソートにはなりません。** 後から呼んだキーだけで全体が並べ替わり、`getSortConditions()` には条件が蓄積されていきます（実測）。
* **ヘッダ込み範囲を `firstRowAsHeader` なし（既定 false）でソートするとヘッダ行もデータとして並びます。** 実測で、昇順ソートすると 1 行目のヘッダ（文字列）が最下部へ沈みました。  
   ヘッダを含む範囲では `getSort(range, true)` または `af.getSort()` を使ってください。
* **フィルター適用中にソートすると、可視行だけが並べ替わります。** 非表示行は元の位置に留まるため、`resetAllFilter()` してからソートするのが安全です（実測）。
* **`resetSort()` / `resetAllSort()` が効くのは AutoFilter / Table 経由の Sort だけです。** `ws.getSort()` で取得した Sort では条件がクリアされないことを実測で確認しました（ドキュメントどおりの仕様）。
* **`getAutoFilter(A1C1, firstRowAsHeader?)` は「取得」であると同時に範囲の設定でもあります。** オートフィルターはシート単位で 1 範囲だけ。  
  別の範囲で呼び出すと設定範囲が置き換わります。  
   第2引数は「取得するかしないか」ではなく `firstRowAsHeader`（先頭行をヘッダとみなすか）です。
* **数値列は `setNumberValue()`（または数値として解釈される値）で入れてください。** `setValue("120")` のような文字列として入れた数値でも大小比較フィルターは動作することを実測しましたが、 Top-N 系などの安定動作のためには数値型での入力を推奨します。

## 関連

* [目次](index.md)
* 前: [大量データの一括投入](general-06-bulk-input.md) / 次: [テーブルとコメント](general-08-tables-comments.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
