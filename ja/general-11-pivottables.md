# ピボットテーブル — 作成・集計・集計方法変更・削除

level: advanced ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](/en/general-11-pivottables.md)

## やりたいこと

ワークシート上のソースデータから**ピボットテーブルを作成**し、 行フィールド・値フィールドを配置して集計する。 あわせて集計方法の変更（合計→平均）とピボットテーブルの削除を行う。

| 手順 | メソッド |
| --- | --- |
| ピボットテーブルの追加 | `ws.getPivotTables().addPivotTable(name, sourceRange, destCell)` |
| フィールド取得 | `pt.getField(fieldName)` |
| フィールド配置 | `pt.setFields(settings, layoutType, rowFields, colFields, filterFields, dataFields)` |
| 集計の実行 | `pt.refreshData()` → `pt.refresh()` |
| 集計方法の変更 | `field.setDataFieldSubTotalsMethod(...)` → 再度 `setFields` → `refreshData` → `refresh` |
| ピボットの削除 | `ws.getPivotTables().removePivotTable(name)` |

## 対象クラス

### `PivotTables` (WorkSheet.getPivotTables() で取得)

| メソッド | 説明 |
| --- | --- |
| `addPivotTable(name, sourceRange, destCell)` | ピボットテーブルを追加。sourceRange は A1 形式のセル範囲、destCell は出力先の左上セル (A1 形式)。 |
| `getPivotTable(name)` | 名前でピボットテーブルを取得。存在しない場合は例外。 |
| `removePivotTable(name)` | ピボットテーブルを削除。 |

### `PivotTable`

| メソッド | 説明 |
| --- | --- |
| `getField(fieldName)` | ソースデータのヘッダ名でフィールドを取得 (`dto.PivotFieldObject`)。 |
| `getFields()` | 全フィールドの配列を返す。 |
| `getPivotTableSetting()` | 設定オブジェクト (`dto.PivotTableSettingObject`) を返す。`setFields` の第 1 引数に使う。 |
| `setFields(settings, layout, rows, cols, filters, data)` | フィールドを配置してレイアウトを確定する。 |
| `refreshData()` | ソースデータを再読込して集計し直す。 |
| `refresh()` | フィールド構成の変更を反映して再集計する。 |
| `getLayoutType()` | 現在のレイアウト種別 (`enums.XlPivotTableLayoutType`) を返す。 |

### `dto.PivotFieldObject`

| メソッド | 説明 |
| --- | --- |
| `setFieldType(type)` | フィールドの役割を設定 (`enums.XlPivotFieldType`)。 |
| `setDataFieldSubTotalsMethod(method)` | 集計方法を設定 (`enums.XlPivotFieldSubtotalsMethodType`)。データフィールドのみ有効。 |
| `getCaption()` | フィールド名を返す。 |

### 主な enum

| enum | 値 | 説明 |
| --- | --- | --- |
| `XlPivotFieldType.PivotFieldTypeRow` | 0 | 行フィールド |
| `XlPivotFieldType.PivotFieldTypeCol` | 1 | 列フィールド |
| `XlPivotFieldType.PivotFieldTypeData` | 3 | データ (値) フィールド |
| `XlPivotTableLayoutType.TabularForm` | 2 | テーブル形式 |
| `XlPivotTableLayoutType.CompactForm` | 0 | コンパクト形式 |
| `XlPivotTableLayoutType.OutlineForm` | 1 | アウトライン形式 |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum` | 1 | 合計 |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage` | 3 | 平均 |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeCount` | 2 | 個数 |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMax` | 4 | 最大 |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMin` | 5 | 最小 |

## コード

実行結果:

#### [Node.js]

```javascript
// pivot-sample.js — ピボットテーブル: 作成・集計・集計方法変更・削除
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("pivot.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
const headers = ["地域", "商品", "数量"];
const srcData = [
["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
["北", "りんご", 5], ["西", "りんご", 10], ["東", "みかん", 30],
["北", "りんご", 20],
];
headers.forEach((h, i) => {
const col = String.fromCharCode(65 + i);
ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) => {
row.forEach((v, ci) => {
const col = String.fromCharCode(65 + ci);
const r = ri + 2;
if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
});
});
const srcTotal = srcData.reduce((s, r) => s + r[2], 0);
console.log(`[1] ソースデータ: ${srcData.length} 行, 数量の合計 = ${srcTotal}`);
// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
const pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");
const rf = pt.getField("地域");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("数量");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");
// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
const pivotRows = {};
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
pivotRows[label] = Number(val);
console.log(` ${label} = ${val}`);
}
}
const grandTotal = pivotRows["総計"];
console.log(`[3] 検算: 総計 ${grandTotal} === ソース合計 ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);
// ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] 集計方法を合計→平均に変更");
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
console.log(` ${label} = ${val}`);
}
}
// ---------- 5. フィールド情報の取得 ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} 個 → ${fields.map(f => f.getCaption()).join(", ")}`);
console.log(` getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);
// ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計");
try {
ws.getPivotTables().getPivotTable("売上集計");
console.log("[6] 削除失敗");
} catch (e) {
console.log("[6] removePivotTable → 削除確認OK");
}
wb.save();
wb.close();
```

#### [Python]

```python
# pivot-sample.py — ピボットテーブル: node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
enums = pyosbxl.enums
app = pyosbxl.App()
wb = app.createWorkBook("pivot-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- 1. ソースデータ ----------
# 地域 / 商品 / 数量 の 10 行
headers = ["地域", "商品", "数量"]
src\_data = [
["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
["北", "りんご", 5], ["西", "りんご", 10], ["東", "みかん", 30],
["北", "りんご", 20],
]
for i, h in enumerate(headers):
col = chr(65 + i)
ws.getRange(f"{col}1:{col}1").setValue(h)
for ri, row in enumerate(src\_data):
for ci, v in enumerate(row):
col = chr(65 + ci)
r = ri + 2
cell = ws.getRange(f"{col}{r}:{col}{r}")
if isinstance(v, int):
cell.setNumberValue(v)
else:
cell.setValue(v)
src\_total = sum(row[2] for row in src\_data)
print(f"[1] ソースデータ: {len(src\_data)} 行, 数量の合計 = {src\_total}")
# ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1")
rf = pt.getField("地域")
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow)
df = pt.getField("数量")
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData)
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)")
# ---------- 3. 集計結果の読み戻しと検算 ----------
# ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
pivot\_rows = {}
for r in range(2, 11):
label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
val = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
if label is not None and label != "":
pivot\_rows[label] = float(val)
print(f" {label} = {val}")
grand\_total = pivot\_rows["総計"]
# node の Number 表示（205）に合わせるため、整数なら int で表示する
gt = int(grand\_total) if grand\_total == int(grand\_total) else grand\_total
print(f"[3] 検算: 総計 {gt} === ソース合計 {src\_total} → {'OK' if grand\_total == src\_total else 'NG'}")
# ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[4] 集計方法を合計→平均に変更")
for r in range(2, 11):
label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
val = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
if label is not None and label != "":
print(f" {label} = {val}")
# ---------- 5. フィールド情報の取得 ----------
fields = pt.getFields()
print(f"[5] getFields: {len(fields)} 個 → " + ", ".join(f.getCaption() for f in fields))
print(f" getLayoutType = {int(pt.getLayoutType())} (TabularForm={int(enums.XlPivotTableLayoutType.TabularForm)})")
# ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計")
try:
ws.getPivotTables().getPivotTable("売上集計")
print("[6] 削除失敗")
except Exception:
print("[6] removePivotTable → 削除確認OK")
wb.save()
wb.close()
```

#### [Java]

```java
// pivot-sample.java — ピボットテーブル: node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("pivot-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
String[] headers = {"地域", "商品", "数量"};
String[][] srcData = {
{"東", "りんご", "10"}, {"東", "みかん", "20"}, {"西", "りんご", "30"},
{"西", "みかん", "15"}, {"東", "りんご", "25"}, {"北", "みかん", "40"},
{"北", "りんご", "5"}, {"西", "りんご", "10"}, {"東", "みかん", "30"},
{"北", "りんご", "20"},
};
for (int i = 0; i < headers.length; i++) {
String col = String.valueOf((char) ('A' + i));
ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
}
int srcTotal = 0;
for (int ri = 0; ri < srcData.length; ri++) {
for (int ci = 0; ci < srcData[ri].length; ci++) {
String col = String.valueOf((char) ('A' + ci));
int r = ri + 2;
String addr = col + r + ":" + col + r;
if (ci == 2) { // 数量列は数値として入力
int v = Integer.parseInt(srcData[ri][ci]);
ws.getRange(addr).setNumberValue(v);
srcTotal += v;
} else {
ws.getRange(addr).setValue(srcData[ri][ci]);
}
}
}
System.out.println("[1] ソースデータ: " + srcData.length + " 行, 数量の合計 = " + srcTotal);
// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
PivotTableWrapper pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");
PivotFieldObjectWrapper rf = pt.getField("地域");
rf.setFieldType(XlPivotFieldType.PivotFieldTypeRow);
PivotFieldObjectWrapper df = pt.getField("数量");
df.setFieldType(XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);
List<PivotFieldObjectWrapper> empty = new ArrayList<>();
pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
Arrays.asList(rf), empty, empty, Arrays.asList(df));
pt.refreshData();
pt.refresh();
System.out.println("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");
// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
Map<String, Double> pivotRows = new LinkedHashMap<>();
for (int r = 2; r <= 10; r++) {
String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
String val = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
if (label != null && !label.isEmpty()) {
pivotRows.put(label, Double.parseDouble(val));
System.out.println(" " + label + " = " + val);
}
}
double grandTotal = pivotRows.get("総計");
// node の Number 表示（205）に合わせるため、整数なら long で表示する
String gt = (grandTotal == Math.rint(grandTotal))
? String.valueOf((long) grandTotal) : String.valueOf(grandTotal);
System.out.println("[3] 検算: 総計 " + gt + " === ソース合計 " + srcTotal
+ " → " + (grandTotal == srcTotal ? "OK" : "NG"));
// ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
Arrays.asList(rf), empty, empty, Arrays.asList(df));
pt.refreshData();
pt.refresh();
System.out.println("[4] 集計方法を合計→平均に変更");
for (int r = 2; r <= 10; r++) {
String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
String val = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
if (label != null && !label.isEmpty()) {
System.out.println(" " + label + " = " + val);
}
}
// ---------- 5. フィールド情報の取得 ----------
List<PivotFieldObjectWrapper> fields = pt.getFields();
StringBuilder caps = new StringBuilder();
for (int i = 0; i < fields.size(); i++) {
if (i > 0) caps.append(", ");
caps.append(fields.get(i).getCaption());
}
System.out.println("[5] getFields: " + fields.size() + " 個 → " + caps);
System.out.println(" getLayoutType = " + pt.getLayoutType().getCode()
+ " (TabularForm=" + XlPivotTableLayoutType.TabularForm.getCode() + ")");
// ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計");
try {
ws.getPivotTables().getPivotTable("売上集計");
System.out.println("[6] 削除失敗");
} catch (Throwable t) {
System.out.println("[6] removePivotTable → 削除確認OK");
}
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// pivot-sample.md — ピボットテーブル: node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("pivot-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
$headers = ["地域", "商品", "数量"];
$srcData = [
["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
["北", "りんご", 5], ["西", "りんご", 10], ["東", "みかん", 30],
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
if (is\_int($v)) $ws->getRange("{$col}{$r}:{$col}{$r}")->setNumberValue($v);
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
$val = $ws->getRange("F{$r}:F{$r}")->getValue()["F{$r}"] ?? null;
if ($label !== null && $label !== "") {
$pivotRows[$label] = (float)$val;
echo " {$label} = {$val}\n";
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
$val = $ws->getRange("F{$r}:F{$r}")->getValue()["F{$r}"] ?? null;
if ($label !== null && $label !== "") {
echo " {$label} = {$val}\n";
}
}
// ---------- 5. フィールド情報の取得 ----------
$fields = $pt->getFields();
$caps = array\_map(fn($f) => $f->getCaption(), $fields);
echo "[5] getFields: " . count($fields) . " 個 → " . implode(", ", $caps) . "\n";
$tabular = \phposbxl\enums\XlPivotTableLayoutType::TabularForm->value;
echo " getLayoutType = {$pt->getLayoutType()} (TabularForm={$tabular})\n";
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
```

#### [WebAssembly]

```javascript
// pivot-sample.wasm.js — ピボットテーブル: node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("pivot-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
const headers = ["地域", "商品", "数量"];
const srcData = [
["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
["北", "りんご", 5], ["西", "りんご", 10], ["東", "みかん", 30],
["北", "りんご", 20],
];
headers.forEach((h, i) => {
const col = String.fromCharCode(65 + i);
ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) => {
row.forEach((v, ci) => {
const col = String.fromCharCode(65 + ci);
const r = ri + 2;
if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
});
});
const srcTotal = srcData.reduce((s, r) => s + r[2], 0);
console.log(`[1] ソースデータ: ${srcData.length} 行, 数量の合計 = ${srcTotal}`);
// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
const pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");
const rf = pt.getField("地域");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("数量");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");
// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
const pivotRows = {};
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
pivotRows[label] = Number(val);
console.log(` ${label} = ${val}`);
}
}
const grandTotal = pivotRows["総計"];
console.log(`[3] 検算: 総計 ${grandTotal} === ソース合計 ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);
// ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] 集計方法を合計→平均に変更");
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
console.log(` ${label} = ${val}`);
}
}
// ---------- 5. フィールド情報の取得 ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} 個 → ${fields.map(f => f.getCaption()).join(", ")}`);
console.log(` getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);
// ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計");
try {
ws.getPivotTables().getPivotTable("売上集計");
console.log("[6] 削除失敗");
} catch (e) {
console.log("[6] removePivotTable → 削除確認OK");
}
wb.save();
wb.close();
```

```plaintext
[1] ソースデータ: 10 行, 数量の合計 = 205
[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)
北 = 65
東 = 85
西 = 55
総計 = 205
[3] 検算: 総計 205 === ソース合計 205 → OK
[4] 集計方法を合計→平均に変更
北 = 21.66666667
東 = 21.25
西 = 18.33333333
総計 = 20.5
[5] getFields: 3 個 → 地域, 商品, 数量
getLayoutType = 2 (TabularForm=2)
[6] removePivotTable → 削除確認OK
```

## 解説

### ピボットテーブル作成の流れ

1. **ソースデータを用意する** — 1 行目にヘッダ（フィールド名）、2 行目以降にデータ。
2. **`addPivotTable(name, sourceRange, destCell)`** — ソース範囲 (A1 形式) と出力先の左上セルを指定。

同じシート内に配置する場合は `destCell` にソースと重ならないセルを指定します。 別シートに配置する場合は `sourceRange` に `Sheet1!A1:C11` のようにシート名を付けます。

1. **フィールドを取得して役割を設定** — `pt.getField("ヘッダ名")` でフィールドを取得し、

`setFieldType()` で 行 / 列 / データ / フィルターのいずれかを割り当てます。

1. **`setFields()` で配置** — 設定オブジェクト、レイアウト種別、各フィールド配列を渡します。

引数の順序は `(settings, layoutType, rowFields, colFields, filterFields, dataFields)` です。

1. **`refreshData()` → `refresh()`** — ソースデータを再読込し、集計結果をシートに書き出します。

**`setFields()` だけでは結果は反映されません。** 必ず `refresh()` を呼んでください。

### 集計方法の変更

データフィールドの集計方法は `setDataFieldSubTotalsMethod()` で変更できます。 変更後は **再度 `setFields()` → `refreshData()` → `refresh()`** が必要です。

| enum 値 | 数値 | 集計方法 |
| --- | --- | --- |
| `SubtotalsMethodTypeSum` | 1 | 合計 (既定) |
| `SubtotalsMethodTypeCount` | 2 | 個数 |
| `SubtotalsMethodTypeAverage` | 3 | 平均 |
| `SubtotalsMethodTypeMax` | 4 | 最大 |
| `SubtotalsMethodTypeMin` | 5 | 最小 |
| `SubtotalsMethodTypeProduct` | 6 | 積 |
| `SubtotalsMethodTypeCountNums` | 7 | 数値の個数 |
| `SubtotalsMethodTypeStdDev` | 8 | 標本標準偏差 |
| `SubtotalsMethodTypeStdDevP` | 9 | 標準偏差 |
| `SubtotalsMethodTypeVar` | 10 | 標本分散 |
| `SubtotalsMethodTypeVarP` | 11 | 分散 |

### 行 × 列のクロス集計

列フィールドも使う場合は、`setFields()` の `colFields` に配列を渡します。

出力は行ラベル (E 列) × 列ラベル (F 列以降) のマトリクスになります。

#### [Node.js]

```javascript
const cf = pt.getField("商品");
cf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeCol);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [cf], [], [df]);
pt.refreshData();
pt.refresh();
```

### 集計結果の読み戻し

ピボットテーブルの出力は通常のセル値として読み取れます。 `ws.getRange("F3:F3").getValue()` のようにして集計値を取得し、 ソースデータの合計と比較することで検算できます。

### ピボットテーブルの削除

`removePivotTable(name)` で削除します。削除後に `getPivotTable(name)` を呼ぶと例外が投げられます。

## 注意点

* **`setFields()` の後に必ず `refresh()` を呼ぶ。** `setFields()` は配置を登録するだけで、シートへの書き出しは `refresh()` が行います。  
   ソースデータを変更した場合は `refreshData()` を先に呼びます。
* **列フィールドなしのピボット (行 + 値のみ) も動作します。** `colFields` に空配列 `[]` を渡します。
* **`getFields()` はソースデータの全ヘッダを返します。** `setFields()` で配置しなかったフィールドも含まれるため、 配置済みフィールドだけを取得したい場合は `getFieldType()` で確認してください。
* **`addPivotTable()` のソース範囲にはヘッダ行を含めます。** `A1:C11` のように 1 行目をヘッダとした範囲を指定します。  
   ヘッダ名がそのままフィールド名になります。
* **集計結果の「総計」行は日本語ロケールで `"総計"` というラベルが付きます。** ラベル文字列は環境に依存する可能性があるため、 読み戻す際は行番号でアクセスするか、ラベルをハードコードせず検索で探すことをおすすめします。
* **同じソース範囲に対して複数のピボットテーブルを作る場合、** 名前 (`pivotTableName`) を変えれば同一シート内に複数配置できます。  
   ただし出力先セルが重ならないようにしてください。
* **`removePivotTable()` で削除しても、ピボットが書き出したセルの値は残ります。** 値も消したい場合は `clearCell()` 等で別途クリアしてください。

## 関連

* [目次](index.md)
* 前: [チャート](general-10-charts.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
