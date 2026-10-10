# ブックとワークシートの操作 — 追加・削除・移動・リネーム

level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-03-sheets.md)

## やりたいこと

1 つのブックの中でワークシートを追加・削除・移動・リネームし、 インデックスでシートを辿ったり、開いたときに表示されるシート（アクティブタブ）を制御する。

## 対象メソッド

| メソッド | 所属 | 説明 |
| --- | --- | --- |
| `addWorkSheet(sheetName, position?)` | WorkBook | ワークシートを追加します。 |
| `deleteWorkSheet(sheetName)` | WorkBook | ワークシートを削除します。 |
| `moveSheet(sheetName, position)` | WorkBook | ワークシートを移動します。 |
| `openWorkSheet(sheetName)` | WorkBook | ワークシートを名前で開きます。 |
| `openWorkSheetByIndex(sheetIndex)` | WorkBook | ワークシートをインデックスで開きます。 |
| `getSheetCount()` | WorkBook | ワークシート数を取得します。 |
| `getSheetNumber(sheetName)` | WorkBook | ワークシートのシート番号を取得します。 |
| `getName()` / `setName(name)` | WorkSheet | シート名称の取得 / 設定を行います。 |
| `isActive()` / `activate(isActive)` | WorkSheet | アクティブタブ（起動時の表示項目）の取得 / 設定を行います。 |

## コード

実行結果:

#### [Node.js]

```javascript
// sheets.js — シートの追加・移動・リネーム・削除と、インデックスアクセス
const { nodeosbxl } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("sheets.xlsx");
// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
const show = (label) => {
const names = [];
for (let i = 1; i <= wb.getSheetCount(); i++) {
names.push(wb.openWorkSheetByIndex(i).getName());
}
console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};
show("[1] 新規作成直後 :");
wb.addWorkSheet("集計"); // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計) :");
wb.addWorkSheet("受注", 1); // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");
wb.moveSheet("受注", 3); // 3 番目へ移動
show("[4] moveSheet(受注,3) :");
const ws = wb.openWorkSheet("受注");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"));
ws.getRange("A1:A1").setValue("受注データ");
wb.openWorkSheet("Sheet1").setName("明細"); // リネーム
show("[6] setName(明細) :");
// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
console.log("[7] isActive: 集計 =", wb.openWorkSheet("集計").isActive(),
"/ 明細 =", wb.openWorkSheet("明細").isActive());
wb.openWorkSheet("集計").activate(true); // 他のシートのアクティブは自動的に解除されます
console.log(" activate後: 集計 =", wb.openWorkSheet("集計").isActive(),
"/ 明細 =", wb.openWorkSheet("明細").isActive());
// --- 削除 ---
wb.addWorkSheet("一時");
wb.deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");
wb.save();
wb.close();
// --- 開き直して確認 ---
const wb2 = app.openWorkBook("sheets.xlsx");
const names = [];
for (let i = 1; i <= wb2.getSheetCount(); i++) names.push(wb2.openWorkSheetByIndex(i).getName());
console.log("[終] シート構成:", names.join(", "));
console.log(" 受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"]);
wb2.close();
```

#### [Python]

```python
# sheets.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("sheets-py.xlsx")
def b(x):
# python の True/False を node の true/false 表記に合わせる
return "true" if x else "false"
# インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
def show(label):
names = [wb.openWorkSheetByIndex(i).getName() for i in range(1, wb.getSheetCount() + 1)]
print(f"{label} count={wb.getSheetCount()} [{', '.join(names)}]")
show("[1] 新規作成直後 :")
wb.addWorkSheet("集計") # position 省略 → 末尾に追加
show("[2] addWorkSheet(集計) :")
wb.addWorkSheet("受注", 1) # position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):")
wb.moveSheet("受注", 3) # 3 番目へ移動
show("[4] moveSheet(受注,3) :")
ws = wb.openWorkSheet("受注")
print("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"))
ws.getRange("A1:A1").setValue("受注データ")
wb.openWorkSheet("Sheet1").setName("明細") # リネーム
show("[6] setName(明細) :")
# --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
print("[7] isActive: 集計 =", b(wb.openWorkSheet("集計").isActive()),
"/ 明細 =", b(wb.openWorkSheet("明細").isActive()))
wb.openWorkSheet("集計").activate(True) # 他のシートのアクティブは自動的に解除されます
print(" activate後: 集計 =", b(wb.openWorkSheet("集計").isActive()),
"/ 明細 =", b(wb.openWorkSheet("明細").isActive()))
# --- 削除 ---
wb.addWorkSheet("一時")
wb.deleteWorkSheet("一時")
show("[8] deleteWorkSheet(一時):")
wb.save()
wb.close()
# --- 開き直して確認 ---
wb2 = app.openWorkBook("sheets-py.xlsx")
names = [wb2.openWorkSheetByIndex(i).getName() for i in range(1, wb2.getSheetCount() + 1)]
print("[終] シート構成:", ", ".join(names))
print(" 受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"])
wb2.close()
```

#### [Java]

```java
// sheets.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
public class Main {
static WorkBookWrapper wb;
// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
static void show(String label) {
List<String> names = new ArrayList<>();
for (int i = 1; i <= wb.getSheetCount(); i++) {
names.add(wb.openWorkSheetByIndex(i).getName());
}
System.out.println(label + " count=" + wb.getSheetCount()
+ " [" + String.join(", ", names) + "]");
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
wb = app.createWorkBook("sheets-java.xlsx");
show("[1] 新規作成直後 :");
wb.addWorkSheet("集計"); // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計) :");
wb.addWorkSheet("受注", 1); // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");
wb.moveSheet("受注", 3); // 3 番目へ移動
show("[4] moveSheet(受注,3) :");
WorkSheetWrapper ws = wb.openWorkSheet("受注");
System.out.println("[5] getName() = " + ws.getName()
+ " / getSheetNumber(受注) = " + wb.getSheetNumber("受注"));
ws.getRange("A1:A1").setValue("受注データ");
wb.openWorkSheet("Sheet1").setName("明細"); // リネーム
show("[6] setName(明細) :");
// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
System.out.println("[7] isActive: 集計 = " + wb.openWorkSheet("集計").isActive()
+ " / 明細 = " + wb.openWorkSheet("明細").isActive());
wb.openWorkSheet("集計").activate(true); // 他のシートのアクティブは自動的に解除されます
System.out.println(" activate後: 集計 = " + wb.openWorkSheet("集計").isActive()
+ " / 明細 = " + wb.openWorkSheet("明細").isActive());
// --- 削除 ---
wb.addWorkSheet("一時");
wb.deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");
wb.save();
wb.close();
// --- 開き直して確認 ---
WorkBookWrapper wb2 = app.openWorkBook("sheets-java.xlsx");
List<String> names = new ArrayList<>();
for (int i = 1; i <= wb2.getSheetCount(); i++) names.add(wb2.openWorkSheetByIndex(i).getName());
System.out.println("[終] シート構成: " + String.join(", ", names));
Map<String, String> v = wb2.openWorkSheet("受注").getRange("A1:A1").getValue();
System.out.println(" 受注!A1 = " + v.get("A1"));
wb2.close();
}
}
```

#### [PHP]

```php
// sheets.md — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sheets-php.xlsx");
// PHP の bool を node の true/false 表記に合わせる
function b($x) {
return $x ? "true" : "false";
}
// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
function show($label) {
global $wb;
$names = [];
for ($i = 1; $i <= $wb->getSheetCount(); $i++) {
$names[] = $wb->openWorkSheetByIndex($i)->getName();
}
echo $label . " count=" . $wb->getSheetCount() . " [" . implode(", ", $names) . "]\n";
}
show("[1] 新規作成直後 :");
$wb->addWorkSheet("集計"); // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計) :");
$wb->addWorkSheet("受注", 1); // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");
$wb->moveSheet("受注", 3); // 3 番目へ移動
show("[4] moveSheet(受注,3) :");
$ws = $wb->openWorkSheet("受注");
echo "[5] getName() = " . $ws->getName() . " / getSheetNumber(受注) = " . $wb->getSheetNumber("受注") . "\n";
$ws->getRange("A1:A1")->setValue("受注データ");
$wb->openWorkSheet("Sheet1")->setName("明細"); // リネーム
show("[6] setName(明細) :");
// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
echo "[7] isActive: 集計 = " . b($wb->openWorkSheet("集計")->isActive())
. " / 明細 = " . b($wb->openWorkSheet("明細")->isActive()) . "\n";
$wb->openWorkSheet("集計")->activate(true); // 他のシートのアクティブは自動的に解除されます
echo " activate後: 集計 = " . b($wb->openWorkSheet("集計")->isActive())
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
for ($i = 1; $i <= $wb2->getSheetCount(); $i++) {
$names[] = $wb2->openWorkSheetByIndex($i)->getName();
}
echo "[終] シート構成: " . implode(", ", $names) . "\n";
$v = $wb2->openWorkSheet("受注")->getRange("A1:A1")->getValue();
echo " 受注!A1 = " . $v["A1"] . "\n";
$wb2->close();
```

#### [WebAssembly]

```javascript
// sheets.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("sheets-wasm.xlsx");
// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
const show = (label) => {
const names = [];
for (let i = 1; i <= wb.getSheetCount(); i++) {
names.push(wb.openWorkSheetByIndex(i).getName());
}
console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};
show("[1] 新規作成直後 :");
wb.addWorkSheet("集計"); // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計) :");
wb.addWorkSheet("受注", 1); // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");
wb.moveSheet("受注", 3); // 3 番目へ移動
show("[4] moveSheet(受注,3) :");
const ws = wb.openWorkSheet("受注");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"));
ws.getRange("A1:A1").setValue("受注データ");
wb.openWorkSheet("Sheet1").setName("明細"); // リネーム
show("[6] setName(明細) :");
// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
console.log("[7] isActive: 集計 =", wb.openWorkSheet("集計").isActive(),
"/ 明細 =", wb.openWorkSheet("明細").isActive());
wb.openWorkSheet("集計").activate(true); // 他のシートのアクティブは自動的に解除されます
console.log(" activate後: 集計 =", wb.openWorkSheet("集計").isActive(),
"/ 明細 =", wb.openWorkSheet("明細").isActive());
// --- 削除 ---
wb.addWorkSheet("一時");
wb.deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");
wb.save();
wb.close();
// --- 開き直して確認 ---
const wb2 = app.openWorkBook("sheets-wasm.xlsx");
const names = [];
for (let i = 1; i <= wb2.getSheetCount(); i++) names.push(wb2.openWorkSheetByIndex(i).getName());
console.log("[終] シート構成:", names.join(", "));
console.log(" 受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"]);
wb2.close();
```

```plaintext
[1] 新規作成直後 : count=1 [Sheet1]
[2] addWorkSheet(集計) : count=2 [Sheet1, 集計]
[3] addWorkSheet(受注,1): count=3 [受注, Sheet1, 集計]
[4] moveSheet(受注,3) : count=3 [Sheet1, 集計, 受注]
[5] getName() = 受注 / getSheetNumber(受注) = 3
[6] setName(明細) : count=3 [明細, 集計, 受注]
[7] isActive: 集計 = false / 明細 = true
activate後: 集計 = true / 明細 = false
[8] deleteWorkSheet(一時): count=3 [明細, 集計, 受注]
[終] シート構成: 明細, 集計, 受注
受注!A1 = 受注データ
```

## 解説

* **`position` とシート番号はどちらも 1 始まり。** `addWorkSheet(name, position)` は **`position` の前に挿入**します（`position` = 1 なら先頭）。  
  省略（既定値 -1）で末尾に追加されます。
* **`moveSheet(name, position)` の `position` は「移動後の位置」**で、先頭が 1 です。
* **`openWorkSheetByIndex(i)` も 1 始まり。** `getSheetCount()` と組み合わせると 全シートを名前で列挙できます（上記の `show()`）。  
   シート名が分からないブックを扱うときは、この形で辿るのが確実です。
* **`getSheetNumber(name)` は現在の並び順での番号**を返します。  
   `moveSheet()` や `deleteWorkSheet()` の後は値が変わります。
* **`setName()` は即座に反映されます。** 以降は新しい名前で `openWorkSheet()` します。  
   古い名前で開こうとすると例外になります。
* **`activate(true)` は排他的に働きます。** 目的のシートを `activate(true)` にすると、 他のシートのアクティブは**自動的に解除されます**（実測）。  
   あらかじめ他を `activate(false)` にする必要はありません。  
   アクティブ状態は `save()` して開き直しても保持されます。

## 注意点

* **存在しないシート名を `openWorkSheet()` に渡すと例外になります。**

  ```javascript
  wb.openWorkSheet("NoSuch"); // → 例外: "NoSuch is not found"
  ```

  ユーザー入力や外部設定からシート名を受ける場合は、事前に `getSheetCount()` + `openWorkSheetByIndex()` で名前を列挙して照合してください。
* **`openWorkSheetByIndex()` の範囲外インデックスも例外になります。** `1` 〜 `getSheetCount()` の範囲で呼び出してください。
* **新規作成直後は既定シート（`Sheet1`）だけが `isActive() === true`** です（実測）。  
   `addWorkSheet()` で追加したシートは `false` になります。  
   ブックを開いたときに別のシートを手前に見せたいなら、そのシートで `activate(true)` を呼んでください。
* **シートを跨いだセルコピーは `copyRow` / `copyCol` / `copyCell` を使います** （コピー元をシート名で指定）。  
  詳細は [行・列・セル範囲の操作](general-05-rows-cols.md)。  
   図形（グラフ含む）・ピボットテーブル・queryTable 形式のテーブル・データテーブルは コピー対象に含まれません。

## 関連

* [目次](index.md)
* 前: [値と数式を入れて表を作る](general-02-values-formulas.md) / 次: [セル値の型と表示書式](general-04-value-types.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
