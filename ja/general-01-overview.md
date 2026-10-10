# 概要 — オブジェクトモデルと基本の流れ

level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](/en/general-01-overview.md)

## やりたいこと

osbxl の全体像をつかむ。

* オブジェクトの階層（`App` → `WorkBook` → `WorkSheet` → `Range`）
* **基本の流れ**（新規作成 / 既存ブックを開く → シート → Range → 値・数式 → save / saveAs / close）
* 4 クラスの主要メソッド

個々の機能の詳細は、このページの後続サンプルで扱います。

## 前提

各言語のモジュールを[ダウンロードページ](https://osb-office.com/pages/download.php)から入手してインストールします。

PHP の OS 別の詳しい導入手順（Linux / macOS / Windows、Composer が未インストールの場合を含む）は[クイックスタート](https://osb-office.com/pages/quickstart.php)を参照してください。

WebAssembly 版は MEMORY64（Wasm 3.0）ビルドのため、実行環境側の対応が必要です （Chrome/Edge 133〜・Firefox 134〜・Node.js 24〜。Safari は現在未対応）。

入口は言語ごとに異なります。namespace 構成（操作対象クラス / dto / chart / enums）は4言語共通です。

#### [Node.js]

```bash
npm install ./nodeosbxl-1.5.0.tgz
```

#### [Python]

```bash
# (例) Python 3.12 以上（abi3 wheel 1 本で 3.12 以降をカバー）
pip install ./pyosbxl-1.5.0-cp312-abi3-linux\_x86\_64.whl
```

#### [Java]

```bash
# JAR をクラスパスに追加（javacpp-1.5.14.jar も必要）
javac -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar Main.java
java -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar:. Main
```

#### [PHP]

```bash
# HP (osb-office.com/files/php/) から PHP バージョン(8.2〜8.5)と OS に合う
# tarball または composer zip をダウンロードする
# 例: phposbxl-1.5.0-php8.5-nts-linux-x86\_64.tar.gz (Linux・PHP 8.5 の場合)
# 展開して composer の path リポジトリに指定し導入する
composer config repositories.mdosbxl path /path/to/phposbxl-1.5.0-php8.5-nts-linux-x86\_64
composer require osb-office/phposbxl
# composer zip の場合: zip を置いたディレクトリを artifact リポジトリに指定
# composer config repositories.mdosbxl artifact /path/to/zip-dir && composer require osb-office/phposbxl
```

#### [WebAssembly]

```bash
# パッケージを展開（osbxl\_wasm.js / osbxl\_em.js / license/ を含む）
tar xzf wasmosbxl-1.5.0.tgz
```
```javascript
// 読み込み例（Node.js 24〜 必須: この wasm は MEMORY64 = Wasm 3.0 ビルドです）
const fs = require("fs");
const createModule = require("./package/osbxl\_wasm.js"); // emscripten のファクトリ
const { createOsbxl } = require("./package/osbxl\_em.js"); // ローダー＋enum 定義
createOsbxl(createModule).then(({ Module, osbxl, dto, chart, enums }) => {
// ライセンスファイルを wasm 仮想 FS へ配置（new App() より前に必須）
const lic = process.env.OSB\_LICENSE\_PATH\_WASM; // 例: /root/lic\_node.osb
if (lic && fs.existsSync(lic)) {
Module.FS.mkdirTree(lic.substring(0, lic.lastIndexOf("/")));
Module.FS.writeFile(lic, fs.readFileSync(lic));
}
const app = new osbxl.App(); // ここから先は Node.js 版と同じ（namespace 名のみ osbxl）
// ...
});
// ブラウザでは <script> で osbxl\_wasm.js を読み込んで同様に createOsbxl を呼びます。
// 対応環境: Chrome/Edge 133〜・Firefox 134〜（Safari は現在未対応）
```

| 言語 | 入口 | 主な namespace |
| --- | --- | --- |
| Node.js | `require("nodeosbxl")` | `nodeosbxl` / `dto` / `chart` / `enums` |
| Python | `import pyosbxl` | `pyosbxl` / `pyosbxl.dto` / `pyosbxl.chart` / `pyosbxl.enums` |
| Java | `import com.osboffice.osbxl.*` | `dto` / `chart` / `enums` 等のパッケージ |
| WebAssembly | `createOsbxl(createModule)` | `osbxl` / `dto` / `chart` / `enums` |

下表は Node.js 表記です。各 namespace の中身の数は4言語共通です。

| namespace | 中身 | 数 |
| --- | --- | --- |
| `nodeosbxl` | 操作の対象になるクラス（App / WorkBook / WorkSheet / Range / Font / Chart 系の入口 など） | 29 クラス |
| `dto` | データの受け渡しに使うオブジェクト（`InputValueObject`, `FontObject`, `ColorObject` など） | 40 クラス |
| `chart` | グラフ専用オブジェクト（`ChartObject`, `Axis`, `ChartTitle` など） | 43 クラス |
| `enums` | Excel 互換の列挙型（`XlFont`, `XlChartType`, `XlIndexColor` など） | 127 種 |

## オブジェクトモデル

```plaintext
App ← すべての入口。ブックを開く/作る、セルアドレスの変換
│
├─ createWorkBook() ─┐
├─ openWorkBook() ─┴→ WorkBook ← xlsx ファイル 1 つ
│ │
│ ├─ openWorkSheet() → WorkSheet ← シート 1 つ
│ │ │
│ │ ├─ getRange("A1:B2") → Range
│ │ │ ※ getCells(1,1,2,2) でも同じ Range が取れる
│ │ │ │
│ │ │ ├─ getValue() / setValue() / setFormula()
│ │ │ ├─ getFont() → Font → getColor() → Color
│ │ │ ├─ getBorders()→ Borders→ getBorder() → Border
│ │ │ ├─ getFill() → Fill
│ │ │ └─ getAlignment() → Alignment
│ │ │
│ │ ├─ getRow() / getCol() → Row / Col
│ │ ├─ getSort() / getAutoFilter()
│ │ ├─ getPivotTables() / getListObjects()
│ │ ├─ addChartObject() → chart.ChartObject
│ │ └─ getComments() / getHyperLinks() / getShapes() ...
│ │
│ ├─ addWorkSheet() / deleteWorkSheet() / moveSheet()
│ ├─ addCustomStyle() / addCustomTableStyle()
│ └─ save() / saveAs() / close()
│
└─ convertFromRowColNumber() などのユーティリティ
```

**この階層を常に上から辿る**のが osbxl の基本です。 `Range` より下（`Font`, `Color`, `Borders` …）は「取得して、そのオブジェクトの setter を呼ぶ」形になります。

## 基本の流れ

```plaintext
[1] new App() App を生成する
[2] createWorkBook(path) 新規ブックを作る
openWorkBook(path) 既存のブックを開く
[3] openWorkSheet(name) ワークシートを開く
[4] getRange("A1:B2") Range を取得する（getCells(1,1,2,2) でも同じ）
[5] setValue() / setFormula() 値・数式を入れる
[6] getValue(true) 結果を読み戻す
[7] save() / saveAs(path) 保存する
[8] close() ブックを閉じる
```

### コード

実行結果:

#### [Node.js]

```javascript
// basic-flow.js — 新規作成 → 編集 → 保存 → 開き直し → 別名保存 の一連の流れ
const { nodeosbxl } = require("nodeosbxl");
const app = new nodeosbxl.App(); // [1] App を生成
// ---------- 新規ブックを作って編集し、保存する ----------
let wb = app.createWorkBook("flow.xlsx"); // [2] 新規作成
console.log("[2] createWorkBook -> getBookPath =", wb.getBookPath());
let ws = wb.openWorkSheet("Sheet1"); // [3] ワークシートを開く
console.log("[3] openWorkSheet -> シート名 =", ws.getName());
// [4] Range を取得して [5] 値を設定
// 数値(行, 列)で取る ws.getCells(1, 1) も getRange("A1:A1") と同じ Range を返します
ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100\*3"); // [5] 数式を設定（先頭の "=" は付けない）
console.log("[6] getValue(true) -> B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);
wb.save(); // [7] 保存
console.log("[7] save() 完了");
wb.close(); // [8] 閉じる（WorkSheet の close は不要）
console.log("[8] close() 完了");
// ---------- 既存ブックを開いて、別名で保存する ----------
wb = app.openWorkBook("flow.xlsx"); // [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1"); // [3']
const read = ws.getRange("A1:B1").getValue(true);
console.log("[2'] openWorkBook -> A1 =", read["A1"], "/ B1 =", read["B1"]);
ws.getRange("A2:A2").setValue("開き直して追記");
wb.saveAs("flow-copy.xlsx"); // [7'] 別名で保存
console.log("[7'] saveAs() 完了 -> getBookPath =", wb.getBookPath());
wb.close(); // [8']
console.log("[8] close() 完了");
```

#### [Python]

```python
# basic-flow.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
app = pyosbxl.App() # [1] App を生成
wb = app.createWorkBook("flow-py.xlsx") # [2] 新規作成
ws = wb.openWorkSheet("Sheet1") # [3] ワークシートを開く
ws.getRange("A1:A1").setValue("新規作成")
ws.getRange("B1:B1").setFormula("100\*3") # [5] 数式（先頭の "=" は付けない）
print("[6] getValue(True) -> B1 =", ws.getRange("B1:B1").getValue(True)["B1"])
wb.save() # [7] 保存
wb.close() # [8] 閉じる
print("[8] close() 完了")
wb = app.openWorkBook("flow-py.xlsx") # [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1")
print("[2'] openWorkBook -> A1 =", ws.getRange("A1:A1").getValue(True)["A1"])
wb.close()
print("[8] close() 完了")
```

#### [Java]

```java
// basic-flow.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper(); // [1] App を生成
WorkBookWrapper wb = app.createWorkBook("flow-java.xlsx"); // [2] 新規作成
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1"); // [3] ワークシートを開く
ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100\*3"); // [5] 数式（先頭の "=" は付けない）
Map<String, String> v = ws.getRange("B1:B1").getValue(true);
System.out.println("[6] getValue(true) -> B1 = " + v.get("B1"));
wb.save(); // [7] 保存
wb.close(); // [8] 閉じる
System.out.println("[8] close() 完了");
wb = app.openWorkBook("flow-java.xlsx"); // [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1");
Map<String, String> r = ws.getRange("A1:A1").getValue(true);
System.out.println("[2'] openWorkBook -> A1 = " + r.get("A1"));
wb.close();
System.out.println("[8] close() 完了");
}
}
```

#### [PHP]

```php
// basic-flow.md — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App(); // [1] App を生成
// ---------- 新規ブックを作って編集し、保存する ----------
$wb = $app->createWorkBook("flow-php.xlsx"); // [2] 新規作成
echo "[2] createWorkBook -> getBookPath = " . $wb->getBookPath() . "\n";
$ws = $wb->openWorkSheet("Sheet1"); // [3] ワークシートを開く
echo "[3] openWorkSheet -> シート名 = " . $ws->getName() . "\n";
// [4] Range を取得して [5] 値を設定
// 数値(行, 列)で取る $ws->getCells(1, 1) も getRange("A1:A1") と同じ Range を返します
$ws->getRange("A1:A1")->setValue("新規作成");
$ws->getRange("B1:B1")->setFormula("100\*3"); // [5] 数式を設定（先頭の "=" は付けない）
$v = $ws->getRange("B1:B1")->getValue(true);
echo "[6] getValue(true) -> B1 = " . $v["B1"] . "\n";
$wb->save(); // [7] 保存
echo "[7] save() 完了\n";
$wb->close(); // [8] 閉じる（WorkSheet の close は不要）
echo "[8] close() 完了\n";
// ---------- 既存ブックを開いて、別名で保存する ----------
$wb = $app->openWorkBook("flow-php.xlsx"); // [2'] 既存のブックを開く
$ws = $wb->openWorkSheet("Sheet1"); // [3']
$read = $ws->getRange("A1:B1")->getValue(true);
echo "[2'] openWorkBook -> A1 = " . $read["A1"] . " / B1 = " . $read["B1"] . "\n";
$ws->getRange("A2:A2")->setValue("開き直して追記");
$wb->saveAs("flow-copy-php.xlsx"); // [7'] 別名で保存
echo "[7'] saveAs() 完了 -> getBookPath = " . $wb->getBookPath() . "\n";
$wb->close(); // [8']
echo "[8] close() 完了\n";
```

#### [WebAssembly]

```javascript
// basic-flow.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App(); // [1] App を生成
const wb = app.createWorkBook("flow-wasm.xlsx"); // [2] 新規作成
const ws = wb.openWorkSheet("Sheet1"); // [3] ワークシートを開く
ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100\*3"); // [5] 数式（先頭の "=" は付けない）
console.log("[6] getValue(true) -> B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);
wb.save(); // [7] 保存
wb.close(); // [8] 閉じる
console.log("[8] close() 完了");
const wb2 = app.openWorkBook("flow-wasm.xlsx"); // [2'] 既存のブックを開く
const ws2 = wb2.openWorkSheet("Sheet1");
console.log("[2'] openWorkBook -> A1 =", ws2.getRange("A1:A1").getValue(true)["A1"]);
wb2.close();
console.log("[8] close() 完了");
```

```plaintext
[2] createWorkBook -> getBookPath = <実行ディレクトリ>/flow.xlsx
[3] openWorkSheet -> シート名 = Sheet1
[6] getValue(true) -> B1 = 300
[7] save() 完了
[8] close() 完了
[2'] openWorkBook -> A1 = 新規作成 / B1 = 300
[7'] saveAs() 完了 -> getBookPath = <実行ディレクトリ>/flow.xlsx
[8] close() 完了
```

`flow.xlsx` と `flow-copy.xlsx` の 2 つができます。

### 流れのポイント

* **`App` は 1 プロセスに 1 つでよい。** ブックを複数扱う場合も同じ `app` から `createWorkBook()` / `openWorkBook()` を呼びます。
* **相対パスは絶対パスに解決される。** `createWorkBook("flow.xlsx")` の後、 `getBookPath()` はカレントディレクトリを含む絶対パスを返します。
* **`WorkSheet` を閉じる処理は無い。** `close()` は `WorkBook` のメソッドだけです。 シートは `wb.close()` でまとめて解放されます。
* **`save()` と `saveAs()` の違いに注意**（下記の「注意点」参照）。

## 主要メソッド

### App（10 メソッド）

`new nodeosbxl.App()` で生成します。ブックの開閉と、アドレス変換ユーティリティを持ちます。

| メソッド | 説明 |
| --- | --- |
| `getVersion()` | バージョン番号を返します。 |
| `createWorkBook(path, defaultFont?, defaultFontSize?)` | ワークブックを作成します。 |
| `openWorkBook(path)` | ワークブックを開きます。 |
| `openPasswordWorkBook(path, password)` | パスワード付きワークブックを開きます。 |
| `convertFromRowColNumber(row, col)` | セルの数値表記を A1C1 表記に変換します。`(1,1)` → `"A1"` |
| `convertFromRowColNumber2(startRow, startCol, endRow, endCol)` | セル範囲の数値表記を A1C1 表記に変換します。`(1,1,2,3)` → `"A1:C2"` |
| `convertToColumnNumber("B")` | 列の A1 表記を数値表記に変換します。→ `2` |
| `convertFromColumnNumber(3)` | 列の数値表記を A1 表記に変換します。→ `"C"` |
| `getNumericValue(dateTimeObject, is1904?)` | 日付時刻オブジェクトから Excel 内部のシリアル値を取得します。 |
| `calculateHexColor(bookPath, colorObj, isForeGroundColor)` | ワークブック固有の色情報を加味した HEX 値を取得します。 |

`convertFromRowColNumber` / `convertFromRowColNumber2` には絶対参照（`"$A$1"`）を返すオーバーロードもあります。

### WorkBook（27 メソッド）— 主要 12

xlsx ファイル 1 つに対応します。

| メソッド | 説明 |
| --- | --- |
| `getBookPath()` | ワークブックのファイルパスを返します。 |
| `openWorkSheet(sheetName)` | ワークシートを開きます。 |
| `openWorkSheetByIndex(sheetIndex)` | ワークシートをインデックスで開きます。 |
| `addWorkSheet(sheetName, position?)` | ワークシートを追加します。 |
| `deleteWorkSheet(sheetName)` | ワークシートを削除します。 |
| `moveSheet(sheetName, position)` | ワークシートを移動します。 |
| `getSheetCount()` | ワークシート数を取得します。 |
| `getSheetNumber(sheetName)` | ワークシートのシート番号を取得します。 |
| `save()` / `save(password)` | ワークブックを同一ファイルに保存します。 |
| `saveAs(path)` / `saveAs(path, password)` | ワークブックを別ファイルに保存します。 |
| `close()` | ワークブックを閉じます。 |
| `isDate1904()` | Excel の日付 1904 年形式であるかを取得します。 |

その他（15）: スタイル系 `getNames` `addCustomStyle` `getCustomStyle` `deleteCustomStyle` `addCustomTableStyle` `addPivotCustomTableStyle` `getCustomTableStyle` `deleteCustomTableStyle` / 外部ブック系 `addExternalWorkBook` `updateExternalWorkBook` `getExternalWorkBookPath` / ファイル属性 `setCompanyName` `setManagerName` `setCreateAuthor` `setLastAuthor`

### WorkSheet（44 メソッド）— 主要 14

シート 1 つに対応します。**ここから下のオブジェクト（Range / Sort / PivotTable / Chart …）へ辿る入口**です。

| メソッド | 説明 |
| --- | --- |
| `getRange("A1:B2")` | セル範囲クラスインスタンスを取得します。 |
| `getCells(row, col)` / `getCells(sr, sc, er, ec)` | セル範囲クラスインスタンスを数値指定で取得します。 |
| `getRow(rowNum)` | 行クラスインスタンスを取得します。 |
| `getCol(colNum)` | 列クラスインスタンスを取得します。 |
| `getName()` / `setName(name)` | シート名称の取得 / 設定を行います。 |
| `isActive()` / `activate(...)` | シートタブが起動時の表示項目かどうかの取得 / 設定を行います。 |
| `insertRow(startRow, numOfRows)` | 行の挿入を行います。 |
| `deleteRow(startRow, numOfRows)` | 行の削除を行います。 |
| `insertCol(startCol, numOfCols)` | 列の挿入を行います。 |
| `deleteCol(startCol, numOfCols)` | 列の削除を行います。 |
| `setValueArray(values)` | セル値の一括入力を行います。 |
| `setFormulaArray(formulas)` | 関数の一括入力を行います。 |

その他（30）: 機能オブジェクト取得 `getFormatConditions` `getAutoFilter` `getSort` `getComments` `getHyperLinks` `getListObjects` `getPivotTables` `getShapes` `getWindow` `getHPageBreaks` `getVPageBreaks` / グラフ `getChartObject` `addChartObject` `deleteChartObject` / 印刷設定 `getPageSetupObject` `setPageSetupObject` / 選択 `getActiveCell` `setActiveCell` / 行・列・セルの操作 `clearRow` `copyRow` `clearCol` `copyCol` `insertCell` `deleteCell` `clearCell` `copyCell` / アウトライン `setRowOutline` `clearRowOutline` `setColOutline` `clearColOutline`

### Range（31 メソッド）— 主要 16

セル範囲に対応します。**値・数式の読み書きと、書式オブジェクトへの入口**です。

| メソッド | 説明 |
| --- | --- |
| `getValue(rawValue?)` | セルの値の取得を行います。`true` で生値（数式の結果）。 |
| `getFormula()` | 数式の取得を行います。 |
| `setValue(value, forceString?, numberFormat?)` | セル値の設定を行います（汎用メソッド）。 |
| `setNumberValue(value, ...)` | セル値の設定を行います（数値設定メソッド）。 |
| `setDateValue(dateTimeObject, ...)` | セル値の設定を行います（日付/時刻設定メソッド）。 |
| `setDateStringValue(str, ...)` | セル値の設定を行います（日付/時刻を文字列から入力）。 |
| `setBooleanValue(bool, ...)` | セル値の設定を行います（Boolean 値設定メソッド）。 |
| `setFormula(formula, isArray?, setAllCell?)` | 数式の設定を行います。 |
| `getFont()` | フォントオブジェクトインスタンスを取得します。 |
| `getBorders()` | 枠線一括オブジェクトインスタンスを取得します。 |
| `getFill()` | フィルオブジェクトインスタンスを取得します。 |
| `getAlignment()` | セルの配置位置指定オブジェクトインスタンスを取得します。 |
| `setNumberFormat(format)` | 表示書式の設定を行います。 |
| `merge(mergeEachRow?)` | セル範囲をマージします。 |
| `clearContent()` | セルのクリアを行います。 |
| `getAddress()` | Range の選択範囲を取得します。 |

その他（15）: `getProtection` `getCharacters` `getPhonetics` `getNumberFormat` `clearNumberFormat` `clearAllFormat` `setDataTable` `clearDataTable` `setBuiltinStyle` `setCustomStyle` `getTotalWidth` `getTotalColumnWidth` `getTotalRowHeight` `unMerge` `replaceValue`

## 解説

* **`Range` は「取得して使い捨て」でよい。** `ws.getRange("A1:A1")` は呼ぶたびに新しいインスタンスを返すので、 変数に保持し続ける必要はありません。
* **`getValue()` の戻り値はマップ。** `{ "A1": "100", "B1": "300" }` の形で、値はすべて**文字列**です。  
   数値として扱うには `Number(...)` で変換します。
* **Range の取り方は 2 通り。**
  + `ws.getRange("A1:B2")` — **文字列アドレス**で指定。単一セルでも `"A1:A1"` と範囲の形で渡します。
  + `ws.getCells(row, col)` / `ws.getCells(startRow, startCol, endRow, endCol)` — **数値（行, 列）**で指定。 行・列とも **1 始まり**で `(1,1)` = A1、`(3,1,3,3)` = A3:C3 です。 ループで行列番号を回すときはアドレス文字列を組み立てずに済むので便利です。
  + 数値 ⇔ 文字列の変換は `app.convertFromRowColNumber(row, col)` / `app.convertFromRowColNumber2(sr, sc, er, ec)` / `app.convertToColumnNumber("B")` を使います。
* **数式は実行時に計算される。** `setFormula()` の直後に `getValue(true)` で計算結果を読めます （`save()` を待つ必要はありません）。

## 注意点

* **ワークブックは、一意性を確保するため内部でロックされ、この処理によって同一ワークブックの同時読込・書込は起こりません。**
* **`close()` は `WorkBook` のメソッドだけ。** `App` / `WorkSheet` / `Range` に close 系のメソッドはありません。
* **`saveAs()` は「その時点の内容を別ファイルに書き出す」だけで、以降の保存先は切り替わりません。** Excel VBA の `SaveAs` とは挙動が違います。

  ```plaintext
  wb.save() → base.xlsx に書かれる
  wb.saveAs("copy.xlsx") → copy.xlsx に「この時点の内容」が書かれる
  wb.getBookPath() → base.xlsx のまま（切り替わらない）
  ws.getRange(...).setValue(...)
  wb.save() → base.xlsx に書かれる（copy.xlsx ではない）
  ```
* **数式に先頭の `=` を付けない。** `setFormula("SUM(A1:A3)")` です。
* **起動時にライセンス表示が stdout に 1 行出力されます**（プロセスにつき 1 回）。  
   標準出力をパースする処理を書く場合は注意してください。

## 関連

* [目次](index.md)
* 次のサンプル: [値と数式を入れて表を作る](general-02-values-formulas.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
