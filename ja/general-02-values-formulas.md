# 値と数式を入れて表を作る

level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](/en/general-02-values-formulas.md)

## やりたいこと

新しい xlsx を作って、表（見出し・文字列・数値・数式）を書き込み、ファイルに保存する。 [概要](general-01-overview.md)の基本の流れを実際の表作成に当てはめたものです。

## 完成イメージ

`sales.xlsx` の `Sheet1` に次の表ができます。

|  | A | B | C | D |
| --- | --- | --- | --- | --- |
| **1** | 商品 | 単価 | 個数 | 金額 |
| **2** | りんご | 100 | 3 | 300 |
| **3** | みかん | 50 | 6 | 300 |
| **4** | ぶどう | 150 | 2 | 300 |
| **5** |  |  | 合計 | 900 |

D 列は数式 `B*C`、D5 は `SUM(D2:D4)` です。

## 前提

各言語のモジュールを[ダウンロードページ](https://osb-office.com/pages/download.php)から入手してインストールします （入口の表記は [概要の前提](general-01-overview.md#前提) 表を参照）。

PHP の OS 別の詳しい導入手順（Linux / macOS / Windows、Composer が未インストールの場合を含む）は[クイックスタート](https://osb-office.com/pages/quickstart.php)を参照してください。

WebAssembly 版は MEMORY64（Wasm 3.0）ビルドのため、実行環境側の対応が必要です （Chrome/Edge 133〜・Firefox 134〜・Node.js 24〜。Safari は現在未対応）。

保存先パスを省略形（`"sales.xlsx"`）にすると、**プロセスのカレントディレクトリ**に書き出されます。

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

## コード

実行結果:

#### [Node.js]

```javascript
// values-formulas.js — ブックを作って、値と数式を入れて保存する
const { nodeosbxl, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("sales.xlsx"); // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1"); // 既定のシートを開く
// --- 見出し (1行目) ---
// app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
["商品", "単価", "個数", "金額"].forEach((text, i) => {
const cell = app.convertFromRowColNumber(1, i + 1);
ws.getRange(`${cell}:${cell}`).setValue(text);
});
// --- 明細 (2〜4行目) ---
const data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
data.forEach(([name, price, qty], i) => {
const row = i + 2;
ws.getRange(`A${row}:A${row}`).setValue(name);
ws.getRange(`B${row}:B${row}`).setNumberValue(price);
ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});
// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((\_, i) => {
const f = new dto.InputFormulaObject();
f.setA1C1(app.convertFromRowColNumber(i + 2, 4)); // D2, D3, D4
f.setFormula(`B${i + 2}\*C${i + 2}`); // "B2\*C2" ...
return f;
});
ws.setFormulaArray(formulas);
// --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");
// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
const values = ws.getRange("A1:D5").getValue(true);
console.log("金額 D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("合計:", values["D5"]);
// --- 保存して閉じる ---
wb.save();
wb.close();
```

#### [Python]

```python
# values-formulas.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("sales-py.xlsx") # 新規ブック作成
ws = wb.openWorkSheet("Sheet1") # 既定のシートを開く
# --- 見出し (1行目) ---
# app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
for i, text in enumerate(["商品", "単価", "個数", "金額"]):
cell = app.convertFromRowColNumber(1, i + 1)
ws.getRange(f"{cell}:{cell}").setValue(text)
# --- 明細 (2〜4行目) ---
data = [("りんご", 100, 3), ("みかん", 50, 6), ("ぶどう", 150, 2)]
for i, (name, price, qty) in enumerate(data):
row = i + 2
ws.getRange(f"A{row}:A{row}").setValue(name)
ws.getRange(f"B{row}:B{row}").setNumberValue(price)
ws.getRange(f"C{row}:C{row}").setNumberValue(qty)
# --- 金額列 (D2:D4) ---
# 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
# 数式文字列に先頭の "=" は付けない。
formulas = []
for i in range(len(data)):
f = pyosbxl.dto.InputFormulaObject()
f.setA1C1(app.convertFromRowColNumber(i + 2, 4)) # D2, D3, D4
f.setFormula(f"B{i + 2}\*C{i + 2}") # "B2\*C2" ...
formulas.append(f)
ws.setFormulaArray(formulas)
# --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計")
ws.getRange("D5:D5").setFormula("SUM(D2:D4)")
# --- 読み戻し ---
# getValue(True) は「生値」= 数式の計算結果を返す
values = ws.getRange("A1:D5").getValue(True)
print("金額 D2:D4 =", ", ".join([values["D2"], values["D3"], values["D4"]]))
print("合計:", values["D5"])
# --- 保存して閉じる ---
wb.save()
wb.close()
```

#### [Java]

```java
// values-formulas.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("sales-java.xlsx"); // 新規ブック作成
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1"); // 既定のシートを開く
// --- 見出し (1行目) ---
// app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
String[] headers = { "商品", "単価", "個数", "金額" };
for (int i = 0; i < headers.length; i++) {
String cell = app.convertFromRowColNumber(1, i + 1);
ws.getRange(cell + ":" + cell).setValue(headers[i]);
}
// --- 明細 (2〜4行目) ---
String[] names = { "りんご", "みかん", "ぶどう" };
double[] prices = { 100, 50, 150 };
double[] qtys = { 3, 6, 2 };
for (int i = 0; i < names.length; i++) {
int row = i + 2;
ws.getRange("A" + row + ":A" + row).setValue(names[i]);
ws.getRange("B" + row + ":B" + row).setNumberValue(prices[i]);
ws.getRange("C" + row + ":C" + row).setNumberValue(qtys[i]);
}
// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
List<InputFormulaObjectWrapper> formulas = new ArrayList<>();
for (int i = 0; i < names.length; i++) {
InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
f.setA1C1(app.convertFromRowColNumber(i + 2, 4)); // D2, D3, D4
f.setFormula("B" + (i + 2) + "\*C" + (i + 2)); // "B2\*C2" ...
formulas.add(f);
}
ws.setFormulaArray(formulas);
// --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");
// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
Map<String, String> values = ws.getRange("A1:D5").getValue(true);
System.out.println("金額 D2:D4 = "
+ String.join(", ", values.get("D2"), values.get("D3"), values.get("D4")));
System.out.println("合計: " + values.get("D5"));
// --- 保存して閉じる ---
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// values-formulas.md — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sales-php.xlsx"); // 新規ブック作成
$ws = $wb->openWorkSheet("Sheet1"); // 既定のシートを開く
// --- 見出し (1行目) ---
// $app->convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
$headers = ["商品", "単価", "個数", "金額"];
foreach ($headers as $i => $text) {
$cell = $app->convertFromRowColNumber(1, $i + 1);
$ws->getRange("{$cell}:{$cell}")->setValue($text);
}
// --- 明細 (2〜4行目) ---
$data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
foreach ($data as $i => [$name, $price, $qty]) {
$row = $i + 2;
$ws->getRange("A{$row}:A{$row}")->setValue($name);
$ws->getRange("B{$row}:B{$row}")->setNumberValue($price);
$ws->getRange("C{$row}:C{$row}")->setNumberValue($qty);
}
// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
$formulas = [];
for ($i = 0; $i < count($data); $i++) {
$f = new \phposbxl\dto\InputFormulaObject();
$row = $i + 2;
$f->setA1C1($app->convertFromRowColNumber($i + 2, 4)); // D2, D3, D4
$f->setFormula("B{$row}\*C{$row}"); // "B2\*C2" ...
$formulas[] = $f;
}
$ws->setFormulaArray($formulas);
// --- 合計 (D5) ---
$ws->getRange("C5:C5")->setValue("合計");
$ws->getRange("D5:D5")->setFormula("SUM(D2:D4)");
// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
$values = $ws->getRange("A1:D5")->getValue(true);
echo "金額 D2:D4 = " . implode(", ", [$values["D2"], $values["D3"], $values["D4"]]) . "\n";
echo "合計: " . $values["D5"] . "\n";
// --- 保存して閉じる ---
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// values-formulas.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("sales-wasm.xlsx"); // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1"); // 既定のシートを開く
// --- 見出し (1行目) ---
// app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
["商品", "単価", "個数", "金額"].forEach((text, i) => {
const cell = app.convertFromRowColNumber(1, i + 1);
ws.getRange(`${cell}:${cell}`).setValue(text);
});
// --- 明細 (2〜4行目) ---
const data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
data.forEach(([name, price, qty], i) => {
const row = i + 2;
ws.getRange(`A${row}:A${row}`).setValue(name);
ws.getRange(`B${row}:B${row}`).setNumberValue(price);
ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});
// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((\_, i) => {
const f = new dto.InputFormulaObject();
f.setA1C1(app.convertFromRowColNumber(i + 2, 4)); // D2, D3, D4
f.setFormula(`B${i + 2}\*C${i + 2}`); // "B2\*C2" ...
return f;
});
ws.setFormulaArray(formulas);
// --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");
// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
const values = ws.getRange("A1:D5").getValue(true);
console.log("金額 D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("合計:", values["D5"]);
// --- 保存して閉じる ---
wb.save();
wb.close();
```

```plaintext
金額 D2:D4 = 300, 300, 300
合計: 900
```

## 解説

* **`new nodeosbxl.App()` がすべての入口。** `App` から `createWorkBook()` / `openWorkBook()` で `WorkBook` を取得し、`openWorkSheet()` で `WorkSheet`、`getRange()` で `Range` を辿ります。
* **セルアドレスは `"A1:B2"` 形式の文字列。** 単一セルでも `"A1:A1"` と範囲の形で渡します。  
   行番号・列番号（1始まり）から作る場合は `app.convertFromRowColNumber(row, col)`、 範囲なら `app.convertFromRowColNumber2(startRow, startCol, endRow, endCol)` を使います。
* **値の種類ごとに setter が分かれています。** `setValue(文字列)` / `setNumberValue(数値)` / `setDateValue(DateTimeObject)` / `setDateStringValue(日付文字列)` / `setBooleanValue(真偽値)`。
* **`getValue(true)` の `true` は「生値」**（数式なら計算結果）。  
  省略すると表示値（書式適用後）になります。  
   戻り値は `{ "A1": "100", ... }` のマップで、値はすべて文字列です。  
  数値比較する場合は `Number(...)` で変換します。
* **`wb.save()` → `wb.close()` で終了。** 別ファイルに書き出したいときは `wb.saveAs(path)` を使いますが、 `saveAs()` は**以降の保存先を切り替えません**（[概要の注意点](general-01-overview.md#注意点)参照）。

## 注意点

* **数式に先頭の `=` を付けない。** `setFormula("=SUM(A1:A3)")` ではなく `setFormula("SUM(A1:A3)")` です。
* **複数セル Range への `setValue` / `setNumberValue` は「全セルに同じ値」を入れます。** `ws.getRange("A1:C1").setValue("X")` → A1, B1, C1 すべて `"X"` になります。
* **`setFormula()` は相対参照を自動展開しません。** `ws.getRange("B1:B3").setFormula("A1*2")` は **B1 だけ**に数式が入り、B2 / B3 は空のままです。  
   第3引数 `setAllCell` を true にした `setFormula("A1*2", false, true)` も 「同じ文字列をそのまま全セルにコピーする」だけで、`A2*2`, `A3*2` には**なりません** （結果 B1:B3 がすべて同じ値になります）。  
   行ごとにずらした数式を入れたい場合は、このサンプルのように **`setFormulaArray()` + `dto.InputFormulaObject`（セルごとに A1C1 と数式を指定）** を使ってください。
* **起動時にライセンス表示が stdout に 1 行出力されます**（プロセスにつき 1 回）。  
   標準出力をパースする処理を書く場合は注意してください。

## 関連

* [目次](index.md)
* 前: [概要 — オブジェクトモデルと基本の流れ](general-01-overview.md) / 次: [ブックとワークシートの操作](general-03-sheets.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
