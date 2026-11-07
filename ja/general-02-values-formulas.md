# 値と数式を入れて表を作る

level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-02-values-formulas.md)


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


各言語のモジュールを[ダウンロードページ](/pages/download.md)から入手してインストールします （入口の表記は [概要の前提](general-01-overview.md#前提) 表を参照）。


PHP の OS 別の詳しい導入手順（Linux / macOS / Windows、Composer が未インストールの場合を含む）は[クイックスタート](/pages/quickstart.md)を参照してください。


WebAssembly 版は MEMORY64（Wasm 3.0）ビルドのため、実行環境側の対応が必要です （Chrome/Edge 133〜・Firefox 134〜・Node.js 24〜。Safari は現在未対応）。


保存先パスを省略形（`"sales.xlsx"`）にすると、**プロセスのカレントディレクトリ**に書き出されます。


#### [Node.js]

```bash
npm install ./nodeosbxl-2.0.0.tgz
```


#### [Python]

```bash
# (例) Python 3.12 以上（abi3 wheel 1 本で 3.12 以降をカバー）
pip install ./pyosbxl-2.0.0-cp312-abi3-linux_x86_64.whl
```


#### [Java]

```bash
# JAR をクラスパスに追加（javacpp-1.5.14.jar も必要）
javac -cp osbxl-2.0.0.jar:javacpp-1.5.14.jar Main.java
java  -cp osbxl-2.0.0.jar:javacpp-1.5.14.jar:. Main
```


#### [PHP]

```bash
# HPからダウンロードした適切なOS・PHPバージョンのzipを作業ディレクトリへ配置
# (例: ~/phposbxl-test に配置した場合)
composer config repositories.phposbxl artifact ~/phposbxl-test/
composer require osb-office/phposbxl
vendor/bin/phposbxl-install
```


#### [WebAssembly]

```bash
# パッケージをインストール
npm install wasmosbxl-2.0.0.tgz
```


## コード


実行結果:


#### [Node.js]

```javascript
// values-formulas.js — ブックを作って、値と数式を入れて保存する
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sales.xlsx");   // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1");         // 既定のシートを開く

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
// 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((_, i) => [`B${i + 2}*C${i + 2}`]);
ws.setRangeFormulas("D2", formulas);

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
wb = app.createWorkBook("sales-py.xlsx")   # 新規ブック作成
ws = wb.openWorkSheet("Sheet1")            # 既定のシートを開く

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
# 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
# 数式文字列に先頭の "=" は付けない。
formulas = [[f"B{i + 2}*C{i + 2}"] for i in range(len(data))]
ws.setRangeFormulas("D2", formulas)

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
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("sales-java.xlsx");   // 新規ブック作成
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");             // 既定のシートを開く

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
        // 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
        // 数式文字列に先頭の "=" は付けない。
        List<List<String>> formulas = new ArrayList<>();
        for (int i = 0; i < names.length; i++) {
            formulas.add(Collections.singletonList("B" + (i + 2) + "*C" + (i + 2)));
        }
        ws.setRangeFormulas("D2", formulas);

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
// values-formulas.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sales-php.xlsx");   // 新規ブック作成
$ws = $wb->openWorkSheet("Sheet1");             // 既定のシートを開く

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
// 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
// 数式文字列に先頭の "=" は付けない。
$formulas = [];
for ($i = 0; $i < count($data); $i++) {
    $row = $i + 2;
    $formulas[] = ["B{$row}*C{$row}"];
}
$ws->setRangeFormulas("D2", $formulas);

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
const wb = app.createWorkBook("sales-wasm.xlsx");   // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1");              // 既定のシートを開く

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
// 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((_, i) => [`B${i + 2}*C${i + 2}`]);
ws.setRangeFormulas("D2", formulas);

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


`金額 D2:D4 = 300, 300, 300
合計: 900`

<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`new nodeosbxl.App()` がすべての入口。** `App` から `createWorkBook()` / `openWorkBook()` で `WorkBook` を取得し、`openWorkSheet()` で `WorkSheet`、`getRange()` で `Range` を辿ります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**セルアドレスは `"A1:B2"` 形式の文字列。** 単一セルでも `"A1:A1"` と範囲の形で渡します。** 行番号・列番号（1始まり）から作る場合は `app.convertFromRowColNumber(row, col)`、 範囲なら `app.convertFromRowColNumber2(startRow, startCol, endRow, endCol)` を使います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>値の種類ごとに setter が分かれています。** `setValue(文字列)` / `setNumberValue(数値)` / `setDateValue(DateTimeObject)` / `setDateStringValue(日付文字列)` / `setBooleanValue(真偽値)`。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`getValue(true)` の `true` は「生値」**（数式なら計算結果）。**省略すると表示値（書式適用後）になります。  
 戻り値は `{ "A1": "100", ... }` のマップで、値はすべて文字列です。  
数値比較する場合は `Number(...)` で変換します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>`wb.save()` → `wb.close()` で終了。** 別ファイルに書き出したいときは `wb.saveAs(path)` を使いますが、 `saveAs()` は**以降の保存先を切り替えません**（[概要の注意点](general-01-overview.md#注意点)参照）。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**数式に先頭の `=` を付けない。** `setFormula("=SUM(A1:A3)")` ではなく `setFormula("SUM(A1:A3)")` です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**複数セル Range への `setValue` / `setNumberValue` は「全セルに同じ値」を入れます。** `ws.getRange("A1:C1").setValue("X")` → A1, B1, C1 すべて `"X"` になります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`setFormula()` は相対参照を自動展開しません。** `ws.getRange("B1:B3").setFormula("A1*2")` は **B1 だけ**に数式が入り、B2 / B3 は空のままです。** 第3引数 `setAllCell` を true にした `setFormula("A1*2", false, true)` も 「同じ文字列をそのまま全セルにコピーする」だけで、`A2*2`, `A3*2` には<strong>なりません** （結果 B1:B3 がすべて同じ値になります）。** 行ごとにずらした数式を入れたい場合は、このサンプルのように <strong>`setRangeFormulas()`（2次元配列で高速一括流し込み）** を使ってください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**起動時にライセンス表示が stdout に 1 行出力されます**（プロセスにつき 1 回）。  
 標準出力をパースする処理を書く場合は注意してください。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li>[目次](index.md)</li>
  <li>前: [概要 — オブジェクトモデルと基本の流れ](general-01-overview.md) / 次: [ブックとワークシートの操作](general-03-sheets.md)</li>
<li>
  API リファレンス
  
    
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
    
    [App](/api-docs/nodejs/classes/nodeosbxl.App.html) / [WorkBook](/api-docs/nodejs/classes/nodeosbxl.WorkBook.html) / [WorkSheet](/api-docs/nodejs/classes/nodeosbxl.WorkSheet.html) / [Range](/api-docs/nodejs/classes/nodeosbxl.Range.html)


[App](/api-docs/python/pyosbxl.App.html) / [WorkBook](/api-docs/python/pyosbxl.WorkBook.html) / [WorkSheet](/api-docs/python/pyosbxl.WorkSheet.html) / [Range](/api-docs/python/pyosbxl.Range.html)


[App](/api-docs/java/com/osboffice/osbxl/AppWrapper.html) / [WorkBook](/api-docs/java/com/osboffice/osbxl/WorkBookWrapper.html) / [WorkSheet](/api-docs/java/com/osboffice/osbxl/WorkSheetWrapper.html) / [Range](/api-docs/java/com/osboffice/osbxl/RangeWrapper.html)


[App](/api-docs/php/classes/phposbxl-App.html) / [WorkBook](/api-docs/php/classes/phposbxl-WorkBook.html) / [WorkSheet](/api-docs/php/classes/phposbxl-WorkSheet.html) / [Range](/api-docs/php/classes/phposbxl-Range.html)


[App](/api-docs/wasm/classes/osbxl.App.html) / [WorkBook](/api-docs/wasm/classes/osbxl.WorkBook.html) / [WorkSheet](/api-docs/wasm/classes/osbxl.WorkSheet.html) / [Range](/api-docs/wasm/classes/osbxl.Range.html)
