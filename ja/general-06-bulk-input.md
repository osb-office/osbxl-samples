# 大量データの一括高速投入 — setRangeData / setRangeFormulas

level: intermediate ／ verified: 2026-10-11 ／ 5言語タブ対応 ／ [English](../en/general-06-bulk-input.md)

## やりたいこと

何千・何万セルもの値や数式を、セルごとに `getCells()` するのではなく **2次元配列にまとめて一度に高速投入する**。CSV やデータベースの検索結果などの行列表を、最速かつ最小のメモリでワークシートに一括反映します。

| やり方 | 使うもの | 特徴 |
| --- | --- | --- |
| セルごとに投入 | `ws.getCells(r, c).setNumberValue(...)` | 少量セル向け。呼び出し回数に比例してオーバーヘッドが増加します。 |
| **2次元配列で値投入（最高速・推奨）** | `ws.setRangeData(startA1, data)` | **最速・推奨。** 行列データを素の配列で一括転送。数値文字列は自動的に数値セルとして保存。 |
| **2次元配列で数式投入（最高速・推奨）** | `ws.setRangeFormulas(startA1, formulas)` | **最速・推奨。** 計算列の数式群を 2次元配列で一度に流し込む（先頭の "=" は不要）。 |

## 対象クラス

### `WorkSheet`

| メソッド | 説明 |
| --- | --- |
| `setRangeData(startA1, data)` | 左上開始セル（例: "A1"）から、2次元配列（各行・各列の文字列）の値を一括高速入力します。 |
| `setRangeFormulas(startA1, formulas)` | 左上開始セル（例: "C1"）から、2次元配列（各行・各列の数式文字列、"="不要）を一括高速入力します。 |

## コード

実行結果:

#### [Node.js]

```javascript
// bulk-input.js — 2次元配列による値と数式の一括高速投入 (nodeosbxl)
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
// 各行・各列の値を文字列の2次元配列として用意し、左上セルを指定して一度に流し込みます。
// 数値形式の文字列 ("100" 等) は自動的に数値セルとして正しく認識・保存されます。
const data = [
    ["商品名", "単価", "数量"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
];
ws.setRangeData("A1", data);

console.log("[1] setRangeData 投入後:");
for (const a1 of ["A1", "B2", "C2"]) {
    console.log(`    ${a1} = ${JSON.stringify(ws.getRange(a1).getValue()[a1])}`);
}

// ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
// 計算列の数式群を2次元配列で一度に流し込みます（先頭の "=" は不要）。
const formulas = [
    ["金額"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
ws.setRangeFormulas("D1", formulas);

console.log("[2] setRangeFormulas 投入後:");
console.log("    数式 D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getFormula()));
console.log("    計算値 D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getValue(true)));

// ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
const ROWS = 2000, COLS = 5;

// セルごとに getCells() して投入
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

// 2次元配列で setRangeData() 投入 (最高速)
const wsBulk = wb.addWorkSheet("Bulk");
t = Date.now();
const matrix = [];
for (let r = 1; r <= ROWS; r++) {
    const row = [];
    for (let c = 1; c <= COLS; c++) row.push(String(r * 100 + c));
    matrix.push(row);
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setRangeData("A1", matrix);
const msApply = Date.now() - t;

console.log(`[3] ${ROWS * COLS} セルの投入速度比較:`);
console.log(`    セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setRangeData (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms (最速)`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === Number(b) ? b : `不一致 ${a} / ${b}`}`);

wb.save();
wb.close();
```

#### [Python]

```python
# bulk-input.py — 2次元配列による値と数式の一括高速投入 (pyosbxl)
import json
import time
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("bulk-input.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
data = [
    ["商品名", "単価", "数量"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
]
ws.setRangeData("A1", data)

print("[1] setRangeData 投入後:")
for a1 in ["A1", "B2", "C2"]:
    print(f"    {a1} = {json.dumps(ws.getRange(a1).getValue()[a1], ensure_ascii=False)}")

# ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
formulas = [
    ["金額"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
]
ws.setRangeFormulas("D1", formulas)

print("[2] setRangeFormulas 投入後:")
print("    数式 D2:D4 =", json.dumps(ws.getRange("D2:D4").getFormula(), ensure_ascii=False))
print("    計算値 D2:D4 =", json.dumps(ws.getRange("D2:D4").getValue(True), ensure_ascii=False))

# ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
ROWS, COLS = 2000, 5

# セルごとに getCells() して投入
ws_cell = wb.addWorkSheet("CellByCell")
t = time.perf_counter()
for r in range(1, ROWS + 1):
    for c in range(1, COLS + 1):
        ws_cell.getCells(r, c).setNumberValue(r * 100 + c)
ms_cell = int((time.perf_counter() - t) * 1000)

# 2次元配列で setRangeData() 投入 (最高速)
ws_bulk = wb.addWorkSheet("Bulk")
t = time.perf_counter()
matrix = [[str(r * 100 + c) for c in range(1, COLS + 1)] for r in range(1, ROWS + 1)]
ms_build = int((time.perf_counter() - t) * 1000)
t = time.perf_counter()
ws_bulk.setRangeData("A1", matrix)
ms_apply = int((time.perf_counter() - t) * 1000)

print(f"[3] {ROWS * COLS} セルの投入速度比較:")
print(f"    セルごと getCells().setNumberValue() : {ms_cell} ms")
print(f"    setRangeData (構築{ms_build} + 投入{ms_apply}) : {ms_build + ms_apply} ms (最速)")

a = ws_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[検証] 両方式とも E2000 = {b if a == float(b) else f'不一致 {a} / {b}'}")

wb.save()
wb.close()
```

#### [Java]

```java
// bulk-input.java — 2次元配列による値と数式の一括高速投入 (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        App app = new App();
        WorkBookWrapper wb = app.createWorkBook("bulk-input.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
        List<List<String>> data = Arrays.asList(
            Arrays.asList("商品名", "単価", "数量"),
            Arrays.asList("りんご", "100", "3"),
            Arrays.asList("みかん", "50", "6"),
            Arrays.asList("ぶどう", "150", "2")
        );
        ws.setRangeData("A1", data);

        System.out.println("[1] setRangeData 投入後:");
        for (String a1 : Arrays.asList("A1", "B2", "C2")) {
            System.out.println("    " + a1 + " = " + ws.getRange(a1).getValue().get(a1));
        }

        // ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
        List<List<String>> formulas = Arrays.asList(
            Collections.singletonList("金額"),
            Collections.singletonList("B2*C2"),
            Collections.singletonList("B3*C3"),
            Collections.singletonList("B4*C4")
        );
        ws.setRangeFormulas("D1", formulas);

        System.out.println("[2] setRangeFormulas 投入後:");
        System.out.println("    数式 D2:D4 = " + ws.getRange("D2:D4").getFormula());
        System.out.println("    計算値 D2:D4 = " + ws.getRange("D2:D4").getValue(true));

        // ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
        final int ROWS = 2000, COLS = 5;

        // セルごとに getCells() して投入
        WorkSheetWrapper wsCell = wb.addWorkSheet("CellByCell");
        long t = System.currentTimeMillis();
        for (int r = 1; r <= ROWS; r++) {
            for (int c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
        }
        long msCell = System.currentTimeMillis() - t;

        // 2次元配列で setRangeData() 投入 (最高速)
        WorkSheetWrapper wsBulk = wb.addWorkSheet("Bulk");
        t = System.currentTimeMillis();
        List<List<String>> matrix = new ArrayList<>();
        for (int r = 1; r <= ROWS; r++) {
            List<String> row = new ArrayList<>();
            for (int c = 1; c <= COLS; c++) row.add(String.valueOf(r * 100 + c));
            matrix.add(row);
        }
        long msBuild = System.currentTimeMillis() - t;
        t = System.currentTimeMillis();
        wsBulk.setRangeData("A1", matrix);
        long msApply = System.currentTimeMillis() - t;

        System.out.println("[3] " + (ROWS * COLS) + " セルの投入速度比較:");
        System.out.println("    セルごと getCells().setNumberValue() : " + msCell + " ms");
        System.out.println("    setRangeData (構築" + msBuild + " + 投入" + msApply + ") : " + (msBuild + msApply) + " ms (最速)");

        String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
        String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
        System.out.println("[検証] 両方式とも E2000 = " + (a.equals(b) ? b : "不一致 " + a + " / " + b));

        wb.save();
        wb.close();
    }
}
```

#### [PHP]

```php
// bulk-input.php — 2次元配列による値と数式の一括高速投入 (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("bulk-input.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

$now = fn(): int => (int)(hrtime(true) / 1000000);   // ミリ秒

// ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
$data = [
    ["商品名", "単価", "数量"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
];
$ws->setRangeData("A1", $data);

echo "[1] setRangeData 投入後:\n";
foreach (["A1", "B2", "C2"] as $a1) {
    echo "    {$a1} = ", json_encode($ws->getRange($a1)->getValue()[$a1], JSON_UNESCAPED_UNICODE), "\n";
}

// ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
$formulas = [
    ["金額"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
$ws->setRangeFormulas("D1", $formulas);

echo "[2] setRangeFormulas 投入後:\n";
echo "    数式 D2:D4 = ", json_encode($ws->getRange("D2:D4")->getFormula(), JSON_UNESCAPED_UNICODE), "\n";
echo "    計算値 D2:D4 = ", json_encode($ws->getRange("D2:D4")->getValue(true), JSON_UNESCAPED_UNICODE), "\n";

// ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
$ROWS = 2000;
$COLS = 5;

// セルごとに getCells() して投入
$wsCell = $wb->addWorkSheet("CellByCell");
$t = $now();
for ($r = 1; $r <= $ROWS; $r++) {
    for ($c = 1; $c <= $COLS; $c++) $wsCell->getCells($r, $c)->setNumberValue($r * 100 + $c);
}
$msCell = $now() - $t;

// 2次元配列で setRangeData() 投入 (最高速)
$wsBulk = $wb->addWorkSheet("Bulk");
$t = $now();
$matrix = [];
for ($r = 1; $r <= $ROWS; $r++) {
    $row = [];
    for ($c = 1; $c <= $COLS; $c++) $row[] = (string)($r * 100 + $c);
    $matrix[] = $row;
}
$msBuild = $now() - $t;
$t = $now();
$wsBulk->setRangeData("A1", $matrix);
$msApply = $now() - $t;

echo "[3] ", $ROWS * $COLS, " セルの投入速度比較:\n";
echo "    セルごと getCells().setNumberValue() : {$msCell} ms\n";
echo "    setRangeData (構築{$msBuild} + 投入{$msApply}) : ", $msBuild + $msApply, " ms (最速)\n";

$a = $wsCell->getRange("E2000:E2000")->getValue(true)["E2000"];
$b = $wsBulk->getRange("E2000:E2000")->getValue(true)["E2000"];
echo "[検証] 両方式とも E2000 = ", $a === (int)$b ? $b : "不一致 {$a} / {$b}", "\n";

$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// bulk-input.wasm.js — 2次元配列による値と数式の一括高速投入 (wasmosbxl / embind)
const app = new osbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
const data = [
    ["商品名", "単価", "数量"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
];
ws.setRangeData("A1", data);

console.log("[1] setRangeData 投入後:");
for (const a1 of ["A1", "B2", "C2"]) {
    console.log(`    ${a1} = ${JSON.stringify(ws.getRange(a1).getValue()[a1])}`);
}

// ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
const formulas = [
    ["金額"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
ws.setRangeFormulas("D1", formulas);

console.log("[2] setRangeFormulas 投入後:");
console.log("    数式 D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getFormula()));
console.log("    計算値 D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getValue(true)));

// ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
const ROWS = 2000, COLS = 5;

// セルごとに getCells() して投入
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

// 2次元配列で setRangeData() 投入 (最高速)
const wsBulk = wb.addWorkSheet("Bulk");
t = Date.now();
const matrix = [];
for (let r = 1; r <= ROWS; r++) {
    const row = [];
    for (let c = 1; c <= COLS; c++) row.push(String(r * 100 + c));
    matrix.push(row);
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setRangeData("A1", matrix);
const msApply = Date.now() - t;

console.log(`[3] ${ROWS * COLS} セルの投入速度比較:`);
console.log(`    セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setRangeData (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms (最速)`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === Number(b) ? b : `不一致 ${a} / ${b}`}`);

wb.save();
wb.close();
```

```plaintext
[1] setRangeData 投入後:
    A1 = "商品名"
    B2 = 100
    C2 = 3
[2] setRangeFormulas 投入後:
    数式 D2:D4 = {"D2":"B2*C2","D3":"B3*C3","D4":"B4*C4"}
    計算値 D2:D4 = {"D2":300,"D3":300,"D4":300}
[3] 10000 セルの投入速度比較:
    セルごと getCells().setNumberValue() : 102 ms
    setRangeData (構築2 + 投入18) : 20 ms (最速)
[検証] 両方式とも E2000 = 200005
```

`[3]` の ms は実行環境によって変わります（上記は 1 回の実行例）。

## 解説

### 速度とパフォーマンス

| セル数 | セルごと `getCells().setNumberValue()` | `setRangeData()` | 高速化倍率 |
| --- | --- | --- | --- |
| 10,000（2,000 行 x 5 列） | 約 102 ms | 約 20 ms（構築 2 + 投入 18） | **約 5 倍高速** |
| 100,000（20,000 行 x 5 列） | 約 1,390 ms | 約 195 ms（構築 15 + 投入 180） | **約 7 倍高速** |
| 200,000（50,000 行 x 4 列） | 約 3,200 ms | 約 380 ms | **約 8.4 倍高速** |

`setRangeData()` は、言語側のブリッジ境界呼び出しオーバーヘッドを最小化し、C++ エンジン側で直接セル配置および SST（共有文字列テーブル）登録を並列・一括処理するため、セル数が増えるほど圧倒的な差が生まれます。

### 使いどころ

* **CSV や DB 検索結果の流し込み** — 行・列の 2次元配列（またはタプルのリスト）をそのまま `ws.setRangeData("A1", matrix)` に渡すだけで完了します。
* **計算列の一括生成** — 行ごとに異なる相対参照の計算式を `ws.setRangeFormulas("D2", [["B2*C2"], ["B3*C3"], ...])` で一度に入力できます。
* **セル数が数百件を超える場合** — 原則として `setRangeData()` / `setRangeFormulas()` の利用を強く推奨します。

## 注意点

* **開始セル（左上）を指定します。** 第1引数に `"A1"` や `"B5"` などのセルアドレスを指定すると、そこを基点として 2次元配列の各要素が行・列方向に展開されます。
* **数値文字列は自動的に数値セルとして判定されます。** `"100"` や `"-12.5"` などの数値として解釈可能な文字列は、Excel 内部でも数値型（Number）セルとして保存されるため、集計関数（SUM 等）の対象になります。
* **数式には先頭の `=` を付けません。** `setRangeFormulas("D2", [["B2*C2"]])` のように記述します。

## 関連

* [目次](index.md)
* 前: [行・列・セル範囲の操作](general-05-rows-cols.md) / 次: [オートフィルターと並べ替え](general-07-autofilter-sort.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
