# 大量データの一括投入 — setValueArray / setFormulaArray

level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-06-bulk-input.md)

## やりたいこと

何千・何万セルもの値や数式を、セルごとに `getRange()` するのではなく **配列にまとめて一度に投入する**。あわせて相対参照を行ごとにずらした数式を一括入力する。

| やり方 | 使うもの |
| --- | --- |
| セルごとに投入 | `ws.getCells(r, c).setNumberValue(...)` |
| **値を一括投入** | `ws.setValueArray(Array<dto.InputValueObject>)` |
| **数式を一括投入** | `ws.setFormulaArray(Array<dto.InputFormulaObject>)` |

## 対象クラス

### `dto.InputValueObject`

| メソッド | 説明 |
| --- | --- |
| `setStringValue(A1C1, val, numberFormat?)` | 文字列値を設定します。 |
| `setNumberValue(A1C1, val, forceString?, numberFormat?)` | 数値を設定します。 |
| `setBooleanValue(A1C1, val, forceString?, numberFormat?)` | 真偽値を設定します。 |
| `setDateTimeValue(A1C1, val, forceString?, numberFormat?)` | 日付時刻値を設定します。 |
| `setEmptyVal(A1C1)` | セル範囲を空に設定します。 |
| `getA1C1()` / `getInputType()` / `getStringValue()` / `getNumberValue()` / `getBooleanValue()` / `getDateTimeValue()` / `getNumberFormat()` / `getForceString()` | 設定内容の取得。 |

### `dto.InputFormulaObject`

| メソッド | 説明 |
| --- | --- |
| `setA1C1(A1C1)` | 対象セルを設定します。 |
| `setFormula(formula)` | 数式を設定します（先頭に `=` は付けない）。 |
| `getA1C1()` / `getFormula()` | 設定内容の取得。 |

## コード

実行結果:

#### [Node.js]

```javascript
// bulk-input.js — 値と数式の一括投入、セルごと投入との速度比較
const { nodeosbxl, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. いろいろな型を一度に投入する ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);
const values = [];
const addString = (a1, val) => {
const o = new dto.InputValueObject();
o.setStringValue(a1, val);
values.push(o);
};
// 末尾の省略可能引数は undefined のまま渡せます（省略として扱われます）
const addNumber = (a1, val, fmt) => {
const o = new dto.InputValueObject();
o.setNumberValue(a1, val, false, fmt);
values.push(o);
};
addString("A1", "商品名");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd"); // 書式を明示
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date); // 省略 → 自動
values.push(dateVal2);
ws.setValueArray(values);
console.log("[1] 型混在の一括投入");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])} 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
}
// ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] setEmptyVal 後 A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));
// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for (let r = 1; r <= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r \* 10);
const formulas = [];
for (let r = 1; r <= 3; r++) {
const f = new dto.InputFormulaObject();
f.setA1C1(app.convertFromRowColNumber(r, 7)); // G1, G2, G3
f.setFormula(`F${r}\*2`); // 行ごとに異なる相対参照
formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] setFormulaArray 後");
console.log(" 数式 =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log(" 値 =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));
// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
const ROWS = 2000, COLS = 5;
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r \* 100 + c);
}
const msCell = Date.now() - t;
const wsBulk = wb.addWorkSheet("Bulk");
const bulk = [];
t = Date.now();
for (let r = 1; r <= ROWS; r++) {
for (let c = 1; c <= COLS; c++) {
const o = new dto.InputValueObject();
o.setNumberValue(app.convertFromRowColNumber(r, c), r \* 100 + c);
bulk.push(o);
}
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setValueArray(bulk);
const msApply = Date.now() - t;
console.log(`[4] ${ROWS \* COLS} セルの投入`);
console.log(` セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(` setValueArray (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms`);
const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === b ? a : `不一致 ${a} / ${b}`}`);
wb.save();
wb.close();
```

#### [Python]

```python
# bulk-input.py — node 版サンプルと同一検証 (pyosbxl)
import json
import time
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("bulk-input-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- 1. いろいろな型を一度に投入する ----------
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 3, 20)
values = []
def add\_string(a1, val):
o = pyosbxl.dto.InputValueObject()
o.setStringValue(a1, val)
values.append(o)
# 末尾の省略可能引数は渡さなければ省略として扱われます
def add\_number(a1, val, fmt=None):
o = pyosbxl.dto.InputValueObject()
if fmt is None:
o.setNumberValue(a1, val)
else:
o.setNumberValue(a1, val, False, fmt)
values.append(o)
add\_string("A1", "商品名")
add\_number("B1", 1234.5)
add\_number("C1", 0.25, "0.0%")
# 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
date\_val = pyosbxl.dto.InputValueObject()
date\_val.setDateTimeValue("D1", date, False, "yyyy-mm-dd") # 書式を明示
values.append(date\_val)
date\_val2 = pyosbxl.dto.InputValueObject()
date\_val2.setDateTimeValue("D2", date) # 省略 → 自動
values.append(date\_val2)
ws.setValueArray(values)
print("[1] 型混在の一括投入")
for a1 in ["A1", "B1", "C1", "D1", "D2"]:
r = ws.getRange(f"{a1}:{a1}")
v = json.dumps(r.getValue()[a1], ensure\_ascii=False)
fmt = json.dumps(r.getNumberFormat()[a1], ensure\_ascii=False)
print(f" {a1}: 表示値={v} 書式={fmt}")
# ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力")
empty = pyosbxl.dto.InputValueObject()
empty.setEmptyVal("A2")
ws.setValueArray([empty])
print("[2] setEmptyVal 後 A2 =", json.dumps(ws.getRange("A2:A2").getValue()["A2"], ensure\_ascii=False))
# ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
# Range.setFormula() では相対参照が展開されないので、こちらを使う
for r in range(1, 4):
ws.getRange(f"F{r}:F{r}").setNumberValue(r \* 10)
formulas = []
for r in range(1, 4):
f = pyosbxl.dto.InputFormulaObject()
f.setA1C1(app.convertFromRowColNumber(r, 7)) # G1, G2, G3
f.setFormula(f"F{r}\*2") # 行ごとに異なる相対参照（先頭の "=" は付けない）
formulas.append(f)
ws.setFormulaArray(formulas)
print("[3] setFormulaArray 後")
fm = ws.getRange("G1:G3").getFormula()
vm = ws.getRange("G1:G3").getValue(True)
keys = ["G1", "G2", "G3"]
print(" 数式 =", json.dumps({k: fm[k] for k in keys}, ensure\_ascii=False, separators=(",", ":")))
print(" 値 =", json.dumps({k: vm[k] for k in keys}, ensure\_ascii=False, separators=(",", ":")))
# ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
ROWS, COLS = 2000, 5
ws\_cell = wb.addWorkSheet("CellByCell")
t = time.perf\_counter()
for r in range(1, ROWS + 1):
for c in range(1, COLS + 1):
ws\_cell.getCells(r, c).setNumberValue(r \* 100 + c)
ms\_cell = int((time.perf\_counter() - t) \* 1000)
ws\_bulk = wb.addWorkSheet("Bulk")
bulk = []
t = time.perf\_counter()
for r in range(1, ROWS + 1):
for c in range(1, COLS + 1):
o = pyosbxl.dto.InputValueObject()
o.setNumberValue(app.convertFromRowColNumber(r, c), r \* 100 + c)
bulk.append(o)
ms\_build = int((time.perf\_counter() - t) \* 1000)
t = time.perf\_counter()
ws\_bulk.setValueArray(bulk)
ms\_apply = int((time.perf\_counter() - t) \* 1000)
print(f"[4] {ROWS \* COLS} セルの投入")
print(f" セルごと getCells().setNumberValue() : {ms\_cell} ms")
print(f" setValueArray (構築{ms\_build} + 投入{ms\_apply}) : {ms\_build + ms\_apply} ms")
a = ws\_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws\_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[検証] 両方式とも E2000 = {a if a == b else f'不一致 {a} / {b}'}")
wb.save()
wb.close()
```

#### [Java]

```java
// bulk-input.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
public class Main {
// JSON.stringify 相当の引用符付き文字列
static String q(String s) { return "\"" + s + "\""; }
static void addString(List<InputValueObjectWrapper> values, String a1, String val) {
InputValueObjectWrapper o = new InputValueObjectWrapper();
o.setStringValue(a1, val);
values.add(o);
}
// fmt == null なら省略可能引数を渡さない（省略として扱われる）
static void addNumber(List<InputValueObjectWrapper> values, String a1, double val, String fmt) {
InputValueObjectWrapper o = new InputValueObjectWrapper();
if (fmt == null) {
o.setNumberValue(a1, val);
} else {
o.setNumberValue(a1, val, false, fmt);
}
values.add(o);
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("bulk-input-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- 1. いろいろな型を一度に投入する ----------
DateTimeObjectWrapper date = new DateTimeObjectWrapper();
date.setYMD(2024, 3, 20);
List<InputValueObjectWrapper> values = new ArrayList<>();
addString(values, "A1", "商品名");
addNumber(values, "B1", 1234.5, null);
addNumber(values, "C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
InputValueObjectWrapper dateVal = new InputValueObjectWrapper();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd"); // 書式を明示
values.add(dateVal);
InputValueObjectWrapper dateVal2 = new InputValueObjectWrapper();
dateVal2.setDateTimeValue("D2", date); // 省略 → 自動
values.add(dateVal2);
ws.setValueArray(values);
System.out.println("[1] 型混在の一括投入");
for (String a1 : new String[]{"A1", "B1", "C1", "D1", "D2"}) {
RangeWrapper r = ws.getRange(a1 + ":" + a1);
System.out.println(" " + a1 + ": 表示値=" + q(r.getValue().get(a1))
+ " 書式=" + q(r.getNumberFormat().get(a1)));
}
// ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力");
InputValueObjectWrapper empty = new InputValueObjectWrapper();
empty.setEmptyVal("A2");
List<InputValueObjectWrapper> emptyList = new ArrayList<>();
emptyList.add(empty);
ws.setValueArray(emptyList);
System.out.println("[2] setEmptyVal 後 A2 = " + q(ws.getRange("A2:A2").getValue().get("A2")));
// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for (int r = 1; r <= 3; r++) ws.getRange("F" + r + ":F" + r).setNumberValue(r \* 10);
List<InputFormulaObjectWrapper> formulas = new ArrayList<>();
for (int r = 1; r <= 3; r++) {
InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
f.setA1C1(app.convertFromRowColNumber(r, 7)); // G1, G2, G3
f.setFormula("F" + r + "\*2"); // 行ごとに異なる相対参照（先頭の "=" は付けない）
formulas.add(f);
}
ws.setFormulaArray(formulas);
System.out.println("[3] setFormulaArray 後");
Map<String, String> fm = ws.getRange("G1:G3").getFormula();
Map<String, String> vm = ws.getRange("G1:G3").getValue(true);
System.out.println(" 数式 = {\"G1\":" + q(fm.get("G1")) + ",\"G2\":" + q(fm.get("G2"))
+ ",\"G3\":" + q(fm.get("G3")) + "}");
System.out.println(" 値 = {\"G1\":" + q(vm.get("G1")) + ",\"G2\":" + q(vm.get("G2"))
+ ",\"G3\":" + q(vm.get("G3")) + "}");
// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
final int ROWS = 2000, COLS = 5;
WorkSheetWrapper wsCell = wb.addWorkSheet("CellByCell");
long t = System.currentTimeMillis();
for (int r = 1; r <= ROWS; r++) {
for (int c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r \* 100 + c);
}
long msCell = System.currentTimeMillis() - t;
WorkSheetWrapper wsBulk = wb.addWorkSheet("Bulk");
List<InputValueObjectWrapper> bulk = new ArrayList<>();
t = System.currentTimeMillis();
for (int r = 1; r <= ROWS; r++) {
for (int c = 1; c <= COLS; c++) {
InputValueObjectWrapper o = new InputValueObjectWrapper();
o.setNumberValue(app.convertFromRowColNumber(r, c), r \* 100 + c);
bulk.add(o);
}
}
long msBuild = System.currentTimeMillis() - t;
t = System.currentTimeMillis();
wsBulk.setValueArray(bulk);
long msApply = System.currentTimeMillis() - t;
System.out.println("[4] " + (ROWS \* COLS) + " セルの投入");
System.out.println(" セルごと getCells().setNumberValue() : " + msCell + " ms");
System.out.println(" setValueArray (構築" + msBuild + " + 投入" + msApply + ") : "
+ (msBuild + msApply) + " ms");
String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
System.out.println("[検証] 両方式とも E2000 = " + (a.equals(b) ? a : "不一致 " + a + " / " + b));
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// bulk-input.md — 値と数式の一括投入、セルごと投入との速度比較 (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("bulk-input-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// JSON.stringify 相当の出力（非 ASCII はそのまま）
$jq = fn($x): string => json\_encode($x, JSON\_UNESCAPED\_UNICODE | JSON\_UNESCAPED\_SLASHES);
$now = fn(): int => (int)(hrtime(true) / 1000000); // ミリ秒
// ---------- 1. いろいろな型を一度に投入する ----------
$date = new \phposbxl\dto\DateTimeObject();
$date->setYMD(2024, 3, 20);
$values = [];
$addString = function (string $a1, string $val) use (&$values): void {
$o = new \phposbxl\dto\InputValueObject();
$o->setStringValue($a1, $val);
$values[] = $o;
};
// 省略可能引数は null のとき引数の少ない呼び出しを使う（省略として扱われる）
$addNumber = function (string $a1, float $val, ?string $fmt = null) use (&$values): void {
$o = new \phposbxl\dto\InputValueObject();
if ($fmt === null) {
$o->setNumberValue($a1, $val);
} else {
$o->setNumberValue($a1, $val, false, $fmt);
}
$values[] = $o;
};
$addString("A1", "商品名");
$addNumber("B1", 1234.5);
$addNumber("C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
$dateVal = new \phposbxl\dto\InputValueObject();
$dateVal->setDateTimeValue("D1", $date, false, "yyyy-mm-dd"); // 書式を明示
$values[] = $dateVal;
$dateVal2 = new \phposbxl\dto\InputValueObject();
$dateVal2->setDateTimeValue("D2", $date); // 省略 → 自動
$values[] = $dateVal2;
$ws->setValueArray($values);
echo "[1] 型混在の一括投入\n";
foreach (["A1", "B1", "C1", "D1", "D2"] as $a1) {
$r = $ws->getRange("{$a1}:{$a1}");
echo " {$a1}: 表示値=", $jq($r->getValue()[$a1]), " 書式=", $jq($r->getNumberFormat()[$a1]), "\n";
}
// ---------- 2. 空にするのも一括投入 ----------
$ws->getRange("A2:A2")->setValue("いったん入力");
$empty = new \phposbxl\dto\InputValueObject();
$empty->setEmptyVal("A2");
$ws->setValueArray([$empty]);
echo "[2] setEmptyVal 後 A2 = ", $jq($ws->getRange("A2:A2")->getValue()["A2"]), "\n";
// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for ($r = 1; $r <= 3; $r++) {
$ws->getRange("F{$r}:F{$r}")->setNumberValue($r \* 10);
}
$formulas = [];
for ($r = 1; $r <= 3; $r++) {
$f = new \phposbxl\dto\InputFormulaObject();
$f->setA1C1($app->convertFromRowColNumber($r, 7)); // G1, G2, G3
$f->setFormula("F{$r}\*2"); // 行ごとに異なる相対参照（先頭の "=" は付けない）
$formulas[] = $f;
}
$ws->setFormulaArray($formulas);
echo "[3] setFormulaArray 後\n";
echo " 数式 = ", $jq($ws->getRange("G1:G3")->getFormula()), "\n";
echo " 値 = ", $jq($ws->getRange("G1:G3")->getValue(true)), "\n";
// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
$ROWS = 2000;
$COLS = 5;
$wsCell = $wb->addWorkSheet("CellByCell");
$t = $now();
for ($r = 1; $r <= $ROWS; $r++) {
for ($c = 1; $c <= $COLS; $c++) {
$wsCell->getCells($r, $c)->setNumberValue($r \* 100 + $c);
}
}
$msCell = $now() - $t;
$wsBulk = $wb->addWorkSheet("Bulk");
$bulk = [];
$t = $now();
for ($r = 1; $r <= $ROWS; $r++) {
for ($c = 1; $c <= $COLS; $c++) {
$o = new \phposbxl\dto\InputValueObject();
$o->setNumberValue($app->convertFromRowColNumber($r, $c), $r \* 100 + $c);
$bulk[] = $o;
}
}
$msBuild = $now() - $t;
$t = $now();
$wsBulk->setValueArray($bulk);
$msApply = $now() - $t;
echo "[4] ", $ROWS \* $COLS, " セルの投入\n";
echo " セルごと getCells().setNumberValue() : {$msCell} ms\n";
echo " setValueArray (構築{$msBuild} + 投入{$msApply}) : ", $msBuild + $msApply, " ms\n";
$a = $wsCell->getRange("E2000:E2000")->getValue(true)["E2000"];
$b = $wsBulk->getRange("E2000:E2000")->getValue(true)["E2000"];
echo "[検証] 両方式とも E2000 = ", $a === $b ? $a : "不一致 {$a} / {$b}", "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// bulk-input.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("bulk-input-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. いろいろな型を一度に投入する ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);
const values = [];
const addString = (a1, val) => {
const o = new dto.InputValueObject();
o.setStringValue(a1, val);
values.push(o);
};
// wasm 版は引数ごとに別オーバーロードが登録されており、node と違い
// 末尾の省略可能引数へ undefined を渡すと "Cannot pass non-string" 例外になる。
// 省略したいときは引数の少ないオーバーロードを呼ぶ。
const addNumber = (a1, val, fmt) => {
const o = new dto.InputValueObject();
if (fmt === undefined) o.setNumberValue(a1, val);
else o.setNumberValue(a1, val, false, fmt);
values.push(o);
};
addString("A1", "商品名");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd"); // 書式を明示
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date); // 省略 → 自動
values.push(dateVal2);
ws.setValueArray(values);
console.log("[1] 型混在の一括投入");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])} 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
}
// ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] setEmptyVal 後 A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));
// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for (let r = 1; r <= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r \* 10);
const formulas = [];
for (let r = 1; r <= 3; r++) {
const f = new dto.InputFormulaObject();
f.setA1C1(app.convertFromRowColNumber(r, 7)); // G1, G2, G3
f.setFormula(`F${r}\*2`); // 行ごとに異なる相対参照
formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] setFormulaArray 後");
console.log(" 数式 =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log(" 値 =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));
// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
const ROWS = 2000, COLS = 5;
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r \* 100 + c);
}
const msCell = Date.now() - t;
const wsBulk = wb.addWorkSheet("Bulk");
const bulk = [];
t = Date.now();
for (let r = 1; r <= ROWS; r++) {
for (let c = 1; c <= COLS; c++) {
const o = new dto.InputValueObject();
o.setNumberValue(app.convertFromRowColNumber(r, c), r \* 100 + c);
bulk.push(o);
}
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setValueArray(bulk);
const msApply = Date.now() - t;
console.log(`[4] ${ROWS \* COLS} セルの投入`);
console.log(` セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(` setValueArray (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms`);
const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === b ? a : `不一致 ${a} / ${b}`}`);
wb.save();
wb.close();
```

```plaintext
[1] 型混在の一括投入
A1: 表示値="商品名" 書式="General"
B1: 表示値="1234.5" 書式="General"
C1: 表示値="25.0%" 書式="0.0%"
D1: 表示値="2024-03-20" 書式="yyyy-mm-dd"
D2: 表示値="2024/3/20" 書式="yyyy/m/d;@"
[2] setEmptyVal 後 A2 = ""
[3] setFormulaArray 後
数式 = {"G1":"F1\*2","G2":"F2\*2","G3":"F3\*2"}
値 = {"G1":"20","G2":"40","G3":"60"}
[4] 10000 セルの投入
セルごと getCells().setNumberValue() : 102 ms
setValueArray (構築18 + 投入58) : 76 ms
[検証] 両方式とも E2000 = 200005
```

`[4]` の ms は実行のたびに変わります（上記は 1 回の実行例）。

## 解説

### 速度

| セル数 | セルごと `getCells().setNumberValue()` | `setValueArray()` | 差 |
| --- | --- | --- | --- |
| 10,000（2,000 行 x 5 列） | 約 110 ms | 76 ms（構築 18 + 投入 58） | 約 1.4 倍 |
| 100,000（20,000 行 x 5 列） | 1,390 ms | 770 ms（構築 130 + 投入 640） | 約 1.8 倍 |

いずれも osbxl 1.5.0 / Node.js v21 での 1 回の実測例で、環境と実行のたびに変わります。 傾向として**セル数が増えるほど差が開きます**。 **JavaScript 側で `InputValueObject` を組み立てる時間も無視できない**ので、 上記では「構築」と「投入」を分けて計測しています。 数万セルを超えるなら `setValueArray()` を使ってください。

### 使いどころ

* **`setValueArray()`** — 配列や CSV、DB の結果など「まとめて入れたい値」があるとき。  
   型（文字列 / 数値 / 真偽値 / 日付）を 1 つの配列に混在させられます。
* **`setFormulaArray()`** — **行ごとに異なる相対参照の数式**を入れるとき。  
   `Range.setFormula()` では相対参照が展開されない（[概要の注意点](general-01-overview.md#注意点)参照）ため、 計算列を作るならこちらが正解です。
* セル数が数十程度なら、可読性を優先して `getCells(r, c)` / `getRange("A1:A1")` で十分です。

## 注意点

* **素のオブジェクトリテラルは渡せません。** `dto.InputValueObject` / `dto.InputFormulaObject` の インスタンスを作る必要があります。

  ```javascript
  // ✗ 例外: Invalid argument
  ws.setValueArray([{ cell: "A1", value: 100 }]);
  // ✓
  const o = new dto.InputValueObject();
  o.setNumberValue("A1", 100);
  ws.setValueArray([o]);
  ```
* **`setDateTimeValue()` は表示書式を自動設定します**（`Range.setDateValue()` と対称）。  
   `numberFormat` を省略すると内容に応じて `yyyy/m/d;@`（日付のみ）/ `h:mm:ss;@`（時刻のみ）/ `yyyy/m/d h:mm:ss;@`（日付+時刻）が自動で付きます。  
   特定の書式にしたいときだけ第4引数で渡します。

  ```plaintext
  setDateTimeValue("D2", date) → 表示値 "2024/3/20" 書式 "yyyy/m/d;@"
  setDateTimeValue("D1", date, false, "yyyy-mm-dd") → 表示値 "2024-03-20" 書式 "yyyy-mm-dd"
  ```
* **同じセルを複数回指定すると、配列の後ろの要素が勝ちます**（実測）。  
   意図しない上書きが起きないよう、投入前に A1C1 の重複が無いことを確認してください。
* **`InputFormulaObject` は `setA1C1()` が必須。** 数式だけ設定してもセルが指定されていなければ反映されません。
* **`setFormulaArray()` の数式にも先頭の `=` は付けません。** `f.setFormula("F1*2")` です。
* **`undefined` は「末尾の」省略可能引数にだけ渡せます**（省略として扱われます）。  
   途中の位置の `undefined` や `null` は `unsupported arguments` 例外になります。

  ```javascript
  // ✓ 末尾 undefined = 省略
  o.setNumberValue("A1", 100, false, undefined);
  // ✗ 例外: 途中の undefined
  o.setNumberValue("A1", undefined, false, "#,##0");
  // ✗ 例外: null は省略として扱われない
  o.setNumberValue("A1", 100, null);
  ```

  これは osbxl 全体（`Range` の setter など含む）に共通の仕様です。

## 関連

* [目次](index.md)
* 前: [行・列・セル範囲の操作](general-05-rows-cols.md) / 次: [オートフィルターと並べ替え](general-07-autofilter-sort.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
