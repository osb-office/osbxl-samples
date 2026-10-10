# Bulk Input — setValueArray / setFormulaArray

level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](/ja/general-06-bulk-input.md)

## What you want to do

Write values and formulas into thousands or tens of thousands of cells **in one shot as an array** instead of calling `getRange()` per cell. Also, bulk-input formulas whose relative references shift row by row.

| Approach | What to use |
| --- | --- |
| Per-cell input | `ws.getCells(r, c).setNumberValue(...)` |
| **Bulk value input** | `ws.setValueArray(Array<dto.InputValueObject>)` |
| **Bulk formula input** | `ws.setFormulaArray(Array<dto.InputFormulaObject>)` |

## Classes covered

### `dto.InputValueObject`

| Method | Description |
| --- | --- |
| `setStringValue(A1C1, val, numberFormat?)` | Sets a string value. |
| `setNumberValue(A1C1, val, forceString?, numberFormat?)` | Sets a number. |
| `setBooleanValue(A1C1, val, forceString?, numberFormat?)` | Sets a boolean. |
| `setDateTimeValue(A1C1, val, forceString?, numberFormat?)` | Sets a date/time value. |
| `setEmptyVal(A1C1)` | Clears the cell range. |
| `getA1C1()` / `getInputType()` / `getStringValue()` / `getNumberValue()` / `getBooleanValue()` / `getDateTimeValue()` / `getNumberFormat()` / `getForceString()` | Read back the configured content. |

### `dto.InputFormulaObject`

| Method | Description |
| --- | --- |
| `setA1C1(A1C1)` | Sets the target cell. |
| `setFormula(formula)` | Sets the formula (no leading `=`). |
| `getA1C1()` / `getFormula()` | Read back the configured content. |

## Code

Output:

#### [Node.js]

```javascript
// bulk-input.js — bulk input of values and formulas, speed comparison against per-cell input
const { nodeosbxl, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. write mixed types in one shot ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);
const values = [];
const addString = (a1, val) => {
const o = new dto.InputValueObject();
o.setStringValue(a1, val);
values.push(o);
};
// Trailing optional arguments can be passed as undefined (treated as omitted)
const addNumber = (a1, val, fmt) => {
const o = new dto.InputValueObject();
o.setNumberValue(a1, val, false, fmt);
values.push(o);
};
addString("A1", "Product");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// Date: omitting the format auto-sets a default format (same as Range.setDateValue)
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd"); // explicit format
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date); // omitted → auto
values.push(dateVal2);
ws.setValueArray(values);
console.log("[1] bulk input with mixed types");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: display=${JSON.stringify(r.getValue()[a1])} format=${JSON.stringify(r.getNumberFormat()[a1])}`);
}
// ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] after setEmptyVal, A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));
// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for (let r = 1; r <= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r \* 10);
const formulas = [];
for (let r = 1; r <= 3; r++) {
const f = new dto.InputFormulaObject();
f.setA1C1(app.convertFromRowColNumber(r, 7)); // G1, G2, G3
f.setFormula(`F${r}\*2`); // different relative reference per row
formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] after setFormulaArray");
console.log(" formulas =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log(" values =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));
// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
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
console.log(`[4] writing ${ROWS \* COLS} cells`);
console.log(` per-cell getCells().setNumberValue() : ${msCell} ms`);
console.log(` setValueArray (build ${msBuild} + write ${msApply}) : ${msBuild + msApply} ms`);
const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[check] both methods agree: E2000 = ${a === b ? a : `mismatch ${a} / ${b}`}`);
wb.save();
wb.close();
```

#### [Python]

```python
# bulk-input.py — same verification as the Node.js sample (pyosbxl)
import json
import time
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("bulk-input-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- 1. write mixed types in one shot ----------
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 3, 20)
values = []
def add\_string(a1, val):
o = pyosbxl.dto.InputValueObject()
o.setStringValue(a1, val)
values.append(o)
# Trailing optional arguments are treated as omitted when not passed
def add\_number(a1, val, fmt=None):
o = pyosbxl.dto.InputValueObject()
if fmt is None:
o.setNumberValue(a1, val)
else:
o.setNumberValue(a1, val, False, fmt)
values.append(o)
add\_string("A1", "Product")
add\_number("B1", 1234.5)
add\_number("C1", 0.25, "0.0%")
# Date: omitting the format auto-sets a default format (same as Range.setDateValue)
date\_val = pyosbxl.dto.InputValueObject()
date\_val.setDateTimeValue("D1", date, False, "yyyy-mm-dd") # explicit format
values.append(date\_val)
date\_val2 = pyosbxl.dto.InputValueObject()
date\_val2.setDateTimeValue("D2", date) # omitted → auto
values.append(date\_val2)
ws.setValueArray(values)
print("[1] bulk input with mixed types")
for a1 in ["A1", "B1", "C1", "D1", "D2"]:
r = ws.getRange(f"{a1}:{a1}")
v = json.dumps(r.getValue()[a1], ensure\_ascii=False)
fmt = json.dumps(r.getNumberFormat()[a1], ensure\_ascii=False)
print(f" {a1}: display={v} format={fmt}")
# ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry")
empty = pyosbxl.dto.InputValueObject()
empty.setEmptyVal("A2")
ws.setValueArray([empty])
print("[2] after setEmptyVal, A2 =", json.dumps(ws.getRange("A2:A2").getValue()["A2"], ensure\_ascii=False))
# ---------- 3. bulk-input formulas with relative references shifted per row ----------
# Range.setFormula() does not expand relative references, so use this instead
for r in range(1, 4):
ws.getRange(f"F{r}:F{r}").setNumberValue(r \* 10)
formulas = []
for r in range(1, 4):
f = pyosbxl.dto.InputFormulaObject()
f.setA1C1(app.convertFromRowColNumber(r, 7)) # G1, G2, G3
f.setFormula(f"F{r}\*2") # different relative reference per row (no leading "=")
formulas.append(f)
ws.setFormulaArray(formulas)
print("[3] after setFormulaArray")
fm = ws.getRange("G1:G3").getFormula()
vm = ws.getRange("G1:G3").getValue(True)
keys = ["G1", "G2", "G3"]
print(" formulas =", json.dumps({k: fm[k] for k in keys}, ensure\_ascii=False, separators=(",", ":")))
print(" values =", json.dumps({k: vm[k] for k in keys}, ensure\_ascii=False, separators=(",", ":")))
# ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
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
print(f"[4] writing {ROWS \* COLS} cells")
print(f" per-cell getCells().setNumberValue() : {ms\_cell} ms")
print(f" setValueArray (build {ms\_build} + write {ms\_apply}) : {ms\_build + ms\_apply} ms")
a = ws\_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws\_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[check] both methods agree: E2000 = {a if a == b else f'mismatch {a} / {b}'}")
wb.save()
wb.close()
```

#### [Java]

```java
// bulk-input.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
public class Main {
// Quoted string, equivalent to JSON.stringify
static String q(String s) { return "\"" + s + "\""; }
static void addString(List<InputValueObjectWrapper> values, String a1, String val) {
InputValueObjectWrapper o = new InputValueObjectWrapper();
o.setStringValue(a1, val);
values.add(o);
}
// If fmt == null, do not pass the optional arguments (treated as omitted)
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
// ---------- 1. write mixed types in one shot ----------
DateTimeObjectWrapper date = new DateTimeObjectWrapper();
date.setYMD(2024, 3, 20);
List<InputValueObjectWrapper> values = new ArrayList<>();
addString(values, "A1", "Product");
addNumber(values, "B1", 1234.5, null);
addNumber(values, "C1", 0.25, "0.0%");
// Date: omitting the format auto-sets a default format (same as Range.setDateValue)
InputValueObjectWrapper dateVal = new InputValueObjectWrapper();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd"); // explicit format
values.add(dateVal);
InputValueObjectWrapper dateVal2 = new InputValueObjectWrapper();
dateVal2.setDateTimeValue("D2", date); // omitted → auto
values.add(dateVal2);
ws.setValueArray(values);
System.out.println("[1] bulk input with mixed types");
for (String a1 : new String[]{"A1", "B1", "C1", "D1", "D2"}) {
RangeWrapper r = ws.getRange(a1 + ":" + a1);
System.out.println(" " + a1 + ": display=" + q(r.getValue().get(a1))
+ " format=" + q(r.getNumberFormat().get(a1)));
}
// ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry");
InputValueObjectWrapper empty = new InputValueObjectWrapper();
empty.setEmptyVal("A2");
List<InputValueObjectWrapper> emptyList = new ArrayList<>();
emptyList.add(empty);
ws.setValueArray(emptyList);
System.out.println("[2] after setEmptyVal, A2 = " + q(ws.getRange("A2:A2").getValue().get("A2")));
// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for (int r = 1; r <= 3; r++) ws.getRange("F" + r + ":F" + r).setNumberValue(r \* 10);
List<InputFormulaObjectWrapper> formulas = new ArrayList<>();
for (int r = 1; r <= 3; r++) {
InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
f.setA1C1(app.convertFromRowColNumber(r, 7)); // G1, G2, G3
f.setFormula("F" + r + "\*2"); // different relative reference per row (no leading "=")
formulas.add(f);
}
ws.setFormulaArray(formulas);
System.out.println("[3] after setFormulaArray");
Map<String, String> fm = ws.getRange("G1:G3").getFormula();
Map<String, String> vm = ws.getRange("G1:G3").getValue(true);
System.out.println(" formulas = {\"G1\":" + q(fm.get("G1")) + ",\"G2\":" + q(fm.get("G2"))
+ ",\"G3\":" + q(fm.get("G3")) + "}");
System.out.println(" values = {\"G1\":" + q(vm.get("G1")) + ",\"G2\":" + q(vm.get("G2"))
+ ",\"G3\":" + q(vm.get("G3")) + "}");
// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
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
System.out.println("[4] writing " + (ROWS \* COLS) + " cells");
System.out.println(" per-cell getCells().setNumberValue() : " + msCell + " ms");
System.out.println(" setValueArray (build " + msBuild + " + write " + msApply + ") : "
+ (msBuild + msApply) + " ms");
String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
System.out.println("[check] both methods agree: E2000 = " + (a.equals(b) ? a : "mismatch " + a + " / " + b));
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// bulk-input.md — bulk input of values and formulas, speed comparison against per-cell input (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("bulk-input-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// JSON.stringify-compatible output (non-ASCII kept as-is)
$jq = fn($x): string => json\_encode($x, JSON\_UNESCAPED\_UNICODE | JSON\_UNESCAPED\_SLASHES);
$now = fn(): int => (int)(hrtime(true) / 1000000); // milliseconds
// ---------- 1. write mixed types in one shot ----------
$date = new \phposbxl\dto\DateTimeObject();
$date->setYMD(2024, 3, 20);
$values = [];
$addString = function (string $a1, string $val) use (&$values): void {
$o = new \phposbxl\dto\InputValueObject();
$o->setStringValue($a1, $val);
$values[] = $o;
};
// When an optional argument is omitted, call with fewer arguments (treated as omitted)
$addNumber = function (string $a1, float $val, ?string $fmt = null) use (&$values): void {
$o = new \phposbxl\dto\InputValueObject();
if ($fmt === null) {
$o->setNumberValue($a1, $val);
} else {
$o->setNumberValue($a1, $val, false, $fmt);
}
$values[] = $o;
};
$addString("A1", "Product");
$addNumber("B1", 1234.5);
$addNumber("C1", 0.25, "0.0%");
// Date: omitting the format auto-sets a default format (same as Range.setDateValue)
$dateVal = new \phposbxl\dto\InputValueObject();
$dateVal->setDateTimeValue("D1", $date, false, "yyyy-mm-dd"); // explicit format
$values[] = $dateVal;
$dateVal2 = new \phposbxl\dto\InputValueObject();
$dateVal2->setDateTimeValue("D2", $date); // omitted → auto
$values[] = $dateVal2;
$ws->setValueArray($values);
echo "[1] bulk input with mixed types\n";
foreach (["A1", "B1", "C1", "D1", "D2"] as $a1) {
$r = $ws->getRange("{$a1}:{$a1}");
echo " {$a1}: display=", $jq($r->getValue()[$a1]), " format=", $jq($r->getNumberFormat()[$a1]), "\n";
}
// ---------- 2. clearing cells is bulk input too ----------
$ws->getRange("A2:A2")->setValue("temporary entry");
$empty = new \phposbxl\dto\InputValueObject();
$empty->setEmptyVal("A2");
$ws->setValueArray([$empty]);
echo "[2] after setEmptyVal, A2 = ", $jq($ws->getRange("A2:A2")->getValue()["A2"]), "\n";
// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for ($r = 1; $r <= 3; $r++) {
$ws->getRange("F{$r}:F{$r}")->setNumberValue($r \* 10);
}
$formulas = [];
for ($r = 1; $r <= 3; $r++) {
$f = new \phposbxl\dto\InputFormulaObject();
$f->setA1C1($app->convertFromRowColNumber($r, 7)); // G1, G2, G3
$f->setFormula("F{$r}\*2"); // different relative reference per row (no leading "=")
$formulas[] = $f;
}
$ws->setFormulaArray($formulas);
echo "[3] after setFormulaArray\n";
echo " formulas = ", $jq($ws->getRange("G1:G3")->getFormula()), "\n";
echo " values = ", $jq($ws->getRange("G1:G3")->getValue(true)), "\n";
// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
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
echo "[4] writing ", $ROWS \* $COLS, " cells\n";
echo " per-cell getCells().setNumberValue() : {$msCell} ms\n";
echo " setValueArray (build {$msBuild} + write {$msApply}) : ", $msBuild + $msApply, " ms\n";
$a = $wsCell->getRange("E2000:E2000")->getValue(true)["E2000"];
$b = $wsBulk->getRange("E2000:E2000")->getValue(true)["E2000"];
echo "[check] both methods agree: E2000 = ", $a === $b ? $a : "mismatch {$a} / {$b}", "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// bulk-input.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("bulk-input-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. write mixed types in one shot ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);
const values = [];
const addString = (a1, val) => {
const o = new dto.InputValueObject();
o.setStringValue(a1, val);
values.push(o);
};
// The wasm build registers a separate overload per argument count; unlike node,
// passing undefined for a trailing optional argument throws a "Cannot pass non-string"
// exception. To omit an argument, call the overload with fewer arguments.
const addNumber = (a1, val, fmt) => {
const o = new dto.InputValueObject();
if (fmt === undefined) o.setNumberValue(a1, val);
else o.setNumberValue(a1, val, false, fmt);
values.push(o);
};
addString("A1", "Product");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// Date: omitting the format auto-sets a default format (same as Range.setDateValue)
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd"); // explicit format
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date); // omitted → auto
values.push(dateVal2);
ws.setValueArray(values);
console.log("[1] bulk input with mixed types");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: display=${JSON.stringify(r.getValue()[a1])} format=${JSON.stringify(r.getNumberFormat()[a1])}`);
}
// ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] after setEmptyVal, A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));
// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for (let r = 1; r <= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r \* 10);
const formulas = [];
for (let r = 1; r <= 3; r++) {
const f = new dto.InputFormulaObject();
f.setA1C1(app.convertFromRowColNumber(r, 7)); // G1, G2, G3
f.setFormula(`F${r}\*2`); // different relative reference per row
formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] after setFormulaArray");
console.log(" formulas =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log(" values =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));
// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
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
console.log(`[4] writing ${ROWS \* COLS} cells`);
console.log(` per-cell getCells().setNumberValue() : ${msCell} ms`);
console.log(` setValueArray (build ${msBuild} + write ${msApply}) : ${msBuild + msApply} ms`);
const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[check] both methods agree: E2000 = ${a === b ? a : `mismatch ${a} / ${b}`}`);
wb.save();
wb.close();
```

```plaintext
[1] bulk input with mixed types
A1: display="Product" format="General"
B1: display="1234.5" format="General"
C1: display="25.0%" format="0.0%"
D1: display="2024-03-20" format="yyyy-mm-dd"
D2: display="2024/3/20" format="yyyy/m/d;@"
[2] after setEmptyVal, A2 = ""
[3] after setFormulaArray
formulas = {"G1":"F1\*2","G2":"F2\*2","G3":"F3\*2"}
values = {"G1":"20","G2":"40","G3":"60"}
[4] writing 10000 cells
per-cell getCells().setNumberValue() : 102 ms
setValueArray (build 18 + write 58) : 76 ms
[check] both methods agree: E2000 = 200005
```

The ms values in `[4]` change with every run (the above is one example run).

## Notes

### Speed

| Cells | Per-cell `getCells().setNumberValue()` | `setValueArray()` | Difference |
| --- | --- | --- | --- |
| 10,000 (2,000 rows x 5 cols) | ~110 ms | 76 ms (build 18 + write 58) | ~1.4x |
| 100,000 (20,000 rows x 5 cols) | 1,390 ms | 770 ms (build 130 + write 640) | ~1.8x |

Both rows are single measured runs on osbxl 1.5.0 / Node.js v21; results vary with the environment and between runs. As a trend, **the gap widens as the cell count grows**. \*\*The time spent assembling `InputValueObject`s on the JavaScript side is not negligible either\*\*, which is why the measurements above separate "build" from "write". If you need to write tens of thousands of cells or more, use `setValueArray()`.

### When to use which

* **`setValueArray()`** — when you have "values to insert in bulk", such as an array, CSV, or DB query results.  
  You can mix types (string / number / boolean / date) in a single array.
* **`setFormulaArray()`** — when inserting **formulas with different relative references per row**. `Range.setFormula()` does not expand relative references (see [overview caveats](general-01-overview.md#notes)), so this is the right tool for building calculated columns.
* For a few dozen cells, `getCells(r, c)` / `getRange("A1:A1")` is perfectly fine — favor readability.

## Caveats

* **Plain object literals are not accepted.** You must create instances of `dto.InputValueObject` / `dto.InputFormulaObject`.

  ```javascript
  // ✗ throws: Invalid argument
  ws.setValueArray([{ cell: "A1", value: 100 }]);
  // ✓
  const o = new dto.InputValueObject();
  o.setNumberValue("A1", 100);
  ws.setValueArray([o]);
  ```
* **`setDateTimeValue()` auto-sets the display format** (symmetric with `Range.setDateValue()`). Omitting `numberFormat` automatically applies `yyyy/m/d;@` (date only) / `h:mm:ss;@` (time only) / `yyyy/m/d h:mm:ss;@` (date + time) depending on the content.  
  Pass the 4th argument only when you want a specific format.

  ```plaintext
  setDateTimeValue("D2", date) → display "2024/3/20" format "yyyy/m/d;@"
  setDateTimeValue("D1", date, false, "yyyy-mm-dd") → display "2024-03-20" format "yyyy-mm-dd"
  ```
* **When the same cell is specified multiple times, the element later in the array wins** (measured). To avoid unintended overwrites, make sure there are no duplicate A1C1 values before writing.
* **`InputFormulaObject` requires `setA1C1()`.** Setting only the formula has no effect if the target cell is not specified.
* **No leading `=` in `setFormulaArray()` formulas either.** Use `f.setFormula("F1*2")`.
* \*\*`undefined` may only be passed for \*trailing\* optional arguments\*\* (treated as omitted). `undefined` or `null` in a middle position throws an `unsupported arguments` exception.

  ```javascript
  // ✓ trailing undefined = omitted
  o.setNumberValue("A1", 100, false, undefined);
  // ✗ throws: undefined in the middle
  o.setNumberValue("A1", undefined, false, "#,##0");
  // ✗ throws: null is not treated as omitted
  o.setNumberValue("A1", 100, null);
  ```

  This is a common specification across osbxl (including the `Range` setters).

## See also

* [Index](index.md)
* Previous: [Rows, columns and ranges](general-05-rows-cols.md) / Next: [AutoFilter and Sorting](general-07-autofilter-sort.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
