# High-Speed Bulk Input — setRangeData / setRangeFormulas

level: intermediate ／ verified: 2026-10-11 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-06-bulk-input.md)

## What you want to do

Write values and formulas into thousands or tens of thousands of cells **in one shot as a 2D matrix array** instead of calling `getCells()` per cell. Populate tabular data from CSV files or database query results into worksheets with maximum speed and minimal memory overhead.

| Approach | What to use | Characteristics |
| --- | --- | --- |
| Per-cell input | `ws.getCells(r, c).setNumberValue(...)` | Suitable for a few cells. Crossing language/C++ boundaries per cell incurs heavy overhead. |
| **2D array value input (Fastest, Recommended)** | `ws.setRangeData(startA1, data)` | **Fastest & Recommended.** Zero DTO creation overhead. Numeric strings are automatically saved as numeric cells. |
| **2D array formula input (Fastest, Recommended)** | `ws.setRangeFormulas(startA1, formulas)` | **Fastest & Recommended.** Stream formulas across calculation columns at once (no leading "=" needed). |

## Classes covered

### `WorkSheet`

| Method | Description |
| --- | --- |
| `setRangeData(startA1, data)` | Bulk enters 2D cell values (array of string rows/columns) from the top-left start cell (e.g., "A1"). |
| `setRangeFormulas(startA1, formulas)` | Bulk enters 2D formulas (array of formula string rows/columns, no "=") from the top-left start cell (e.g., "C1"). |

## Code

Output:

#### [Node.js]

```javascript
// bulk-input.js — high-speed bulk input of values and formulas with 2D array (nodeosbxl)
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. Bulk value input with 2D array (setRangeData) ----------
// Prepare a 2D array of rows and columns as strings and write all at once starting from top-left cell.
// Numeric strings like "100" are automatically stored as Excel numbers.
const data = [
    ["Product", "Price", "Qty"],
    ["Apple", "100", "3"],
    ["Orange", "50", "6"],
    ["Grape", "150", "2"],
];
ws.setRangeData("A1", data);

console.log("[1] After setRangeData:");
for (const a1 of ["A1", "B2", "C2"]) {
    console.log(`    ${a1} = ${JSON.stringify(ws.getRange(a1).getValue()[a1])}`);
}

// ---------- 2. Bulk formula input with 2D array (setRangeFormulas) ----------
// Bulk input formulas across calculation columns (no leading "=" needed).
const formulas = [
    ["Subtotal"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
ws.setRangeFormulas("D1", formulas);

console.log("[2] After setRangeFormulas:");
console.log("    Formulas D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getFormula()));
console.log("    Values   D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getValue(true)));

// ---------- 3. Performance comparison (2,000 rows x 5 cols = 10,000 cells) ----------
const ROWS = 2000, COLS = 5;

// Per-cell input via getCells()
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

// 2D array input via setRangeData() (Fastest)
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

console.log(`[3] Performance for ${ROWS * COLS} cells:`);
console.log(`    Per-cell getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setRangeData (build ${msBuild} + write ${msApply}) : ${msBuild + msApply} ms (fastest)`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[Verification] Both match E2000 = ${a === Number(b) ? b : `mismatch ${a} / ${b}`}`);

wb.save();
wb.close();
```

#### [Python]

```python
# bulk-input.py — high-speed bulk input of values and formulas with 2D array (pyosbxl)
import json
import time
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("bulk-input.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. Bulk value input with 2D array (setRangeData) ----------
data = [
    ["Product", "Price", "Qty"],
    ["Apple", "100", "3"],
    ["Orange", "50", "6"],
    ["Grape", "150", "2"],
]
ws.setRangeData("A1", data)

print("[1] After setRangeData:")
for a1 in ["A1", "B2", "C2"]:
    print(f"    {a1} = {json.dumps(ws.getRange(a1).getValue()[a1], ensure_ascii=False)}")

# ---------- 2. Bulk formula input with 2D array (setRangeFormulas) ----------
formulas = [
    ["Subtotal"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
]
ws.setRangeFormulas("D1", formulas)

print("[2] After setRangeFormulas:")
print("    Formulas D2:D4 =", json.dumps(ws.getRange("D2:D4").getFormula(), ensure_ascii=False))
print("    Values   D2:D4 =", json.dumps(ws.getRange("D2:D4").getValue(True), ensure_ascii=False))

# ---------- 3. Performance comparison (2,000 rows x 5 cols = 10,000 cells) ----------
ROWS, COLS = 2000, 5

# Per-cell input via getCells()
ws_cell = wb.addWorkSheet("CellByCell")
t = time.perf_counter()
for r in range(1, ROWS + 1):
    for c in range(1, COLS + 1):
        ws_cell.getCells(r, c).setNumberValue(r * 100 + c)
ms_cell = int((time.perf_counter() - t) * 1000)

# 2D array input via setRangeData() (Fastest)
ws_bulk = wb.addWorkSheet("Bulk")
t = time.perf_counter()
matrix = [[str(r * 100 + c) for c in range(1, COLS + 1)] for r in range(1, ROWS + 1)]
ms_build = int((time.perf_counter() - t) * 1000)
t = time.perf_counter()
ws_bulk.setRangeData("A1", matrix)
ms_apply = int((time.perf_counter() - t) * 1000)

print(f"[3] Performance for {ROWS * COLS} cells:")
print(f"    Per-cell getCells().setNumberValue() : {ms_cell} ms")
print(f"    setRangeData (build {ms_build} + write {ms_apply}) : {ms_build + ms_apply} ms (fastest)")

a = ws_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[Verification] Both match E2000 = {b if a == float(b) else f'mismatch {a} / {b}'}")

wb.save()
wb.close()
```

#### [Java]

```java
// bulk-input.java — high-speed bulk input of values and formulas with 2D array (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        App app = new App();
        WorkBookWrapper wb = app.createWorkBook("bulk-input.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. Bulk value input with 2D array (setRangeData) ----------
        List<List<String>> data = Arrays.asList(
            Arrays.asList("Product", "Price", "Qty"),
            Arrays.asList("Apple", "100", "3"),
            Arrays.asList("Orange", "50", "6"),
            Arrays.asList("Grape", "150", "2")
        );
        ws.setRangeData("A1", data);

        System.out.println("[1] After setRangeData:");
        for (String a1 : Arrays.asList("A1", "B2", "C2")) {
            System.out.println("    " + a1 + " = " + ws.getRange(a1).getValue().get(a1));
        }

        // ---------- 2. Bulk formula input with 2D array (setRangeFormulas) ----------
        List<List<String>> formulas = Arrays.asList(
            Collections.singletonList("Subtotal"),
            Collections.singletonList("B2*C2"),
            Collections.singletonList("B3*C3"),
            Collections.singletonList("B4*C4")
        );
        ws.setRangeFormulas("D1", formulas);

        System.out.println("[2] After setRangeFormulas:");
        System.out.println("    Formulas D2:D4 = " + ws.getRange("D2:D4").getFormula());
        System.out.println("    Values   D2:D4 = " + ws.getRange("D2:D4").getValue(true));

        // ---------- 3. Performance comparison (2,000 rows x 5 cols = 10,000 cells) ----------
        final int ROWS = 2000, COLS = 5;

        // Per-cell input via getCells()
        WorkSheetWrapper wsCell = wb.addWorkSheet("CellByCell");
        long t = System.currentTimeMillis();
        for (int r = 1; r <= ROWS; r++) {
            for (int c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
        }
        long msCell = System.currentTimeMillis() - t;

        // 2D array input via setRangeData() (Fastest)
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

        System.out.println("[3] Performance for " + (ROWS * COLS) + " cells:");
        System.out.println("    Per-cell getCells().setNumberValue() : " + msCell + " ms");
        System.out.println("    setRangeData (build " + msBuild + " + write " + msApply + ") : " + (msBuild + msApply) + " ms (fastest)");

        String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
        String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
        System.out.println("[Verification] Both match E2000 = " + (a.equals(b) ? b : "mismatch " + a + " / " + b));

        wb.save();
        wb.close();
    }
}
```

#### [PHP]

```php
// bulk-input.php — high-speed bulk input of values and formulas with 2D array (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("bulk-input.xlsx");
$ws = $wb->openWorkSheet("Sheet1");

$now = fn(): int => (int)(hrtime(true) / 1000000);   // ms

// ---------- 1. Bulk value input with 2D array (setRangeData) ----------
$data = [
    ["Product", "Price", "Qty"],
    ["Apple", "100", "3"],
    ["Orange", "50", "6"],
    ["Grape", "150", "2"],
];
$ws->setRangeData("A1", $data);

echo "[1] After setRangeData:\n";
foreach (["A1", "B2", "C2"] as $a1) {
    echo "    {$a1} = ", json_encode($ws->getRange($a1)->getValue()[$a1], JSON_UNESCAPED_UNICODE), "\n";
}

// ---------- 2. Bulk formula input with 2D array (setRangeFormulas) ----------
$formulas = [
    ["Subtotal"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
$ws->setRangeFormulas("D1", $formulas);

echo "[2] After setRangeFormulas:\n";
echo "    Formulas D2:D4 = ", json_encode($ws->getRange("D2:D4")->getFormula(), JSON_UNESCAPED_UNICODE), "\n";
echo "    Values   D2:D4 = ", json_encode($ws->getRange("D2:D4")->getValue(true), JSON_UNESCAPED_UNICODE), "\n";

// ---------- 3. Performance comparison (2,000 rows x 5 cols = 10,000 cells) ----------
$ROWS = 2000;
$COLS = 5;

// Per-cell input via getCells()
$wsCell = $wb->addWorkSheet("CellByCell");
$t = $now();
for ($r = 1; $r <= $ROWS; $r++) {
    for ($c = 1; $c <= $COLS; $c++) $wsCell->getCells($r, $c)->setNumberValue($r * 100 + $c);
}
$msCell = $now() - $t;

// 2D array input via setRangeData() (Fastest)
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

echo "[3] Performance for ", $ROWS * $COLS, " cells:\n";
echo "    Per-cell getCells().setNumberValue() : {$msCell} ms\n";
echo "    setRangeData (build {$msBuild} + write {$msApply}) : ", $msBuild + $msApply, " ms (fastest)\n";

$a = $wsCell->getRange("E2000:E2000")->getValue(true)["E2000"];
$b = $wsBulk->getRange("E2000:E2000")->getValue(true)["E2000"];
echo "[Verification] Both match E2000 = ", $a === (int)$b ? $b : "mismatch {$a} / {$b}", "\n";

$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// bulk-input.wasm.js — high-speed bulk input of values and formulas with 2D array (wasmosbxl / embind)
const app = new osbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. Bulk value input with 2D array (setRangeData) ----------
const data = [
    ["Product", "Price", "Qty"],
    ["Apple", "100", "3"],
    ["Orange", "50", "6"],
    ["Grape", "150", "2"],
];
ws.setRangeData("A1", data);

console.log("[1] After setRangeData:");
for (const a1 of ["A1", "B2", "C2"]) {
    console.log(`    ${a1} = ${JSON.stringify(ws.getRange(a1).getValue()[a1])}`);
}

// ---------- 2. Bulk formula input with 2D array (setRangeFormulas) ----------
const formulas = [
    ["Subtotal"],
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
ws.setRangeFormulas("D1", formulas);

console.log("[2] After setRangeFormulas:");
console.log("    Formulas D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getFormula()));
console.log("    Values   D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getValue(true)));

// ---------- 3. Performance comparison (2,000 rows x 5 cols = 10,000 cells) ----------
const ROWS = 2000, COLS = 5;

// Per-cell input via getCells()
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

// 2D array input via setRangeData() (Fastest)
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

console.log(`[3] Performance for ${ROWS * COLS} cells:`);
console.log(`    Per-cell getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setRangeData (build ${msBuild} + write ${msApply}) : ${msBuild + msApply} ms (fastest)`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[Verification] Both match E2000 = ${a === Number(b) ? b : `mismatch ${a} / ${b}`}`);

wb.save();
wb.close();
```

```plaintext
[1] After setRangeData:
    A1 = "Product"
    B2 = 100
    C2 = 3
[2] After setRangeFormulas:
    Formulas D2:D4 = {"D2":"B2*C2","D3":"B3*C3","D4":"B4*C4"}
    Values   D2:D4 = {"D2":300,"D3":300,"D4":300}
[3] Performance for 10000 cells:
    Per-cell getCells().setNumberValue() : 102 ms
    setRangeData (build 2 + write 18) : 20 ms (fastest)
[Verification] Both match E2000 = 200005
```

`[3]` execution time varies depending on the runtime environment.

## Explanation

### Speed and Performance

| Cell Count | Per-cell `getCells().setNumberValue()` | `setRangeData()` | Speedup |
| --- | --- | --- | --- |
| 10,000 (2,000 rows x 5 cols) | ~102 ms | ~20 ms (build 2 + write 18) | **~5x faster** |
| 100,000 (20,000 rows x 5 cols) | ~1,390 ms | ~195 ms (build 15 + write 180) | **~7x faster** |
| 200,000 (50,000 rows x 4 cols) | ~3,200 ms | ~380 ms | **~8.4x faster** |

`setRangeData()` minimizes language-boundary crossing overhead and performs direct SST (Shared String Table) registration and cell placement inside C++, widening the performance gap as cell count grows.

### Best Practices

* **Populating CSV / Database Records** — Pass a 2D array of rows and columns directly into `ws.setRangeData("A1", matrix)`.
* **Generating Calculation Columns** — Stream formulas row by row with `ws.setRangeFormulas("D2", [["B2*C2"], ["B3*C3"], ...])`.
* **When Writing Hundreds or More Cells** — Using `setRangeData()` / `setRangeFormulas()` is strongly recommended over per-cell setters.

## Notes

* **Specify the top-left start cell.** The first argument (e.g. `"A1"` or `"B5"`) acts as the anchor point from which the 2D matrix expands across rows and columns.
* **Numeric strings are automatically saved as numeric cells.** Strings that represent valid numbers (e.g., `"100"`, `"-12.5"`) are recognized as numeric cells in Excel and can be calculated by functions like SUM.
* **Do not prefix formulas with `=`.** Write `setRangeFormulas("D2", [["B2*C2"]])`.

## Related

* [Index](index.md)
* Previous: [Rows, Columns, and Ranges](general-05-rows-cols.md) / Next: [AutoFilter and Sort](general-07-autofilter-sort.md)
* API Reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
