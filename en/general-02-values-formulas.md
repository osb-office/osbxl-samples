# Building a Table with Values and Formulas

level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-02-values-formulas.md)


## What you want to do


Create a new xlsx, write a table (headers, strings, numbers, formulas) and save the file. This applies the basic flow from the [overview](general-01-overview.md) to real table building.


## Resulting image


`Sheet1` of `sales.xlsx` will contain this table:


|  | A | B | C | D |
| --- | --- | --- | --- | --- |
| **1** | Product | Unit Price | Qty | Amount |
| **2** | Apple | 100 | 3 | 300 |
| **3** | Orange | 50 | 6 | 300 |
| **4** | Grape | 150 | 2 | 300 |
| **5** |  |  | Total | 900 |


Column D holds the formula `B*C`; D5 holds `SUM(D2:D4)`.


## Prerequisites


Download the module for your language from the [download page](/pages/download.md) and install it (entry points are listed in the [overview Prerequisites](general-01-overview.md#prerequisites) table).


For detailed per-OS PHP installation steps (Linux / macOS / Windows, including when Composer is not installed), see the [Quickstart](/en/pages/quickstart.md).


The WebAssembly build uses MEMORY64 (Wasm 3.0), so the runtime must support it (Chrome/Edge 133+, Firefox 134+, Node.js 24+. Safari is not supported yet).


If you omit the directory in the path (`"sales.xlsx"`), the file is written to the **process current working directory**.


#### [Node.js]

```bash
npm install ./nodeosbxl-2.0.0.tgz
```


#### [Python]

```bash
# (example) Python 3.12+ (a single abi3 wheel covers 3.12 and later)
pip install ./pyosbxl-2.0.0-cp312-abi3-linux_x86_64.whl
```


#### [Java]

```bash
# Add the JAR to your classpath (javacpp-1.5.14.jar is also required)
javac -cp osbxl-2.0.0.jar:javacpp-1.5.14.jar Main.java
java  -cp osbxl-2.0.0.jar:javacpp-1.5.14.jar:. Main
```


#### [PHP]

```bash
# Place the downloaded .zip matching your OS and PHP version into the working directory
# (e.g. ~/phposbxl-test)
composer config repositories.phposbxl artifact ~/phposbxl-test/
composer require osb-office/phposbxl
vendor/bin/phposbxl-install
```


#### [WebAssembly]

```bash
# Install the package
npm install wasmosbxl-2.0.0.tgz
```


## Code


Output:


#### [Node.js]

```javascript
// values-formulas.js — create a book, write values and formulas, save
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sales.xlsx");   // new workbook
const ws = wb.openWorkSheet("Sheet1");         // open the default sheet

// --- header (row 1) ---
// app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -> "A1"
["Product", "Unit Price", "Qty", "Amount"].forEach((text, i) => {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- detail rows (2-4) ---
const data = [["Apple", 100, 3], ["Orange", 50, 6], ["Grape", 150, 2]];
data.forEach(([name, price, qty], i) => {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- amount column (D2:D4) ---
// Each row needs its own relative reference, so use setRangeFormulas.
// Formula strings do NOT take a leading "=".
const formulas = data.map((_, i) => [`B${i + 2}*C${i + 2}`]);
ws.setRangeFormulas("D2", formulas);

// --- total (D5) ---
ws.getRange("C5:C5").setValue("Total");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- read back ---
// getValue(true) returns raw values = formula results
const values = ws.getRange("A1:D5").getValue(true);
console.log("Amount D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("Total:", values["D5"]);

// --- save and close ---
wb.save();
wb.close();
```


#### [Python]

```python
# values-formulas.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sales-py.xlsx")   # new workbook
ws = wb.openWorkSheet("Sheet1")            # open the default sheet

# --- header (row 1) ---
# app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -> "A1"
for i, text in enumerate(["Product", "Unit Price", "Qty", "Amount"]):
    cell = app.convertFromRowColNumber(1, i + 1)
    ws.getRange(f"{cell}:{cell}").setValue(text)

# --- detail rows (2-4) ---
data = [("Apple", 100, 3), ("Orange", 50, 6), ("Grape", 150, 2)]
for i, (name, price, qty) in enumerate(data):
    row = i + 2
    ws.getRange(f"A{row}:A{row}").setValue(name)
    ws.getRange(f"B{row}:B{row}").setNumberValue(price)
    ws.getRange(f"C{row}:C{row}").setNumberValue(qty)

# --- amount column (D2:D4) ---
# Each row needs its own relative reference, so use setRangeFormulas.
# Formula strings do NOT take a leading "=".
formulas = [[f"B{i + 2}*C{i + 2}"] for i in range(len(data))]
ws.setRangeFormulas("D2", formulas)

# --- total (D5) ---
ws.getRange("C5:C5").setValue("Total")
ws.getRange("D5:D5").setFormula("SUM(D2:D4)")

# --- read back ---
# getValue(True) returns raw values = formula results
values = ws.getRange("A1:D5").getValue(True)
print("Amount D2:D4 =", ", ".join([values["D2"], values["D3"], values["D4"]]))
print("Total:", values["D5"])

# --- save and close ---
wb.save()
wb.close();
```


#### [Java]

```java
// values-formulas.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("sales-java.xlsx");   // new workbook
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");             // open the default sheet

        // --- header (row 1) ---
        // app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -> "A1"
        String[] headers = { "Product", "Unit Price", "Qty", "Amount" };
        for (int i = 0; i < headers.length; i++) {
            String cell = app.convertFromRowColNumber(1, i + 1);
            ws.getRange(cell + ":" + cell).setValue(headers[i]);
        }

        // --- detail rows (2-4) ---
        String[] names = { "Apple", "Orange", "Grape" };
        double[] prices = { 100, 50, 150 };
        double[] qtys = { 3, 6, 2 };
        for (int i = 0; i < names.length; i++) {
            int row = i + 2;
            ws.getRange("A" + row + ":A" + row).setValue(names[i]);
            ws.getRange("B" + row + ":B" + row).setNumberValue(prices[i]);
            ws.getRange("C" + row + ":C" + row).setNumberValue(qtys[i]);
        }

        // --- amount column (D2:D4) ---
        // Each row needs its own relative reference, so use setRangeFormulas.
        // Formula strings do NOT take a leading "=".
        List<List<String>> formulas = new ArrayList<>();
        for (int i = 0; i < names.length; i++) {
            formulas.add(Collections.singletonList("B" + (i + 2) + "*C" + (i + 2)));
        }
        ws.setRangeFormulas("D2", formulas);

        // --- total (D5) ---
        ws.getRange("C5:C5").setValue("Total");
        ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

        // --- read back ---
        // getValue(true) returns raw values = formula results
        Map<String, String> values = ws.getRange("A1:D5").getValue(true);
        System.out.println("Amount D2:D4 = "
                + String.join(", ", values.get("D2"), values.get("D3"), values.get("D4")));
        System.out.println("Total: " + values.get("D5"));

        // --- save and close ---
        wb.save();
        wb.close();
    }
}
```


#### [PHP]

```php
// values-formulas.php — same verification as the Node.js sample (phposbxl)
// Body only (<?php and the autoload require are added by the harness)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sales-php.xlsx");   // new workbook
$ws = $wb->openWorkSheet("Sheet1");             // open the default sheet

// --- header (row 1) ---
// $app->convertFromRowColNumber(row, col) gives the A1 address for (1,1) -> "A1"
$headers = ["Product", "Unit Price", "Qty", "Amount"];
foreach ($headers as $i => $text) {
    $cell = $app->convertFromRowColNumber(1, $i + 1);
    $ws->getRange("{$cell}:{$cell}")->setValue($text);
}

// --- detail rows (2-4) ---
$data = [["Apple", 100, 3], ["Orange", 50, 6], ["Grape", 150, 2]];
foreach ($data as $i => [$name, $price, $qty]) {
    $row = $i + 2;
    $ws->getRange("A{$row}:A{$row}")->setValue($name);
    $ws->getRange("B{$row}:B{$row}")->setNumberValue($price);
    $ws->getRange("C{$row}:C{$row}")->setNumberValue($qty);
}

// --- amount column (D2:D4) ---
// Each row needs its own relative reference, so use setRangeFormulas.
// Formula strings do NOT take a leading "=".
$formulas = [];
for ($i = 0; $i < count($data); $i++) {
    $row = $i + 2;
    $formulas[] = ["B{$row}*C{$row}"];
}
$ws->setRangeFormulas("D2", $formulas);

// --- total (D5) ---
$ws->getRange("C5:C5")->setValue("Total");
$ws->getRange("D5:D5")->setFormula("SUM(D2:D4)");

// --- read back ---
// getValue(true) returns raw values = formula results
$values = $ws->getRange("A1:D5")->getValue(true);
echo "Amount D2:D4 = " . implode(", ", [$values["D2"], $values["D3"], $values["D4"]]) . "\n";
echo "Total: " . $values["D5"] . "\n";

// --- save and close ---
$wb->save();
$wb->close();
```


#### [WebAssembly]

```javascript
// values-formulas.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("sales-wasm.xlsx");   // new workbook
const ws = wb.openWorkSheet("Sheet1");              // open the default sheet

// --- header (row 1) ---
// app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -> "A1"
["Product", "Unit Price", "Qty", "Amount"].forEach((text, i) => {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- detail rows (2-4) ---
const data = [["Apple", 100, 3], ["Orange", 50, 6], ["Grape", 150, 2]];
data.forEach(([name, price, qty], i) => {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- amount column (D2:D4) ---
// Each row needs its own relative reference, so use setRangeFormulas.
// Formula strings do NOT take a leading "=".
const formulas = data.map((_, i) => [`B${i + 2}*C${i + 2}`]);
ws.setRangeFormulas("D2", formulas);

// --- total (D5) ---
ws.getRange("C5:C5").setValue("Total");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- read back ---
// getValue(true) returns raw values = formula results
const values = ws.getRange("A1:D5").getValue(true);
console.log("Amount D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("Total:", values["D5"]);

// --- save and close ---
wb.save();
wb.close();
```


`Amount D2:D4 = 300, 300, 300
Total: 900`

<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`new nodeosbxl.App()` is the single entry point.** From `App` you get a `WorkBook` via `createWorkBook()` / `openWorkBook()`, then a `WorkSheet` via `openWorkSheet()`, then a `Range` via `getRange()`.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Cell addresses are strings like `"A1:B2"`.** Even a single cell uses the range form `"A1:A1"`. To build them from 1-based row/column numbers use `app.convertFromRowColNumber(row, col)`, or `app.convertFromRowColNumber2(startRow, startCol, endRow, endCol)` for ranges.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Each value type has its own setter.** `setValue(string)` / `setNumberValue(number)` / `setDateValue(DateTimeObject)` / `setDateStringValue(date string)` / `setBooleanValue(bool)`.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**The `true` in `getValue(true)` means "raw value"** (formula result). Without it you get the display value (after number formatting). The return value is a map like `{ "A1": "100", ... }` and all values are strings; convert with `Number(...)` when comparing.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Finish with `wb.save()` then `wb.close()`.** Use `wb.saveAs(path)` to write a second file, but note that `saveAs()` does **not** switch the destination of later saves (see [overview caveats](general-01-overview.md#notes)).</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Do not prefix formulas with `=`.** Use `setFormula("SUM(A1:A3)")`, not `setFormula("=SUM(A1:A3)")`.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`setValue` / `setNumberValue` on a multi-cell Range writes the same value to every cell.** `ws.getRange("A1:C1").setValue("X")` puts `"X"` into A1, B1 and C1.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`setFormula()` does not auto-fill relative references.** `ws.getRange("B1:B3").setFormula("A1*2")` puts the formula into **B1 only**; B2 / B3 stay empty.**The third argument variant `setFormula("A1*2", false, true)` (`setAllCell = true`) merely copies the same string into every cell — it does <strong>not** become `A2*2`, `A3*2` (all of B1:B3 end up with the same value). To shift formulas per row, use **`setRangeFormulas()` (high-speed bulk write via 2D array)** as in this sample.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**A license banner line is printed to stdout once per process.** Keep it in mind when parsing standard output.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li>[Index](index.md)</li>
  <li>Previous: [Overview — object model and basic flow](general-01-overview.md) / Next: [Working with books and sheets](general-03-sheets.md)</li>
<li>
  API reference
  
    
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
    
    [App](/api-docs/en/nodejs/classes/nodeosbxl.App.html) / [WorkBook](/api-docs/en/nodejs/classes/nodeosbxl.WorkBook.html) / [WorkSheet](/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html) / [Range](/api-docs/en/nodejs/classes/nodeosbxl.Range.html)


[App](/api-docs/en/python/pyosbxl.App.html) / [WorkBook](/api-docs/en/python/pyosbxl.WorkBook.html) / [WorkSheet](/api-docs/en/python/pyosbxl.WorkSheet.html) / [Range](/api-docs/en/python/pyosbxl.Range.html)


[App](/api-docs/en/java/com/osboffice/osbxl/AppWrapper.html) / [WorkBook](/api-docs/en/java/com/osboffice/osbxl/WorkBookWrapper.html) / [WorkSheet](/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html) / [Range](/api-docs/en/java/com/osboffice/osbxl/RangeWrapper.html)


[App](/api-docs/en/php/classes/phposbxl-App.html) / [WorkBook](/api-docs/en/php/classes/phposbxl-WorkBook.html) / [WorkSheet](/api-docs/en/php/classes/phposbxl-WorkSheet.html) / [Range](/api-docs/en/php/classes/phposbxl-Range.html)


[App](/api-docs/en/wasm/classes/osbxl.App.html) / [WorkBook](/api-docs/en/wasm/classes/osbxl.WorkBook.html) / [WorkSheet](/api-docs/en/wasm/classes/osbxl.WorkSheet.html) / [Range](/api-docs/en/wasm/classes/osbxl.Range.html)
