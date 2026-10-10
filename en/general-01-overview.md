# Overview — Object Model and Basic Flow

level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-01-overview.md)

## What you want to do

Get the big picture of nodeosbxl.

* The object hierarchy (`App` → `WorkBook` → `WorkSheet` → `Range`)
* The **basic flow** (create new / open existing book → sheet → Range → values & formulas → save / saveAs / close)
* The main methods of the four classes

Details of individual features are covered by the samples that follow this page.

## Prerequisites

Download the module for your language from the [download page](https://osb-office.com/pages/download.php) and install it.

For detailed per-OS PHP installation steps (Linux / macOS / Windows, including when Composer is not installed), see the [Quickstart](/en/pages/quickstart.md).

The WebAssembly build uses MEMORY64 (Wasm 3.0), so the runtime must support it (Chrome/Edge 133+, Firefox 134+, Node.js 24+. Safari is not supported yet).

The entry point differs per language; the namespace layout (operation classes / dto / chart / enums) is common to all four.

#### [Node.js]

```bash
npm install ./nodeosbxl-1.5.0.tgz
```

#### [Python]

```bash
# (example) Python 3.12+ (a single abi3 wheel covers 3.12 and later)
pip install ./pyosbxl-1.5.0-cp312-abi3-linux\_x86\_64.whl
```

#### [Java]

```bash
# Add the JAR to your classpath (javacpp-1.5.14.jar is also required)
javac -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar Main.java
java -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar:. Main
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
npm install wasmosbxl-1.5.0.tgz
```
```javascript
// Loading example (requires Node.js 24+: this wasm is a MEMORY64 = Wasm 3.0 build)
const fs = require("fs");
const createModule = require("./package/osbxl\_wasm.js"); // emscripten factory
const { createOsbxl } = require("./package/osbxl\_em.js"); // loader + enum definitions
createOsbxl(createModule).then(({ Module, osbxl, dto, chart, enums }) => {
// Stage the license file into the wasm virtual FS (required before new App())
const lic = process.env.OSB\_LICENSE\_PATH\_WASM; // e.g. /root/lic\_node.osb
if (lic && fs.existsSync(lic)) {
Module.FS.mkdirTree(lic.substring(0, lic.lastIndexOf("/")));
Module.FS.writeFile(lic, fs.readFileSync(lic));
}
const app = new osbxl.App(); // from here on it matches the Node.js build (namespace name is osbxl)
// ...
});
// In the browser, load osbxl\_wasm.js via <script> and call createOsbxl the same way.
// Supported: Chrome/Edge 133+, Firefox 134+ (Safari not supported yet)
```

| Language | Entry point | Main namespaces |
| --- | --- | --- |
| Node.js | `require("nodeosbxl")` | `nodeosbxl` / `dto` / `chart` / `enums` |
| Python | `import pyosbxl` | `pyosbxl` / `pyosbxl.dto` / `pyosbxl.chart` / `pyosbxl.enums` |
| Java | `import com.osboffice.osbxl.*` | packages such as `dto` / `chart` / `enums` |
| WebAssembly | `createOsbxl(createModule)` | `osbxl` / `dto` / `chart` / `enums` |

The table below uses Node.js names; the counts per namespace are common to all four languages.

| Namespace | Contents | Count |
| --- | --- | --- |
| `nodeosbxl` | Classes you operate on (App / WorkBook / WorkSheet / Range / Font / Chart entry points, etc.) | 29 classes |
| `dto` | Objects used to pass data around (`InputValueObject`, `FontObject`, `ColorObject`, etc.) | 40 classes |
| `chart` | Chart-specific objects (`ChartObject`, `Axis`, `ChartTitle`, etc.) | 43 classes |
| `enums` | Excel-compatible enumerations (`XlFont`, `XlChartType`, `XlIndexColor`, etc.) | 127 kinds |

## Object model

```plaintext
App ← the single entry point. Open/create books, convert cell addresses
│
├─ createWorkBook() ─┐
├─ openWorkBook() ─┴→ WorkBook ← one xlsx file
│ │
│ ├─ openWorkSheet() → WorkSheet ← one sheet
│ │ │
│ │ ├─ getRange("A1:B2") → Range
│ │ │ ※ getCells(1,1,2,2) returns the same Range
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
└─ utilities such as convertFromRowColNumber()
```

**Always walking this hierarchy top-down** is the osbxl basics. Below `Range` (`Font`, `Color`, `Borders`, …) the pattern is "get the object, then call its setters".

## Basic flow

```plaintext
[1] new App() create the App
[2] createWorkBook(path) create a new book
openWorkBook(path) open an existing book
[3] openWorkSheet(name) open a worksheet
[4] getRange("A1:B2") get a Range (getCells(1,1,2,2) is the same)
[5] setValue() / setFormula() write values / formulas
[6] getValue(true) read the results back
[7] save() / saveAs(path) save
[8] close() close the book
```

### Code

Output:

#### [Node.js]

```javascript
// basic-flow.js — the full flow: create → edit → save → reopen → save as
const { nodeosbxl } = require("nodeosbxl");
const app = new nodeosbxl.App(); // [1] create the App
// ---------- create a new book, edit it and save ----------
let wb = app.createWorkBook("flow.xlsx"); // [2] create new
console.log("[2] createWorkBook -> getBookPath =", wb.getBookPath());
let ws = wb.openWorkSheet("Sheet1"); // [3] open a worksheet
console.log("[3] openWorkSheet -> sheet name =", ws.getName());
// [4] get a Range and [5] set values
// ws.getCells(1, 1) (numeric row/col) returns the same Range as getRange("A1:A1")
ws.getRange("A1:A1").setValue("created new");
ws.getRange("B1:B1").setFormula("100\*3"); // [5] set a formula (no leading "=")
console.log("[6] getValue(true) -> B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);
wb.save(); // [7] save
console.log("[7] save() done");
wb.close(); // [8] close (no close needed for WorkSheet)
console.log("[8] close() done");
// ---------- open the existing book and save it under a new name ----------
wb = app.openWorkBook("flow.xlsx"); // [2'] open an existing book
ws = wb.openWorkSheet("Sheet1"); // [3']
const read = ws.getRange("A1:B1").getValue(true);
console.log("[2'] openWorkBook -> A1 =", read["A1"], "/ B1 =", read["B1"]);
ws.getRange("A2:A2").setValue("appended after reopen");
wb.saveAs("flow-copy.xlsx"); // [7'] save under a new name
console.log("[7'] saveAs() done -> getBookPath =", wb.getBookPath());
wb.close(); // [8']
console.log("[8] close() done");
```

#### [Python]

```python
# basic-flow.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl
app = pyosbxl.App() # [1] create the App
wb = app.createWorkBook("flow-py.xlsx") # [2] create new
ws = wb.openWorkSheet("Sheet1") # [3] open a worksheet
ws.getRange("A1:A1").setValue("created new")
ws.getRange("B1:B1").setFormula("100\*3") # [5] formula (no leading "=")
print("[6] getValue(True) -> B1 =", ws.getRange("B1:B1").getValue(True)["B1"])
wb.save() # [7] save
wb.close() # [8] close
print("[8] close() done")
wb = app.openWorkBook("flow-py.xlsx") # [2'] open an existing book
ws = wb.openWorkSheet("Sheet1")
print("[2'] openWorkBook -> A1 =", ws.getRange("A1:A1").getValue(True)["A1"])
wb.close()
print("[8] close() done")
```

#### [Java]

```java
// basic-flow.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper(); // [1] create the App
WorkBookWrapper wb = app.createWorkBook("flow-java.xlsx"); // [2] create new
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1"); // [3] open a worksheet
ws.getRange("A1:A1").setValue("created new");
ws.getRange("B1:B1").setFormula("100\*3"); // [5] formula (no leading "=")
Map<String, String> v = ws.getRange("B1:B1").getValue(true);
System.out.println("[6] getValue(true) -> B1 = " + v.get("B1"));
wb.save(); // [7] save
wb.close(); // [8] close
System.out.println("[8] close() done");
wb = app.openWorkBook("flow-java.xlsx"); // [2'] open an existing book
ws = wb.openWorkSheet("Sheet1");
Map<String, String> r = ws.getRange("A1:A1").getValue(true);
System.out.println("[2'] openWorkBook -> A1 = " + r.get("A1"));
wb.close();
System.out.println("[8] close() done");
}
}
```

#### [PHP]

```php
// basic-flow.md — same verification as the Node.js sample (phposbxl)
// Body only (<?php and the autoload require are added by the harness)
$app = new \phposbxl\App(); // [1] create the App
// ---------- create a new book, edit it and save ----------
$wb = $app->createWorkBook("flow-php.xlsx"); // [2] create new
echo "[2] createWorkBook -> getBookPath = " . $wb->getBookPath() . "\n";
$ws = $wb->openWorkSheet("Sheet1"); // [3] open a worksheet
echo "[3] openWorkSheet -> sheet name = " . $ws->getName() . "\n";
// [4] get a Range and [5] set values
// $ws->getCells(1, 1) (numeric row/col) returns the same Range as getRange("A1:A1")
$ws->getRange("A1:A1")->setValue("created new");
$ws->getRange("B1:B1")->setFormula("100\*3"); // [5] set a formula (no leading "=")
$v = $ws->getRange("B1:B1")->getValue(true);
echo "[6] getValue(true) -> B1 = " . $v["B1"] . "\n";
$wb->save(); // [7] save
echo "[7] save() done\n";
$wb->close(); // [8] close (no close needed for WorkSheet)
echo "[8] close() done\n";
// ---------- open the existing book and save it under a new name ----------
$wb = $app->openWorkBook("flow-php.xlsx"); // [2'] open an existing book
$ws = $wb->openWorkSheet("Sheet1"); // [3']
$read = $ws->getRange("A1:B1")->getValue(true);
echo "[2'] openWorkBook -> A1 = " . $read["A1"] . " / B1 = " . $read["B1"] . "\n";
$ws->getRange("A2:A2")->setValue("appended after reopen");
$wb->saveAs("flow-copy-php.xlsx"); // [7'] save under a new name
echo "[7'] saveAs() done -> getBookPath = " . $wb->getBookPath() . "\n";
$wb->close(); // [8']
echo "[8] close() done\n";
```

#### [WebAssembly]

```javascript
// basic-flow.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App(); // [1] create the App
const wb = app.createWorkBook("flow-wasm.xlsx"); // [2] create new
const ws = wb.openWorkSheet("Sheet1"); // [3] open a worksheet
ws.getRange("A1:A1").setValue("created new");
ws.getRange("B1:B1").setFormula("100\*3"); // [5] formula (no leading "=")
console.log("[6] getValue(true) -> B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);
wb.save(); // [7] save
wb.close(); // [8] close
console.log("[8] close() done");
const wb2 = app.openWorkBook("flow-wasm.xlsx"); // [2'] open an existing book
const ws2 = wb2.openWorkSheet("Sheet1");
console.log("[2'] openWorkBook -> A1 =", ws2.getRange("A1:A1").getValue(true)["A1"]);
wb2.close();
console.log("[8] close() done");
```

```plaintext
[2] createWorkBook -> getBookPath = <working directory>/flow.xlsx
[3] openWorkSheet -> sheet name = Sheet1
[6] getValue(true) -> B1 = 300
[7] save() done
[8] close() done
[2'] openWorkBook -> A1 = created new / B1 = 300
[7'] saveAs() done -> getBookPath = <working directory>/flow.xlsx
[8] close() done
```

Two files are produced: `flow.xlsx` and `flow-copy.xlsx`.

### Flow highlights

* **One `App` per process is enough.** Even when handling multiple books, call `createWorkBook()` / `openWorkBook()` on the same `app`.
* **Relative paths are resolved to absolute paths.** After `createWorkBook("flow.xlsx")`, `getBookPath()` returns an absolute path including the current directory.
* **There is no operation to close a `WorkSheet`.** `close()` exists only as a `WorkBook` method. Sheets are all released together by `wb.close()`.
* **Note the difference between `save()` and `saveAs()`** (see "Caveats" below).

## Main methods

### App (10 methods)

Created with `new nodeosbxl.App()`. Handles opening/closing books and address conversion utilities.

| Method | Description |
| --- | --- |
| `getVersion()` | Returns the version number. |
| `createWorkBook(path, defaultFont?, defaultFontSize?)` | Creates a workbook. |
| `openWorkBook(path)` | Opens a workbook. |
| `openPasswordWorkBook(path, password)` | Opens a password-protected workbook. |
| `convertFromRowColNumber(row, col)` | Converts numeric cell coordinates to A1C1 notation. `(1,1)` → `"A1"` |
| `convertFromRowColNumber2(startRow, startCol, endRow, endCol)` | Converts a numeric cell range to A1C1 notation. `(1,1,2,3)` → `"A1:C2"` |
| `convertToColumnNumber("B")` | Converts a column letter to its numeric index. → `2` |
| `convertFromColumnNumber(3)` | Converts a numeric column index to its letter. → `"C"` |
| `getNumericValue(dateTimeObject, is1904?)` | Gets the Excel internal serial value from a date-time object. |
| `calculateHexColor(bookPath, colorObj, isForeGroundColor)` | Gets the HEX value taking workbook-specific color information into account. |

`convertFromRowColNumber` / `convertFromRowColNumber2` also have overloads that return absolute references (`"$A$1"`).

### WorkBook (27 methods) — main 12

Corresponds to one xlsx file.

| Method | Description |
| --- | --- |
| `getBookPath()` | Returns the workbook's file path. |
| `openWorkSheet(sheetName)` | Opens a worksheet. |
| `openWorkSheetByIndex(sheetIndex)` | Opens a worksheet by index. |
| `addWorkSheet(sheetName, position?)` | Adds a worksheet. |
| `deleteWorkSheet(sheetName)` | Deletes a worksheet. |
| `moveSheet(sheetName, position)` | Moves a worksheet. |
| `getSheetCount()` | Gets the number of worksheets. |
| `getSheetNumber(sheetName)` | Gets the sheet number of a worksheet. |
| `save()` / `save(password)` | Saves the workbook to the same file. |
| `saveAs(path)` / `saveAs(path, password)` | Saves the workbook to a different file. |
| `close()` | Closes the workbook. |
| `isDate1904()` | Gets whether the workbook uses the Excel 1904 date system. |

Others (15): styles — `getNames` `addCustomStyle` `getCustomStyle` `deleteCustomStyle` `addCustomTableStyle` `addPivotCustomTableStyle` `getCustomTableStyle` `deleteCustomTableStyle` / external books — `addExternalWorkBook` `updateExternalWorkBook` `getExternalWorkBookPath` / file properties — `setCompanyName` `setManagerName` `setCreateAuthor` `setLastAuthor`

### WorkSheet (44 methods) — main 14

Corresponds to one sheet. **This is the gateway to the objects below it (Range / Sort / PivotTable / Chart …).**

| Method | Description |
| --- | --- |
| `getRange("A1:B2")` | Gets a cell range class instance. |
| `getCells(row, col)` / `getCells(sr, sc, er, ec)` | Gets a cell range class instance from numeric coordinates. |
| `getRow(rowNum)` | Gets a row class instance. |
| `getCol(colNum)` | Gets a column class instance. |
| `getName()` / `setName(name)` | Gets / sets the sheet name. |
| `isActive()` / `activate(...)` | Gets / sets whether the sheet tab is the one shown when the book opens. |
| `insertRow(startRow, numOfRows)` | Inserts rows. |
| `deleteRow(startRow, numOfRows)` | Deletes rows. |
| `insertCol(startCol, numOfCols)` | Inserts columns. |
| `deleteCol(startCol, numOfCols)` | Deletes columns. |
| `setValueArray(values)` | Bulk-sets cell values. |
| `setFormulaArray(formulas)` | Bulk-sets formulas. |

Others (30): feature objects — `getFormatConditions` `getAutoFilter` `getSort` `getComments` `getHyperLinks` `getListObjects` `getPivotTables` `getShapes` `getWindow` `getHPageBreaks` `getVPageBreaks` / charts — `getChartObject` `addChartObject` `deleteChartObject` / print setup — `getPageSetupObject` `setPageSetupObject` / selection — `getActiveCell` `setActiveCell` / row, column and cell operations — `clearRow` `copyRow` `clearCol` `copyCol` `insertCell` `deleteCell` `clearCell` `copyCell` / outlining — `setRowOutline` `clearRowOutline` `setColOutline` `clearColOutline`

### Range (31 methods) — main 16

Corresponds to a cell range. **Reads/writes values and formulas, and is the gateway to formatting objects.**

| Method | Description |
| --- | --- |
| `getValue(rawValue?)` | Gets cell values. `true` returns raw values (formula results). |
| `getFormula()` | Gets the formula. |
| `setValue(value, forceString?, numberFormat?)` | Sets a cell value (general-purpose method). |
| `setNumberValue(value, ...)` | Sets a cell value (numeric setter). |
| `setDateValue(dateTimeObject, ...)` | Sets a cell value (date/time setter). |
| `setDateStringValue(str, ...)` | Sets a cell value (date/time from a string). |
| `setBooleanValue(bool, ...)` | Sets a cell value (boolean setter). |
| `setFormula(formula, isArray?, setAllCell?)` | Sets a formula. |
| `getFont()` | Gets the font object instance. |
| `getBorders()` | Gets the borders collection object instance. |
| `getFill()` | Gets the fill object instance. |
| `getAlignment()` | Gets the cell alignment object instance. |
| `setNumberFormat(format)` | Sets the number format. |
| `merge(mergeEachRow?)` | Merges the cell range. |
| `clearContent()` | Clears the cells. |
| `getAddress()` | Gets the selected range of the Range. |

Others (15): `getProtection` `getCharacters` `getPhonetics` `getNumberFormat` `clearNumberFormat` `clearAllFormat` `setDataTable` `clearDataTable` `setBuiltinStyle` `setCustomStyle` `getTotalWidth` `getTotalColumnWidth` `getTotalRowHeight` `unMerge` `replaceValue`

## Notes

* **`Range` is "get it and throw it away".** `ws.getRange("A1:A1")` returns a new instance on every call, so there is no need to keep holding it in a variable.
* **`getValue()` returns a map.** Shaped like `{ "A1": "100", "B1": "300" }`, and all values are **strings**. Convert with `Number(...)` to treat them as numbers.
* **Two ways to get a Range.**
  + `ws.getRange("A1:B2")` — by **string address**. Even a single cell is passed in range form as `"A1:A1"`.
  + `ws.getCells(row, col)` / `ws.getCells(startRow, startCol, endRow, endCol)` — by **numbers (row, col)**. Rows and columns are **1-based**: `(1,1)` = A1, `(3,1,3,3)` = A3:C3. Handy when looping over row/column numbers because you don't have to assemble address strings.
  + For number ⇔ string conversion use `app.convertFromRowColNumber(row, col)` / `app.convertFromRowColNumber2(sr, sc, er, ec)` / `app.convertToColumnNumber("B")`.
* **Formulas are computed at run time.** Right after `setFormula()` you can read the result with `getValue(true)` (no need to wait for `save()`).

## Caveats

* **A workbook is locked internally to guarantee uniqueness, so simultaneous read/write of the same workbook never occurs.**
* **`close()` exists only as a `WorkBook` method.** `App` / `WorkSheet` / `Range` have no close-style methods.
* **`saveAs()` only "writes the current content to a different file"; it does not switch the destination of later saves.** This differs from Excel VBA's `SaveAs`.

  ```plaintext
  wb.save() → written to base.xlsx
  wb.saveAs("copy.xlsx") → "the content at this point" is written to copy.xlsx
  wb.getBookPath() → still base.xlsx (not switched)
  ws.getRange(...).setValue(...)
  wb.save() → written to base.xlsx (not copy.xlsx)
  ```
* **Do not prefix formulas with `=`.** Use `setFormula("SUM(A1:A3)")`.
* **A license banner line is printed to stdout at startup** (once per process). Keep it in mind when writing code that parses standard output.

## See also

* [Index](index.md)
* Next sample: [Building a table with values and formulas](general-02-values-formulas.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
