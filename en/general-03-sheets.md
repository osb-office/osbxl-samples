# Books and Worksheets — add, delete, move, rename

level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-03-sheets.md)


## What you want to do


Add, delete, move and rename worksheets within a single book, walk sheets by index, and control which sheet is shown when the book opens (the active tab).


## Methods covered


| Method | Belongs to | Description |
| --- | --- | --- |
| `addWorkSheet(sheetName, position?)` | WorkBook | Adds a worksheet. |
| `deleteWorkSheet(sheetName)` | WorkBook | Deletes a worksheet. |
| `moveSheet(sheetName, position)` | WorkBook | Moves a worksheet. |
| `openWorkSheet(sheetName)` | WorkBook | Opens a worksheet by name. |
| `openWorkSheet(sheetIndex)` | WorkBook | Opens a worksheet by index. |
| `getSheetCount()` | WorkBook | Gets the number of worksheets. |
| `getSheetNumber(sheetName)` | WorkBook | Gets the sheet number of a worksheet. |
| `getName()` / `setName(name)` | WorkSheet | Gets / sets the sheet name. |
| `isActive()` / `activate(isActive)` | WorkSheet | Gets / sets the active tab (the sheet shown when the book opens). |


## Code


Output:


#### [Node.js]

```javascript
// sheets.js — add, move, rename and delete sheets, plus index access
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sheets.xlsx");

// Indexes are 1-based. Open 1..getSheetCount() in order and list the names
const show = (label) => {
    const names = [];
    for (let i = 0; i < wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheet(i).getName());
    }
    console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};

show("[1] after creation:");

wb.addWorkSheet("Summary");                     // position omitted → appended at the end
show("[2] addWorkSheet(Summary):");

wb.addWorkSheet("Orders", 1);                   // position=1 → inserted at the front
show("[3] addWorkSheet(Orders,1):");

wb.moveSheet("Orders", 3);                      // move to position 3
show("[4] moveSheet(Orders,3):");

const ws = wb.openWorkSheet("Orders");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(Orders) =", wb.getSheetNumber("Orders"));
ws.getRange("A1:A1").setValue("order data");

wb.openWorkSheet("Sheet1").setName("Details");  // rename
show("[6] setName(Details):");

// --- active tab (the sheet shown in front when the book is opened) ---
console.log("[7] isActive: Summary =", wb.openWorkSheet("Summary").isActive(),
            "/ Details =", wb.openWorkSheet("Details").isActive());
wb.openWorkSheet("Summary").activate(true);  // the other sheets are deactivated automatically
console.log("    after activate: Summary =", wb.openWorkSheet("Summary").isActive(),
            "/ Details =", wb.openWorkSheet("Details").isActive());

// --- delete ---
wb.addWorkSheet("Temp");
wb.deleteWorkSheet("Temp");
show("[8] deleteWorkSheet(Temp):");

wb.save();
wb.close();

// --- reopen and verify ---
const wb2 = app.openWorkBook("sheets.xlsx");
const names = [];
for (let i = 0; i < wb2.getSheetCount(); i++) names.push(wb2.openWorkSheet(i).getName());
console.log("[End] sheets:", names.join(", "));
console.log("     Orders!A1 =", wb2.openWorkSheet("Orders").getRange("A1:A1").getValue()["A1"]);
wb2.close();
```


#### [Python]

```python
# sheets.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sheets-py.xlsx")

def b(x):
    # render Python's True/False using node-style true/false notation
    return "true" if x else "false"

# Indexes are 1-based. Open 1..getSheetCount() in order and list the names
def show(label):
    names = [wb.openWorkSheet(i).getName() for i in range(1, wb.getSheetCount() + 1)]
    print(f"{label} count={wb.getSheetCount()} [{', '.join(names)}]")

show("[1] after creation:")

wb.addWorkSheet("Summary")                      # position omitted → appended at the end
show("[2] addWorkSheet(Summary):")

wb.addWorkSheet("Orders", 1)                    # position=1 → inserted at the front
show("[3] addWorkSheet(Orders,1):")

wb.moveSheet("Orders", 3)                       # move to position 3
show("[4] moveSheet(Orders,3):")

ws = wb.openWorkSheet("Orders")
print("[5] getName() =", ws.getName(), "/ getSheetNumber(Orders) =", wb.getSheetNumber("Orders"))
ws.getRange("A1:A1").setValue("order data")

wb.openWorkSheet("Sheet1").setName("Details")   # rename
show("[6] setName(Details):")

# --- active tab (the sheet shown in front when the book is opened) ---
print("[7] isActive: Summary =", b(wb.openWorkSheet("Summary").isActive()),
      "/ Details =", b(wb.openWorkSheet("Details").isActive()))
wb.openWorkSheet("Summary").activate(True)      # the other sheets are deactivated automatically
print("    after activate: Summary =", b(wb.openWorkSheet("Summary").isActive()),
      "/ Details =", b(wb.openWorkSheet("Details").isActive()))

# --- delete ---
wb.addWorkSheet("Temp")
wb.deleteWorkSheet("Temp")
show("[8] deleteWorkSheet(Temp):")

wb.save()
wb.close()

# --- reopen and verify ---
wb2 = app.openWorkBook("sheets-py.xlsx")
names = [wb2.openWorkSheet(i).getName() for i in range(1, wb2.getSheetCount() + 1)]
print("[End] sheets:", ", ".join(names))
print("     Orders!A1 =", wb2.openWorkSheet("Orders").getRange("A1:A1").getValue()["A1"])
wb2.close()
```


#### [Java]

```java
// sheets.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    static WorkBookWrapper wb;

    // Indexes are 1-based. Open 1..getSheetCount() in order and list the names
    static void show(String label) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb.getSheetCount(); i++) {
            names.add(wb.openWorkSheet(i).getName());
        }
        System.out.println(label + " count=" + wb.getSheetCount()
                + " [" + String.join(", ", names) + "]");
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        wb = app.createWorkBook("sheets-java.xlsx");

        show("[1] after creation:");

        wb.addWorkSheet("Summary");                     // position omitted → appended at the end
        show("[2] addWorkSheet(Summary):");

        wb.addWorkSheet("Orders", 1);                   // position=1 → inserted at the front
        show("[3] addWorkSheet(Orders,1):");

        wb.moveSheet("Orders", 3);                      // move to position 3
        show("[4] moveSheet(Orders,3):");

        WorkSheetWrapper ws = wb.openWorkSheet("Orders");
        System.out.println("[5] getName() = " + ws.getName()
                + " / getSheetNumber(Orders) = " + wb.getSheetNumber("Orders"));
        ws.getRange("A1:A1").setValue("order data");

        wb.openWorkSheet("Sheet1").setName("Details");  // rename
        show("[6] setName(Details):");

        // --- active tab (the sheet shown in front when the book is opened) ---
        System.out.println("[7] isActive: Summary = " + wb.openWorkSheet("Summary").isActive()
                + " / Details = " + wb.openWorkSheet("Details").isActive());
        wb.openWorkSheet("Summary").activate(true);     // the other sheets are deactivated automatically
        System.out.println("    after activate: Summary = " + wb.openWorkSheet("Summary").isActive()
                + " / Details = " + wb.openWorkSheet("Details").isActive());

        // --- delete ---
        wb.addWorkSheet("Temp");
        wb.deleteWorkSheet("Temp");
        show("[8] deleteWorkSheet(Temp):");

        wb.save();
        wb.close();

        // --- reopen and verify ---
        WorkBookWrapper wb2 = app.openWorkBook("sheets-java.xlsx");
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb2.getSheetCount(); i++) names.add(wb2.openWorkSheet(i).getName());
        System.out.println("[End] sheets: " + String.join(", ", names));
        Map<String, String> v = wb2.openWorkSheet("Orders").getRange("A1:A1").getValue();
        System.out.println("     Orders!A1 = " + v.get("A1"));
        wb2.close();
    }
}
```


#### [PHP]

```php
// sheets.php — same verification as the Node.js sample (phposbxl)
// Body only (<?php and the autoload require are added by the harness)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("sheets-php.xlsx");

// Render PHP booleans using node-style true/false notation
function b($x) {
    return $x ? "true" : "false";
}

// Indexes are 1-based. Open 1..getSheetCount() in order and list the names
function show($label) {
    global $wb;
    $names = [];
    for ($i = 0; $i < $wb->getSheetCount(); $i++) {
        $names[] = $wb->openWorkSheet($i)->getName();
    }
    echo $label . " count=" . $wb->getSheetCount() . " [" . implode(", ", $names) . "]\n";
}

show("[1] after creation:");

$wb->addWorkSheet("Summary");                     // position omitted → appended at the end
show("[2] addWorkSheet(Summary):");

$wb->addWorkSheet("Orders", 1);                   // position=1 → inserted at the front
show("[3] addWorkSheet(Orders,1):");

$wb->moveSheet("Orders", 3);                      // move to position 3
show("[4] moveSheet(Orders,3):");

$ws = $wb->openWorkSheet("Orders");
echo "[5] getName() = " . $ws->getName() . " / getSheetNumber(Orders) = " . $wb->getSheetNumber("Orders") . "\n";
$ws->getRange("A1:A1")->setValue("order data");

$wb->openWorkSheet("Sheet1")->setName("Details"); // rename
show("[6] setName(Details):");

// --- active tab (the sheet shown in front when the book is opened) ---
echo "[7] isActive: Summary = " . b($wb->openWorkSheet("Summary")->isActive())
   . " / Details = " . b($wb->openWorkSheet("Details")->isActive()) . "\n";
$wb->openWorkSheet("Summary")->activate(true);    // the other sheets are deactivated automatically
echo "    after activate: Summary = " . b($wb->openWorkSheet("Summary")->isActive())
   . " / Details = " . b($wb->openWorkSheet("Details")->isActive()) . "\n";

// --- delete ---
$wb->addWorkSheet("Temp");
$wb->deleteWorkSheet("Temp");
show("[8] deleteWorkSheet(Temp):");

$wb->save();
$wb->close();

// --- reopen and verify ---
$wb2 = $app->openWorkBook("sheets-php.xlsx");
$names = [];
for ($i = 0; $i < $wb2->getSheetCount(); $i++) {
    $names[] = $wb2->openWorkSheet($i)->getName();
}
echo "[End] sheets: " . implode(", ", $names) . "\n";
$v = $wb2->openWorkSheet("Orders")->getRange("A1:A1")->getValue();
echo "     Orders!A1 = " . $v["A1"] . "\n";
$wb2->close();
```


#### [WebAssembly]

```javascript
// sheets.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("sheets-wasm.xlsx");

// Indexes are 1-based. Open 1..getSheetCount() in order and list the names
const show = (label) => {
    const names = [];
    for (let i = 0; i < wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheet(i).getName());
    }
    console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};

show("[1] after creation:");

wb.addWorkSheet("Summary");                     // position omitted → appended at the end
show("[2] addWorkSheet(Summary):");

wb.addWorkSheet("Orders", 1);                   // position=1 → inserted at the front
show("[3] addWorkSheet(Orders,1):");

wb.moveSheet("Orders", 3);                      // move to position 3
show("[4] moveSheet(Orders,3):");

const ws = wb.openWorkSheet("Orders");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(Orders) =", wb.getSheetNumber("Orders"));
ws.getRange("A1:A1").setValue("order data");

wb.openWorkSheet("Sheet1").setName("Details");  // rename
show("[6] setName(Details):");

// --- active tab (the sheet shown in front when the book is opened) ---
console.log("[7] isActive: Summary =", wb.openWorkSheet("Summary").isActive(),
            "/ Details =", wb.openWorkSheet("Details").isActive());
wb.openWorkSheet("Summary").activate(true);  // the other sheets are deactivated automatically
console.log("    after activate: Summary =", wb.openWorkSheet("Summary").isActive(),
            "/ Details =", wb.openWorkSheet("Details").isActive());

// --- delete ---
wb.addWorkSheet("Temp");
wb.deleteWorkSheet("Temp");
show("[8] deleteWorkSheet(Temp):");

wb.save();
wb.close();

// --- reopen and verify ---
const wb2 = app.openWorkBook("sheets-wasm.xlsx");
const names = [];
for (let i = 0; i < wb2.getSheetCount(); i++) names.push(wb2.openWorkSheet(i).getName());
console.log("[End] sheets:", names.join(", "));
console.log("     Orders!A1 =", wb2.openWorkSheet("Orders").getRange("A1:A1").getValue()["A1"]);
wb2.close();
```


`[1] after creation: count=1 [Sheet1]
[2] addWorkSheet(Summary): count=2 [Sheet1, Summary]
[3] addWorkSheet(Orders,1): count=3 [Orders, Sheet1, Summary]
[4] moveSheet(Orders,3): count=3 [Sheet1, Summary, Orders]
[5] getName() = Orders / getSheetNumber(Orders) = 3
[6] setName(Details): count=3 [Details, Summary, Orders]
[7] isActive: Summary = false / Details = true
    after activate: Summary = true / Details = false
[8] deleteWorkSheet(Temp): count=3 [Details, Summary, Orders]
[End] sheets: Details, Summary, Orders
     Orders!A1 = order data`

<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Both `position` and sheet numbers are 1-based.** `addWorkSheet(name, position)` **inserts before `position`** (`position` = 1 means the front). Omitting it (default -1) appends at the end.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**The `position` in `moveSheet(name, position)` is the "position after the move"**, with the front being 1.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`openWorkSheet(i)` is also 1-based.** Combined with `getSheetCount()` it lets you enumerate all sheets by name (the `show()` above). When handling a book whose sheet names you don't know, walking it this way is the reliable approach.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`getSheetNumber(name)` returns the number in the current ordering.** The value changes after `moveSheet()` or `deleteWorkSheet()`.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**`setName()` takes effect immediately.** From then on, `openWorkSheet()` with the new name.**Trying to open it by the old name raises an exception.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>`activate(true)` works exclusively.** When you `activate(true)` the target sheet, the other sheets are **deactivated automatically** (measured). You do not need to `activate(false)` the others beforehand.**The active state is kept across `save()` and reopening.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Passing a nonexistent sheet name to `openWorkSheet()` raises an exception.**`wb.openWorkSheet("NoSuch");   // → exception: "NoSuch is not found"`When sheet names come from user input or external configuration, first enumerate the names with `getSheetCount()` + `openWorkSheet()` and match against them.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**An out-of-range index for `openWorkSheet()` also raises an exception.** Call it with a value in the range `1` to `getSheetCount()`.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Right after creating a new book, only the default sheet (`Sheet1`) has `isActive() === true`** (measured). Sheets added with `addWorkSheet()` are `false`. If you want a different sheet shown in front when the book opens, call `activate(true)` on that sheet.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**Use `copyRow` / `copyCol` / `copyCell` to copy cells across sheets** (the copy source is specified by sheet name). See [Working with rows, columns and cell ranges](general-05-rows-cols.md) for details.  
Shapes (including charts), pivot tables, queryTable-style tables and data tables are not included in what gets copied.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li>[Index](index.md)</li>
  <li>Previous: [Building a table with values and formulas](general-02-values-formulas.md) / Next: [Cell value types and number formats](general-04-value-types.md)</li>
<li>
  API reference
  
    
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
    
    [WorkBook](/api-docs/en/nodejs/classes/nodeosbxl.WorkBook.html) / [WorkSheet](/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html)


[WorkBook](/api-docs/en/python/pyosbxl.WorkBook.html) / [WorkSheet](/api-docs/en/python/pyosbxl.WorkSheet.html)


[WorkBook](/api-docs/en/java/com/osboffice/osbxl/WorkBookWrapper.html) / [WorkSheet](/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html)


[WorkBook](/api-docs/en/php/classes/phposbxl-WorkBook.html) / [WorkSheet](/api-docs/en/php/classes/phposbxl-WorkSheet.html)


[WorkBook](/api-docs/en/wasm/classes/osbxl.WorkBook.html) / [WorkSheet](/api-docs/en/wasm/classes/osbxl.WorkSheet.html)
