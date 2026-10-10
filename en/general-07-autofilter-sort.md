# AutoFilter and Sorting — filter, sort, read back

level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-07-autofilter-sort.md)

## What you want to do

Set an **AutoFilter** on a data table so that only rows matching a condition are shown, and sort (ascending / descending) with **Sort**. A filter works by **hiding the rows** that do not match the condition; whether a row is hidden can be read with `Row.getHidden()`.

| Approach | What to use |
| --- | --- |
| Set filter / conditions | `ws.getAutoFilter(range, firstRowAsHeader)` → the `set*Filter` methods on `AutoFilter` |
| Clear conditions | `af.resetAllFilter()` (conditions only) / `af.removeFilter()` (removes the filter itself) |
| Sort | `ws.getSort(range)` / `af.getSort()` → `executeSortAscending` / `executeSortDescending` |
| Check hidden rows | `ws.getRow(r).getHidden()` |

## Classes covered

### `nodeosbxl.AutoFilter` — obtained via `ws.getAutoFilter(A1C1, firstRowAsHeader?)`

| Method | Description |
| --- | --- |
| `getAddress()` | Returns the target cell range. |
| `setCustomFilter(op, criteria, colId?)` | Filter with an operator + criteria. `criteria` supports the `*` and `?` wildcards. |
| `setAndCustomFilter(op1, c1, op2, c2, colId?)` / `setOrCustomFilter(...)` | AND / OR combination of two conditions. |
| `setTop10ValueFilter(n, colId?)` / `setBottom10ValueFilter(n, colId?)` | Top / bottom n rows by value. |
| `setTop10PercentFilter(p, colId?)` / `setBottom10PercentFilter(p, colId?)` | Top / bottom p percent by value. |
| `setAverageFilter(aboveAverage, colId?)` | Above average (true) / below average (false). |
| `setUniqueValuesNumberFilter(vals, colId?)` | Show only rows matching one of the numbers in the list. |
| `setUniqueValuesStringFilter` / `setUniqueValuesDateTimeFilter` | String / date-time versions (see [Caveats](#caveats)). |
| `setDateTimeFilter` / `setDateTimeGroupingFilter` | Date period / grouping conditions. |
| `setFontColorFilter` / `setCellColorFilter` / `setIconFilter` | Conditions based on color / icon. |
| `getSort()` | Returns a `Sort` that protects the header row. |
| `resetFilter(colId)` / `resetAllFilter()` | Clear conditions (the filter itself remains). |
| `removeFilter()` | Removes the AutoFilter itself and restores hidden rows. |

`op` is `enums.XlAutoFilterOperator`:

| Value | Number | Meaning |
| --- | --- | --- |
| `Equal` | 1 | equal to |
| `LessThan` | 2 | less than |
| `LessThanOrEqual` | 3 | less than or equal to |
| `NotEqual` | 4 | not equal to |
| `GreaterThanOrEqual` | 5 | greater than or equal to |
| `GreaterThan` | 6 | greater than |

### `nodeosbxl.Sort` — obtained via `ws.getSort(A1C1, firstRowAsHeader?)` or `af.getSort()`

| Method | Description |
| --- | --- |
| `executeSortAscending(target, direction?, matchCase?)` | Sorts ascending. `target` is the column number within the range (first column = 1). |
| `executeSortDescending(target, direction?, matchCase?)` | Sorts descending. |
| `execute(sortFieldObject)` | Sorts by a single `dto.SortFieldObject` (configure it with `setSortOnValues(target, ascending)`). |
| `executeMultiple([...])` | Multi-key sort (works correctly since the 2026-10-01 fix. See [Caveats](#caveats)). |
| `getSortConditions()` | Reads back the configured conditions as an array of `dto.SortFieldObject`. |
| `resetSort(target)` / `resetAllSort()` | Clear conditions. **Only effective on a Sort obtained via AutoFilter / Table.** |

`direction` is `enums.XlRowCol` (`Rows` = 1 / `Columns` = 2). The default is `Rows` (sort rows).

## Code

Output:

#### [Node.js]

```javascript
// af-sort.js — AutoFilter and sorting
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("af-sort.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. prepare data (header + 10 rows) ----------
const DATA = [
["Apple", "Tokyo", 120, 100],
["Orange", "Osaka", 40, 80],
["Banana", "Tokyo", 200, 150],
["Grape", "Fukuoka", 60, 300],
["Pear", "Osaka", 90, 120],
["Peach", "Fukuoka", 150, 400],
["Kiwi", "Tokyo", 30, 90],
["Mango", "Osaka", 70, 500],
["Strawberry", "Fukuoka", 110, 350],
["Melon", "Tokyo", 25, 800],
];
["Product", "Region", "Qty", "Unit Price"].forEach((h, i) =>
ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) => {
const r = i + 2;
ws.getRange(`A${r}:A${r}`).setValue(row[0]);
ws.getRange(`B${r}:B${r}`).setValue(row[1]);
ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] data ready: A1:D11 (header + 10 rows)");
// display helpers: row number → "Product(value)" / list of rows that are not hidden
const text = (r, col) => ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) => `${text(r, "A")}(${text(r, col)})`;
const visible = (col) => {
const out = [];
for (let r = 2; r <= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
return out.join(" ");
};
const hiddenList = () => {
const out = [];
for (let r = 2; r <= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
return out.length ? out.join(",") : "none";
};
// ---------- 2. set up the AutoFilter ----------
const af = ws.getAutoFilter("A1:D11", true); // first row is the header
console.log(`[2] getAutoFilter → range = ${af.getAddress()} / hidden rows = ${hiddenList()}`);
// ---------- 3. custom filter: Qty > 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] visible rows with Qty>100: ${visible("C")}`);
console.log(` hidden rows = ${hiddenList()} / getValue("A2:A11") count = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter(); // clear conditions only (the filter itself remains)
// ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3); // top 3 by Qty
console.log(`[4] visible rows in top 3 by Qty: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3); // Qty is 120 or 40
console.log(`[5] visible rows with Qty in {120,40}: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2); // Region is Tokyo or Fukuoka (string list)
console.log(`[6] visible rows with Region in {Tokyo,Fukuoka}: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "Osaka", 2); // Region = Osaka (string match)
console.log(`[7] visible rows with Region="Osaka": ${visible("B")}`);
af.resetAllFilter();
// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
ws.getSort("A2:D11").executeSortAscending(3); // ascending by Qty (3rd column within the range)
console.log(`[8] ascending by Qty: ${visible("C")}`);
// a Sort obtained via the AutoFilter protects the header row automatically
const sort = af.getSort();
sort.executeSortDescending(4); // descending by Unit Price (4th column)
console.log(`[9] descending by Unit Price: ${visible("D")}`);
// ---------- 6. read the sort conditions back ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] condition read back: target column = ${cond.getTarget()} / ascending = ${cond.getSortAscending()}`);
// ---------- 7. remove the AutoFilter itself ----------
af.removeFilter();
console.log(`[11] after removeFilter: hidden rows = ${hiddenList()}`);
wb.save();
wb.close();
```

#### [Python]

```python
# af-sort.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("af-sort-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- 1. prepare data (header + 10 rows) ----------
DATA = [
["Apple", "Tokyo", 120, 100],
["Orange", "Osaka", 40, 80],
["Banana", "Tokyo", 200, 150],
["Grape", "Fukuoka", 60, 300],
["Pear", "Osaka", 90, 120],
["Peach", "Fukuoka", 150, 400],
["Kiwi", "Tokyo", 30, 90],
["Mango", "Osaka", 70, 500],
["Strawberry", "Fukuoka", 110, 350],
["Melon", "Tokyo", 25, 800],
]
for i, h in enumerate(["Product", "Region", "Qty", "Unit Price"]):
col = "ABCD"[i]
ws.getRange(f"{col}1:{col}1").setValue(h)
for i, row in enumerate(DATA):
r = i + 2
ws.getRange(f"A{r}:A{r}").setValue(row[0])
ws.getRange(f"B{r}:B{r}").setValue(row[1])
ws.getRange(f"C{r}:C{r}").setNumberValue(row[2])
ws.getRange(f"D{r}:D{r}").setNumberValue(row[3])
print("[1] data ready: A1:D11 (header + 10 rows)")
# display helpers: row number → "Product(value)" / list of rows that are not hidden
def text(r, col):
return ws.getRange(f"{col}{r}:{col}{r}").getValue()[f"{col}{r}"]
def label(r, col):
return f"{text(r, 'A')}({text(r, col)})"
def visible(col):
out = []
for r in range(2, 12):
if not ws.getRow(r).getHidden():
out.append(label(r, col))
return " ".join(out)
def hidden\_list():
out = [str(r) for r in range(2, 12) if ws.getRow(r).getHidden()]
return ",".join(out) if out else "none"
# ---------- 2. set up the AutoFilter ----------
af = ws.getAutoFilter("A1:D11", True) # first row is the header
print(f"[2] getAutoFilter → range = {af.getAddress()} / hidden rows = {hidden\_list()}")
# ---------- 3. custom filter: Qty > 100 ----------
# colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.GreaterThan, "100", 3)
print(f"[3] visible rows with Qty>100: {visible('C')}")
print(f' hidden rows = {hidden\_list()} / getValue("A2:A11") count = {len(ws.getRange("A2:A11").getValue())}')
af.resetAllFilter() # clear conditions only (the filter itself remains)
# ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3) # top 3 by Qty
print(f"[4] visible rows in top 3 by Qty: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesNumberFilter([120, 40], 3) # Qty is 120 or 40
print(f"[5] visible rows with Qty in {{120,40}}: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2) # Region is Tokyo or Fukuoka (string list)
print(f"[6] visible rows with Region in {{Tokyo,Fukuoka}}: {visible('B')}")
af.resetAllFilter()
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.Equal, "Osaka", 2) # Region = Osaka (string match)
print(f'[7] visible rows with Region="Osaka": {visible("B")}')
af.resetAllFilter()
# ---------- 5. sorting (single key) ----------
# target only the data rows without the header (firstRowAsHeader omitted = False)
ws.getSort("A2:D11").executeSortAscending(3) # ascending by Qty (3rd column within the range)
print(f"[8] ascending by Qty: {visible('C')}")
# a Sort obtained via the AutoFilter protects the header row automatically
sort = af.getSort()
sort.executeSortDescending(4) # descending by Unit Price (4th column)
print(f"[9] descending by Unit Price: {visible('D')}")
# ---------- 6. read the sort conditions back ----------
cond = sort.getSortConditions()[0]
# Python prints booleans as True/False, so use lower() to match the JS style
print(f"[10] condition read back: target column = {cond.getTarget()} / ascending = {str(cond.getSortAscending()).lower()}")
# ---------- 7. remove the AutoFilter itself ----------
af.removeFilter()
print(f"[11] after removeFilter: hidden rows = {hidden\_list()}")
wb.save()
wb.close()
```

#### [Java]

```java
// af-sort.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
public class Main {
static WorkSheetWrapper ws;
// display helpers: row number → "Product(value)" / list of rows that are not hidden
static String text(int r, String col) {
return ws.getRange(col + r + ":" + col + r).getValue().get(col + r);
}
static String label(int r, String col) {
return text(r, "A") + "(" + text(r, col) + ")";
}
static String visible(String col) {
List<String> out = new ArrayList<>();
for (int r = 2; r <= 11; r++) {
if (!ws.getRow(r).getHidden()) out.add(label(r, col));
}
return String.join(" ", out);
}
static String hiddenList() {
List<String> out = new ArrayList<>();
for (int r = 2; r <= 11; r++) {
if (ws.getRow(r).getHidden()) out.add(String.valueOf(r));
}
return out.isEmpty() ? "none" : String.join(",", out);
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("af-sort-java.xlsx");
ws = wb.openWorkSheet("Sheet1");
// ---------- 1. prepare data (header + 10 rows) ----------
Object[][] DATA = {
{"Apple", "Tokyo", 120.0, 100.0},
{"Orange", "Osaka", 40.0, 80.0},
{"Banana", "Tokyo", 200.0, 150.0},
{"Grape", "Fukuoka", 60.0, 300.0},
{"Pear", "Osaka", 90.0, 120.0},
{"Peach", "Fukuoka", 150.0, 400.0},
{"Kiwi", "Tokyo", 30.0, 90.0},
{"Mango", "Osaka", 70.0, 500.0},
{"Strawberry", "Fukuoka", 110.0, 350.0},
{"Melon", "Tokyo", 25.0, 800.0},
};
String[] headers = {"Product", "Region", "Qty", "Unit Price"};
for (int i = 0; i < headers.length; i++) {
String col = String.valueOf("ABCD".charAt(i));
ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
}
for (int i = 0; i < DATA.length; i++) {
int r = i + 2;
ws.getRange("A" + r + ":A" + r).setValue((String) DATA[i][0]);
ws.getRange("B" + r + ":B" + r).setValue((String) DATA[i][1]);
ws.getRange("C" + r + ":C" + r).setNumberValue((Double) DATA[i][2]);
ws.getRange("D" + r + ":D" + r).setNumberValue((Double) DATA[i][3]);
}
System.out.println("[1] data ready: A1:D11 (header + 10 rows)");
// ---------- 2. set up the AutoFilter ----------
AutoFilterWrapper af = ws.getAutoFilter("A1:D11", true); // first row is the header
System.out.println("[2] getAutoFilter → range = " + af.getAddress() + " / hidden rows = " + hiddenList());
// ---------- 3. custom filter: Qty > 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(XlAutoFilterOperator.GreaterThan, "100", 3);
System.out.println("[3] visible rows with Qty>100: " + visible("C"));
System.out.println(" hidden rows = " + hiddenList()
+ " / getValue(\"A2:A11\") count = " + ws.getRange("A2:A11").getValue().size());
af.resetAllFilter(); // clear conditions only (the filter itself remains)
// ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3); // top 3 by Qty
System.out.println("[4] visible rows in top 3 by Qty: " + visible("C"));
af.resetAllFilter();
af.setUniqueValuesNumberFilter(Arrays.asList(120.0, 40.0), 3); // Qty is 120 or 40
System.out.println("[5] visible rows with Qty in {120,40}: " + visible("C"));
af.resetAllFilter();
af.setUniqueValuesStringFilter(Arrays.asList("Tokyo", "Fukuoka"), 2); // Region is Tokyo or Fukuoka (string list)
System.out.println("[6] visible rows with Region in {Tokyo,Fukuoka}: " + visible("B"));
af.resetAllFilter();
af.setCustomFilter(XlAutoFilterOperator.Equal, "Osaka", 2); // Region = Osaka (string match)
System.out.println("[7] visible rows with Region=\"Osaka\": " + visible("B"));
af.resetAllFilter();
// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
ws.getSort("A2:D11").executeSortAscending(3); // ascending by Qty (3rd column within the range)
System.out.println("[8] ascending by Qty: " + visible("C"));
// a Sort obtained via the AutoFilter protects the header row automatically
SortWrapper sort = af.getSort();
sort.executeSortDescending(4); // descending by Unit Price (4th column)
System.out.println("[9] descending by Unit Price: " + visible("D"));
// ---------- 6. read the sort conditions back ----------
SortFieldObjectWrapper cond = sort.getSortConditions().get(0);
System.out.println("[10] condition read back: target column = " + cond.getTarget()
+ " / ascending = " + cond.getSortAscending());
// ---------- 7. remove the AutoFilter itself ----------
af.removeFilter();
System.out.println("[11] after removeFilter: hidden rows = " + hiddenList());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// af-sort.md — AutoFilter and sorting (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("af-sort-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// ---------- 1. prepare data (header + 10 rows) ----------
$DATA = [
["Apple", "Tokyo", 120, 100],
["Orange", "Osaka", 40, 80],
["Banana", "Tokyo", 200, 150],
["Grape", "Fukuoka", 60, 300],
["Pear", "Osaka", 90, 120],
["Peach", "Fukuoka", 150, 400],
["Kiwi", "Tokyo", 30, 90],
["Mango", "Osaka", 70, 500],
["Strawberry", "Fukuoka", 110, 350],
["Melon", "Tokyo", 25, 800],
];
foreach (["Product", "Region", "Qty", "Unit Price"] as $i => $h) {
$col = "ABCD"[$i];
$ws->getRange("{$col}1:{$col}1")->setValue($h);
}
foreach ($DATA as $i => $row) {
$r = $i + 2;
$ws->getRange("A{$r}:A{$r}")->setValue($row[0]);
$ws->getRange("B{$r}:B{$r}")->setValue($row[1]);
$ws->getRange("C{$r}:C{$r}")->setNumberValue($row[2]);
$ws->getRange("D{$r}:D{$r}")->setNumberValue($row[3]);
}
echo "[1] data ready: A1:D11 (header + 10 rows)\n";
// display helpers: row number → "Product(value)" / list of rows that are not hidden
$text = fn(int $r, string $col): string => $ws->getRange("{$col}{$r}:{$col}{$r}")->getValue()["{$col}{$r}"];
$label = fn(int $r, string $col): string => $text($r, "A") . "(" . $text($r, $col) . ")";
$visible = function (string $col) use ($ws, $label): string {
$out = [];
for ($r = 2; $r <= 11; $r++) {
if (!$ws->getRow($r)->getHidden()) {
$out[] = $label($r, $col);
}
}
return implode(" ", $out);
};
$hiddenList = function () use ($ws): string {
$out = [];
for ($r = 2; $r <= 11; $r++) {
if ($ws->getRow($r)->getHidden()) {
$out[] = (string)$r;
}
}
return count($out) > 0 ? implode(",", $out) : "none";
};
// ---------- 2. set up the AutoFilter ----------
$af = $ws->getAutoFilter("A1:D11", true); // first row is the header
echo "[2] getAutoFilter → range = ", $af->getAddress(), " / hidden rows = ", $hiddenList(), "\n";
// ---------- 3. custom filter: Qty > 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
$af->setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::GreaterThan, "100", 3);
echo "[3] visible rows with Qty>100: ", $visible("C"), "\n";
echo " hidden rows = ", $hiddenList(),
" / getValue(\"A2:A11\") count = ", count($ws->getRange("A2:A11")->getValue()), "\n";
$af->resetAllFilter(); // clear conditions only (the filter itself remains)
// ---------- 4. Top-N / value-list filters ----------
$af->setTop10ValueFilter(3, 3); // top 3 by Qty
echo "[4] visible rows in top 3 by Qty: ", $visible("C"), "\n";
$af->resetAllFilter();
$af->setUniqueValuesNumberFilter([120, 40], 3); // Qty is 120 or 40
echo "[5] visible rows with Qty in {120,40}: ", $visible("C"), "\n";
$af->resetAllFilter();
$af->setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2); // Region is Tokyo or Fukuoka (string list)
echo "[6] visible rows with Region in {Tokyo,Fukuoka}: ", $visible("B"), "\n";
$af->resetAllFilter();
$af->setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::Equal, "Osaka", 2); // Region = Osaka (string match)
echo "[7] visible rows with Region=\"Osaka\": ", $visible("B"), "\n";
$af->resetAllFilter();
// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
$ws->getSort("A2:D11")->executeSortAscending(3); // ascending by Qty (3rd column within the range)
echo "[8] ascending by Qty: ", $visible("C"), "\n";
// a Sort obtained via the AutoFilter protects the header row automatically
$sort = $af->getSort();
$sort->executeSortDescending(4); // descending by Unit Price (4th column)
echo "[9] descending by Unit Price: ", $visible("D"), "\n";
// ---------- 6. read the sort conditions back ----------
$cond = $sort->getSortConditions()[0];
// echo of a PHP bool prints 1/empty string, so render JS-style true/false
echo "[10] condition read back: target column = ", $cond->getTarget(),
" / ascending = ", $cond->getSortAscending() ? "true" : "false", "\n";
// ---------- 7. remove the AutoFilter itself ----------
$af->removeFilter();
echo "[11] after removeFilter: hidden rows = ", $hiddenList(), "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// af-sort.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("af-sort-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. prepare data (header + 10 rows) ----------
const DATA = [
["Apple", "Tokyo", 120, 100],
["Orange", "Osaka", 40, 80],
["Banana", "Tokyo", 200, 150],
["Grape", "Fukuoka", 60, 300],
["Pear", "Osaka", 90, 120],
["Peach", "Fukuoka", 150, 400],
["Kiwi", "Tokyo", 30, 90],
["Mango", "Osaka", 70, 500],
["Strawberry", "Fukuoka", 110, 350],
["Melon", "Tokyo", 25, 800],
];
["Product", "Region", "Qty", "Unit Price"].forEach((h, i) =>
ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) => {
const r = i + 2;
ws.getRange(`A${r}:A${r}`).setValue(row[0]);
ws.getRange(`B${r}:B${r}`).setValue(row[1]);
ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] data ready: A1:D11 (header + 10 rows)");
// display helpers: row number → "Product(value)" / list of rows that are not hidden
const text = (r, col) => ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) => `${text(r, "A")}(${text(r, col)})`;
const visible = (col) => {
const out = [];
for (let r = 2; r <= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
return out.join(" ");
};
const hiddenList = () => {
const out = [];
for (let r = 2; r <= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
return out.length ? out.join(",") : "none";
};
// ---------- 2. set up the AutoFilter ----------
const af = ws.getAutoFilter("A1:D11", true); // first row is the header
console.log(`[2] getAutoFilter → range = ${af.getAddress()} / hidden rows = ${hiddenList()}`);
// ---------- 3. custom filter: Qty > 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] visible rows with Qty>100: ${visible("C")}`);
console.log(` hidden rows = ${hiddenList()} / getValue("A2:A11") count = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter(); // clear conditions only (the filter itself remains)
// ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3); // top 3 by Qty
console.log(`[4] visible rows in top 3 by Qty: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3); // Qty is 120 or 40
console.log(`[5] visible rows with Qty in {120,40}: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2); // Region is Tokyo or Fukuoka (string list)
console.log(`[6] visible rows with Region in {Tokyo,Fukuoka}: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "Osaka", 2); // Region = Osaka (string match)
console.log(`[7] visible rows with Region="Osaka": ${visible("B")}`);
af.resetAllFilter();
// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
ws.getSort("A2:D11").executeSortAscending(3); // ascending by Qty (3rd column within the range)
console.log(`[8] ascending by Qty: ${visible("C")}`);
// a Sort obtained via the AutoFilter protects the header row automatically
const sort = af.getSort();
sort.executeSortDescending(4); // descending by Unit Price (4th column)
console.log(`[9] descending by Unit Price: ${visible("D")}`);
// ---------- 6. read the sort conditions back ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] condition read back: target column = ${cond.getTarget()} / ascending = ${cond.getSortAscending()}`);
// ---------- 7. remove the AutoFilter itself ----------
af.removeFilter();
console.log(`[11] after removeFilter: hidden rows = ${hiddenList()}`);
wb.save();
wb.close();
```

```plaintext
[1] data ready: A1:D11 (header + 10 rows)
[2] getAutoFilter → range = A1:D11 / hidden rows = none
[3] visible rows with Qty>100: Apple(120) Banana(200) Peach(150) Strawberry(110)
hidden rows = 3,5,6,8,9,11 / getValue("A2:A11") count = 10
[4] visible rows in top 3 by Qty: Apple(120) Banana(200) Peach(150)
[5] visible rows with Qty in {120,40}: Apple(120) Orange(40)
[6] visible rows with Region in {Tokyo,Fukuoka}: Apple(Tokyo) Banana(Tokyo) Grape(Fukuoka) Peach(Fukuoka) Kiwi(Tokyo) Strawberry(Fukuoka) Melon(Tokyo)
[7] visible rows with Region="Osaka": Orange(Osaka) Pear(Osaka) Mango(Osaka)
[8] ascending by Qty: Melon(25) Kiwi(30) Orange(40) Grape(60) Mango(70) Pear(90) Strawberry(110) Apple(120) Peach(150) Banana(200)
[9] descending by Unit Price: Melon(800) Mango(500) Peach(400) Strawberry(350) Grape(300) Banana(150) Pear(120) Apple(100) Kiwi(90) Orange(80)
[10] condition read back: target column = 4 / ascending = false
[11] after removeFilter: hidden rows = none
```

## Notes

### Filters = hiding rows

An AutoFilter only **hides the rows that do not match the condition**; no data is deleted.

* Check the hidden state with `ws.getRow(r).getHidden()` (see [Rows, columns and ranges](general-05-rows-cols.md) for `Row` usage).
* **`getValue()` returns the values of hidden rows as well** (output `[3]`: the count is 10 even though 6 rows are hidden). If you want "only the visible rows", narrow them down with `getHidden()` before reading.
* `resetFilter(colId)` / `resetAllFilter()` clear **only the conditions**; the filter (frame) remains.
* `removeFilter()` removes the filter itself and makes all hidden rows visible again (output `[11]`).
* The filter range, conditions, and hidden state are **restored after save → reopen** (measured).

### Writing filter conditions

* Both `colId` in `setCustomFilter(op, criteria, colId)` and `target` in `Sort` are **the column number within the specified range (first column = 1)** — not the sheet column number itself, so watch out for the offset when the range starts at column `B`.
* Write **only the value** in `criteria` (`"100"`). The operator is specified via `op` (`enums.XlAutoFilterOperator`).
* Numeric comparisons (`GreaterThan` etc.), Top-N, average, and number-list matching (`setUniqueValuesNumberFilter`) all work as expected (output `[3]`–`[5]`).

### Sorting

* The basic form is to **pass only the data rows as the range**, as in `ws.getSort("A2:D11")`. If you pass a range that includes the header, use `ws.getSort("A1:D11", true)` (2nd argument `firstRowAsHeader`) or **`af.getSort()` (the header is protected automatically)**.
* Sorting moves **whole rows**. Columns other than the key column (Product, Region, etc.) move along with their row.
* After running, read the conditions back with `getSortConditions()` (`getTarget()` / `getSortAscending()` / `getDirection()` / `getMatchCase()`). The read-back conditions also survive save → reopen.
* The `direction` in `executeSortAscending(target, direction, matchCase)` is `enums.XlRowCol.Rows`(1) / `Columns`(2). The default is `Rows`.

## Caveats

* **Pass only the value in `criteria` (no operator symbols, no wildcards).** The operator is specified via `enums.XlAutoFilterOperator` by design.  
  Mixing symbols in, as in `setCustomFilter(GreaterThan, ">100", 3)`, makes the condition be treated as a string; on a numeric column it matches nothing and all rows are hidden (`"100"` works fine).
* **String match filters work correctly** (fixed on 2026-10-01. In the old version 1.5.0 there was a bug where string cells newly created within the session did not match and all rows were hidden). `setCustomFilter(Equal, "Tokyo", ...)` / `setUniqueValuesStringFilter(["Tokyo","Fukuoka"], ...)` / `setOrCustomFilter(...)` all work as shown in output `[6]` `[7]`.
* **`executeMultiple()` (multi-key sort) works correctly** (fixed on 2026-10-01). In the old version 1.5.0 rows were duplicated or lost.  
  Build each condition with `SortFieldObject.setSortOnValues(column, ascending?)` and pass them as an array in priority order.  
  Acceptance test: `test/sort_af_fc_ext.test.js` "Sort executeMultiple (multiple keys, A-2 fix)".
* **Calling `executeSortAscending()` multiple times does not produce a multi-key sort.** The whole table is re-sorted by the key called last, while conditions keep accumulating in `getSortConditions()` (measured).
* **Sorting a header-inclusive range without `firstRowAsHeader` (default false) sorts the header row as data.** In a measured run, an ascending sort sank the row-1 header (a string) to the very bottom.  
  For ranges that include the header, use `getSort(range, true)` or `af.getSort()`.
* **Sorting while a filter is applied reorders only the visible rows.** Hidden rows stay at their original positions, so it is safer to `resetAllFilter()` before sorting (measured).
* **`resetSort()` / `resetAllSort()` work only on a Sort obtained via AutoFilter / Table.** We confirmed by measurement that conditions are not cleared on a Sort obtained via `ws.getSort()` (the specification as documented).
* **`getAutoFilter(A1C1, firstRowAsHeader?)` is a "getter" that also sets the range.** There is only one AutoFilter range per sheet; calling it with a different range replaces the configured range.  
  The 2nd argument is not "whether to get" but `firstRowAsHeader` (whether the first row is treated as the header).
* **Put numbers in numeric columns with `setNumberValue()` (or values interpreted as numbers).** We measured that comparison filters do work on numbers entered as strings such as `setValue("120")`, but numeric-typed input is recommended for stable behavior of Top-N and similar filters.

## See also

* [Index](index.md)
* Previous: [Bulk Input of Large Data](general-06-bulk-input.md) / Next: [Tables and Comments](general-08-tables-comments.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
