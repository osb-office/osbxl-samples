# Pivot Tables — create, aggregate, delete

level: advanced ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-11-pivottables.md)

## What you want to do

**Create a pivot table** from source data on a worksheet, place row and data (value) fields and aggregate. Also change the aggregation method (Sum → Average) and delete the pivot table.

| Step | Method |
| --- | --- |
| Add a pivot table | `ws.getPivotTables().addPivotTable(name, sourceRange, destCell)` |
| Get a field | `pt.getField(fieldName)` |
| Place fields | `pt.setFields(settings, layoutType, rowFields, colFields, filterFields, dataFields)` |
| Run the aggregation | `pt.refreshData()` → `pt.refresh()` |
| Change the aggregation method | `field.setDataFieldSubTotalsMethod(...)` → `setFields` again → `refreshData` → `refresh` |
| Delete the pivot table | `ws.getPivotTables().removePivotTable(name)` |

## Classes involved

### `PivotTables` (obtained via WorkSheet.getPivotTables())

| Method | Description |
| --- | --- |
| `addPivotTable(name, sourceRange, destCell)` | Adds a pivot table. sourceRange is an A1-style cell range; destCell is the top-left output cell (A1 style). |
| `getPivotTable(name)` | Gets a pivot table by name. Throws if it does not exist. |
| `removePivotTable(name)` | Deletes a pivot table. |

### `PivotTable`

| Method | Description |
| --- | --- |
| `getField(fieldName)` | Gets a field by its source-data header name (`dto.PivotFieldObject`). |
| `getFields()` | Returns an array of all fields. |
| `getPivotTableSetting()` | Returns the settings object (`dto.PivotTableSettingObject`). Used as the first argument of `setFields`. |
| `setFields(settings, layout, rows, cols, filters, data)` | Places fields and fixes the layout. |
| `refreshData()` | Reloads the source data and re-aggregates. |
| `refresh()` | Reflects field-layout changes and re-aggregates. |
| `getLayoutType()` | Returns the current layout type (`enums.XlPivotTableLayoutType`). |

### `dto.PivotFieldObject`

| Method | Description |
| --- | --- |
| `setFieldType(type)` | Sets the field's role (`enums.XlPivotFieldType`). |
| `setDataFieldSubTotalsMethod(method)` | Sets the aggregation method (`enums.XlPivotFieldSubtotalsMethodType`). Only valid for data fields. |
| `getCaption()` | Returns the field name. |

### Main enums

| enum | Value | Description |
| --- | --- | --- |
| `XlPivotFieldType.PivotFieldTypeRow` | 0 | Row field |
| `XlPivotFieldType.PivotFieldTypeCol` | 1 | Column field |
| `XlPivotFieldType.PivotFieldTypeData` | 3 | Data (value) field |
| `XlPivotTableLayoutType.TabularForm` | 2 | Tabular form |
| `XlPivotTableLayoutType.CompactForm` | 0 | Compact form |
| `XlPivotTableLayoutType.OutlineForm` | 1 | Outline form |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum` | 1 | Sum |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage` | 3 | Average |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeCount` | 2 | Count |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMax` | 4 | Max |
| `XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMin` | 5 | Min |

## Code

Output:

#### [Node.js]

```javascript
// pivot-sample.js — pivot tables: create, aggregate, change aggregation method, delete
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("pivot.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. Source data ----------
// 10 rows of Region / Product / Quantity
const headers = ["Region", "Product", "Quantity"];
const srcData = [
["East", "Apple", 10], ["East", "Orange", 20], ["West", "Apple", 30],
["West", "Orange", 15], ["East", "Apple", 25], ["North", "Orange", 40],
["North", "Apple", 5], ["West", "Apple", 10], ["East", "Orange", 30],
["North", "Apple", 20],
];
headers.forEach((h, i) => {
const col = String.fromCharCode(65 + i);
ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) => {
row.forEach((v, ci) => {
const col = String.fromCharCode(65 + ci);
const r = ri + 2;
if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
});
});
const srcTotal = srcData.reduce((s, r) => s + r[2], 0);
console.log(`[1] Source data: ${srcData.length} rows, quantity total = ${srcTotal}`);
// ---------- 2. Create the pivot table (rows=Region, values=Quantity sum) ----------
const pt = ws.getPivotTables().addPivotTable("SalesSummary", "A1:C11", "E1");
const rf = pt.getField("Region");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("Quantity");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)");
// ---------- 3. Read back the results and cross-check ----------
// Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
// The grand-total label ("総計") is written by the library itself, not by our data.
const pivotRows = {};
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
pivotRows[label] = Number(val);
console.log(` ${label} = ${val}`);
}
}
const grandTotal = pivotRows["総計"];
console.log(`[3] Check: grand total ${grandTotal} === source total ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);
// ---------- 4. Change the aggregation method to Average ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] Changed subtotal method from Sum to Average");
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
console.log(` ${label} = ${val}`);
}
}
// ---------- 5. Get field information ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} fields → ${fields.map(f => f.getCaption()).join(", ")}`);
console.log(` getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);
// ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary");
try {
ws.getPivotTables().getPivotTable("SalesSummary");
console.log("[6] Removal failed");
} catch (e) {
console.log("[6] removePivotTable → removal confirmed OK");
}
wb.save();
wb.close();
```

#### [Python]

```python
# pivot-sample.py — pivot tables: same verification as the Node.js sample (pyosbxl)
import pyosbxl
enums = pyosbxl.enums
app = pyosbxl.App()
wb = app.createWorkBook("pivot-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- 1. Source data ----------
# 10 rows of Region / Product / Quantity
headers = ["Region", "Product", "Quantity"]
src\_data = [
["East", "Apple", 10], ["East", "Orange", 20], ["West", "Apple", 30],
["West", "Orange", 15], ["East", "Apple", 25], ["North", "Orange", 40],
["North", "Apple", 5], ["West", "Apple", 10], ["East", "Orange", 30],
["North", "Apple", 20],
]
for i, h in enumerate(headers):
col = chr(65 + i)
ws.getRange(f"{col}1:{col}1").setValue(h)
for ri, row in enumerate(src\_data):
for ci, v in enumerate(row):
col = chr(65 + ci)
r = ri + 2
cell = ws.getRange(f"{col}{r}:{col}{r}")
if isinstance(v, int):
cell.setNumberValue(v)
else:
cell.setValue(v)
src\_total = sum(row[2] for row in src\_data)
print(f"[1] Source data: {len(src\_data)} rows, quantity total = {src\_total}")
# ---------- 2. Create the pivot table (rows=Region, values=Quantity sum) ----------
pt = ws.getPivotTables().addPivotTable("SalesSummary", "A1:C11", "E1")
rf = pt.getField("Region")
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow)
df = pt.getField("Quantity")
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData)
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)")
# ---------- 3. Read back the results and cross-check ----------
# Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
# The grand-total label ("総計") is written by the library itself, not by our data.
pivot\_rows = {}
for r in range(2, 11):
label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
val = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
if label is not None and label != "":
pivot\_rows[label] = float(val)
print(f" {label} = {val}")
grand\_total = pivot\_rows["総計"]
# To match node's Number display (205), print as int when the value is integral
gt = int(grand\_total) if grand\_total == int(grand\_total) else grand\_total
print(f"[3] Check: grand total {gt} === source total {src\_total} → {'OK' if grand\_total == src\_total else 'NG'}")
# ---------- 4. Change the aggregation method to Average ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[4] Changed subtotal method from Sum to Average")
for r in range(2, 11):
label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
val = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
if label is not None and label != "":
print(f" {label} = {val}")
# ---------- 5. Get field information ----------
fields = pt.getFields()
print(f"[5] getFields: {len(fields)} fields → " + ", ".join(f.getCaption() for f in fields))
print(f" getLayoutType = {int(pt.getLayoutType())} (TabularForm={int(enums.XlPivotTableLayoutType.TabularForm)})")
# ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary")
try:
ws.getPivotTables().getPivotTable("SalesSummary")
print("[6] Removal failed")
except Exception:
print("[6] removePivotTable → removal confirmed OK")
wb.save()
wb.close()
```

#### [Java]

```java
// pivot-sample.java — pivot tables: same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("pivot-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- 1. Source data ----------
// 10 rows of Region / Product / Quantity
String[] headers = {"Region", "Product", "Quantity"};
String[][] srcData = {
{"East", "Apple", "10"}, {"East", "Orange", "20"}, {"West", "Apple", "30"},
{"West", "Orange", "15"}, {"East", "Apple", "25"}, {"North", "Orange", "40"},
{"North", "Apple", "5"}, {"West", "Apple", "10"}, {"East", "Orange", "30"},
{"North", "Apple", "20"},
};
for (int i = 0; i < headers.length; i++) {
String col = String.valueOf((char) ('A' + i));
ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
}
int srcTotal = 0;
for (int ri = 0; ri < srcData.length; ri++) {
for (int ci = 0; ci < srcData[ri].length; ci++) {
String col = String.valueOf((char) ('A' + ci));
int r = ri + 2;
String addr = col + r + ":" + col + r;
if (ci == 2) { // the Quantity column is entered as numbers
int v = Integer.parseInt(srcData[ri][ci]);
ws.getRange(addr).setNumberValue(v);
srcTotal += v;
} else {
ws.getRange(addr).setValue(srcData[ri][ci]);
}
}
}
System.out.println("[1] Source data: " + srcData.length + " rows, quantity total = " + srcTotal);
// ---------- 2. Create the pivot table (rows=Region, values=Quantity sum) ----------
PivotTableWrapper pt = ws.getPivotTables().addPivotTable("SalesSummary", "A1:C11", "E1");
PivotFieldObjectWrapper rf = pt.getField("Region");
rf.setFieldType(XlPivotFieldType.PivotFieldTypeRow);
PivotFieldObjectWrapper df = pt.getField("Quantity");
df.setFieldType(XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);
List<PivotFieldObjectWrapper> empty = new ArrayList<>();
pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
Arrays.asList(rf), empty, empty, Arrays.asList(df));
pt.refreshData();
pt.refresh();
System.out.println("[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)");
// ---------- 3. Read back the results and cross-check ----------
// Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
// The grand-total label ("総計") is written by the library itself, not by our data.
Map<String, Double> pivotRows = new LinkedHashMap<>();
for (int r = 2; r <= 10; r++) {
String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
String val = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
if (label != null && !label.isEmpty()) {
pivotRows.put(label, Double.parseDouble(val));
System.out.println(" " + label + " = " + val);
}
}
double grandTotal = pivotRows.get("総計");
// To match node's Number display (205), print as long when the value is integral
String gt = (grandTotal == Math.rint(grandTotal))
? String.valueOf((long) grandTotal) : String.valueOf(grandTotal);
System.out.println("[3] Check: grand total " + gt + " === source total " + srcTotal
+ " → " + (grandTotal == srcTotal ? "OK" : "NG"));
// ---------- 4. Change the aggregation method to Average ----------
df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
Arrays.asList(rf), empty, empty, Arrays.asList(df));
pt.refreshData();
pt.refresh();
System.out.println("[4] Changed subtotal method from Sum to Average");
for (int r = 2; r <= 10; r++) {
String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
String val = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
if (label != null && !label.isEmpty()) {
System.out.println(" " + label + " = " + val);
}
}
// ---------- 5. Get field information ----------
List<PivotFieldObjectWrapper> fields = pt.getFields();
StringBuilder caps = new StringBuilder();
for (int i = 0; i < fields.size(); i++) {
if (i > 0) caps.append(", ");
caps.append(fields.get(i).getCaption());
}
System.out.println("[5] getFields: " + fields.size() + " fields → " + caps);
System.out.println(" getLayoutType = " + pt.getLayoutType().getCode()
+ " (TabularForm=" + XlPivotTableLayoutType.TabularForm.getCode() + ")");
// ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary");
try {
ws.getPivotTables().getPivotTable("SalesSummary");
System.out.println("[6] Removal failed");
} catch (Throwable t) {
System.out.println("[6] removePivotTable → removal confirmed OK");
}
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// pivot-sample.md — pivot tables: same verification as the Node.js sample (phposbxl). Body only (<?php / autoload provided by the harness)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("pivot-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// ---------- 1. Source data ----------
// 10 rows of Region / Product / Quantity
$headers = ["Region", "Product", "Quantity"];
$srcData = [
["East", "Apple", 10], ["East", "Orange", 20], ["West", "Apple", 30],
["West", "Orange", 15], ["East", "Apple", 25], ["North", "Orange", 40],
["North", "Apple", 5], ["West", "Apple", 10], ["East", "Orange", 30],
["North", "Apple", 20],
];
foreach ($headers as $i => $h) {
$col = chr(65 + $i);
$ws->getRange("{$col}1:{$col}1")->setValue($h);
}
$srcTotal = 0;
foreach ($srcData as $ri => $row) {
foreach ($row as $ci => $v) {
$col = chr(65 + $ci);
$r = $ri + 2;
if (is\_int($v)) $ws->getRange("{$col}{$r}:{$col}{$r}")->setNumberValue($v);
else $ws->getRange("{$col}{$r}:{$col}{$r}")->setValue($v);
}
$srcTotal += $row[2];
}
echo "[1] Source data: " . count($srcData) . " rows, quantity total = {$srcTotal}\n";
// ---------- 2. Create the pivot table (rows=Region, values=Quantity sum) ----------
$pt = $ws->getPivotTables()->addPivotTable("SalesSummary", "A1:C11", "E1");
$rf = $pt->getField("Region");
$rf->setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeRow);
$df = $pt->getField("Quantity");
$df->setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeData);
$df->setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeSum);
$pt->setFields($pt->getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
[$rf], [], [], [$df]);
$pt->refreshData();
$pt->refresh();
echo "[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)\n";
// ---------- 3. Read back the results and cross-check ----------
// Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
// The grand-total label ("総計") is written by the library itself, not by our data.
// getValue() returns an associative array (key=cell address, value=displayed value as string).
$pivotRows = [];
for ($r = 2; $r <= 10; $r++) {
$label = $ws->getRange("E{$r}:E{$r}")->getValue()["E{$r}"] ?? null;
$val = $ws->getRange("F{$r}:F{$r}")->getValue()["F{$r}"] ?? null;
if ($label !== null && $label !== "") {
$pivotRows[$label] = (float)$val;
echo " {$label} = {$val}\n";
}
}
$grandTotal = $pivotRows["総計"];
// print like a JS Number (205): use int representation when the value is integral
$gt = ($grandTotal == (int)$grandTotal) ? (string)(int)$grandTotal : (string)$grandTotal;
echo "[3] Check: grand total {$gt} === source total {$srcTotal} → " . ($grandTotal == $srcTotal ? "OK" : "NG") . "\n";
// ---------- 4. Change the aggregation method to Average ----------
$df->setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeAverage);
$pt->setFields($pt->getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
[$rf], [], [], [$df]);
$pt->refreshData();
$pt->refresh();
echo "[4] Changed subtotal method from Sum to Average\n";
for ($r = 2; $r <= 10; $r++) {
$label = $ws->getRange("E{$r}:E{$r}")->getValue()["E{$r}"] ?? null;
$val = $ws->getRange("F{$r}:F{$r}")->getValue()["F{$r}"] ?? null;
if ($label !== null && $label !== "") {
echo " {$label} = {$val}\n";
}
}
// ---------- 5. Get field information ----------
$fields = $pt->getFields();
$caps = array\_map(fn($f) => $f->getCaption(), $fields);
echo "[5] getFields: " . count($fields) . " fields → " . implode(", ", $caps) . "\n";
$tabular = \phposbxl\enums\XlPivotTableLayoutType::TabularForm->value;
echo " getLayoutType = {$pt->getLayoutType()} (TabularForm={$tabular})\n";
// ---------- 6. Delete the pivot table ----------
$ws->getPivotTables()->removePivotTable("SalesSummary");
try {
$ws->getPivotTables()->getPivotTable("SalesSummary");
echo "[6] Removal failed\n";
} catch (\Throwable $e) {
echo "[6] removePivotTable → removal confirmed OK\n";
}
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// pivot-sample.wasm.js — pivot tables: same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("pivot-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- 1. Source data ----------
// 10 rows of Region / Product / Quantity
const headers = ["Region", "Product", "Quantity"];
const srcData = [
["East", "Apple", 10], ["East", "Orange", 20], ["West", "Apple", 30],
["West", "Orange", 15], ["East", "Apple", 25], ["North", "Orange", 40],
["North", "Apple", 5], ["West", "Apple", 10], ["East", "Orange", 30],
["North", "Apple", 20],
];
headers.forEach((h, i) => {
const col = String.fromCharCode(65 + i);
ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) => {
row.forEach((v, ci) => {
const col = String.fromCharCode(65 + ci);
const r = ri + 2;
if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
});
});
const srcTotal = srcData.reduce((s, r) => s + r[2], 0);
console.log(`[1] Source data: ${srcData.length} rows, quantity total = ${srcTotal}`);
// ---------- 2. Create the pivot table (rows=Region, values=Quantity sum) ----------
const pt = ws.getPivotTables().addPivotTable("SalesSummary", "A1:C11", "E1");
const rf = pt.getField("Region");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("Quantity");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)");
// ---------- 3. Read back the results and cross-check ----------
// Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
// The grand-total label ("総計") is written by the library itself, not by our data.
const pivotRows = {};
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
pivotRows[label] = Number(val);
console.log(` ${label} = ${val}`);
}
}
const grandTotal = pivotRows["総計"];
console.log(`[3] Check: grand total ${grandTotal} === source total ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);
// ---------- 4. Change the aggregation method to Average ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] Changed subtotal method from Sum to Average");
for (let r = 2; r <= 10; r++) {
const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
const val = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
if (label !== undefined && label !== "") {
console.log(` ${label} = ${val}`);
}
}
// ---------- 5. Get field information ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} fields → ${fields.map(f => f.getCaption()).join(", ")}`);
console.log(` getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);
// ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary");
try {
ws.getPivotTables().getPivotTable("SalesSummary");
console.log("[6] Removal failed");
} catch (e) {
console.log("[6] removePivotTable → removal confirmed OK");
}
wb.save();
wb.close();
```

```plaintext
[1] Source data: 10 rows, quantity total = 205
[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)
East = 85
North = 65
West = 55
総計 = 205
[3] Check: grand total 205 === source total 205 → OK
[4] Changed subtotal method from Sum to Average
East = 21.25
North = 21.66666667
West = 18.33333333
総計 = 20.5
[5] getFields: 3 fields → Region, Product, Quantity
getLayoutType = 2 (TabularForm=2)
[6] removePivotTable → removal confirmed OK
```

## Notes

### Flow for creating a pivot table

1. **Prepare the source data** — headers (field names) in row 1, data from row 2 onward.
2. **`addPivotTable(name, sourceRange, destCell)`** — specify the source range (A1 style) and the top-left output cell.

When placing the pivot on the same sheet, choose a `destCell` that does not overlap the source. When placing it on another sheet, prefix `sourceRange` with the sheet name, e.g. `Sheet1!A1:C11`.

1. **Get fields and set their roles** — get a field with `pt.getField("header name")` and assign

row / column / data / filter with `setFieldType()`.

1. **Place them with `setFields()`** — pass the settings object, the layout type and the field arrays.

The argument order is `(settings, layoutType, rowFields, colFields, filterFields, dataFields)`.

1. **`refreshData()` → `refresh()`** — reloads the source data and writes the aggregation results to the sheet.

**`setFields()` alone does not produce results.** Always call `refresh()`.

### Changing the aggregation method

The aggregation method of a data field is changed with `setDataFieldSubTotalsMethod()`. After changing it you must run **`setFields()` → `refreshData()` → `refresh()`** again.

| enum value | Number | Aggregation method |
| --- | --- | --- |
| `SubtotalsMethodTypeSum` | 1 | Sum (default) |
| `SubtotalsMethodTypeCount` | 2 | Count |
| `SubtotalsMethodTypeAverage` | 3 | Average |
| `SubtotalsMethodTypeMax` | 4 | Max |
| `SubtotalsMethodTypeMin` | 5 | Min |
| `SubtotalsMethodTypeProduct` | 6 | Product |
| `SubtotalsMethodTypeCountNums` | 7 | Count numbers |
| `SubtotalsMethodTypeStdDev` | 8 | Sample standard deviation |
| `SubtotalsMethodTypeStdDevP` | 9 | Standard deviation |
| `SubtotalsMethodTypeVar` | 10 | Sample variance |
| `SubtotalsMethodTypeVarP` | 11 | Variance |

### Row × column cross-tabulation

To also use a column field, pass an array as `colFields` of `setFields()`.

The output becomes a matrix of row labels (column E) × column labels (column F onward).

#### [Node.js]

```javascript
const cf = pt.getField("Product");
cf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeCol);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
[rf], [cf], [], [df]);
pt.refreshData();
pt.refresh();
```

### Reading back the aggregated results

Pivot table output can be read as ordinary cell values. Get aggregated values with e.g. `ws.getRange("F3:F3").getValue()` and cross-check them by comparing against the sum of the source data.

### Deleting the pivot table

Delete with `removePivotTable(name)`. Calling `getPivotTable(name)` afterwards throws.

## Caveats

* **Always call `refresh()` after `setFields()`.** `setFields()` only registers the layout; writing to the sheet is done by `refresh()`. If you changed the source data, call `refreshData()` first.
* **A pivot without a column field (rows + values only) also works.** Pass an empty array `[]` for `colFields`.
* **`getFields()` returns every header of the source data.** It includes fields you never placed with `setFields()`, so if you want only the placed fields, check each one with `getFieldType()`.
* **Include the header row in the source range of `addPivotTable()`.** Specify a range whose first row holds the headers, e.g. `A1:C11`. The header names become the field names as-is.
* **The "grand total" row of the results is labeled by the library — currently `"総計"` (Japanese for "Grand Total").** Since the label string may depend on the environment/build, when reading results back either access rows by row number or search for the label instead of hardcoding a translated word (this sample looks up the exact label the library writes).
* **To build multiple pivot tables over the same source range,** you can place several on one sheet as long as the names (`pivotTableName`) differ.  
  Make sure their output cells do not overlap.
* **`removePivotTable()` does not erase the cell values the pivot wrote.** To clear them too, use `clearCell()` or similar separately.

## See also

* [Index](index.md)
* Previous: [Charts Basics](general-10-charts.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
