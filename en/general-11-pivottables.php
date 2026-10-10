<?php
$lang = 'en';
$page_title = 'Pivot Tables — create, aggregate, delete | osbxl Samples';
$no_tagline = true;
$meta_description = 'Pivot Tables — create, aggregate, delete. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,pivottable,pivottables,addpivottable,setfields,refreshdata,refresh,removepivottable,getfield,getfields,pivotfieldobject,subtotalsmethodtype,layouttype';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Pivot Tables — create, aggregate, delete</h1>
<p class="text-sm text-text-light mb-6">level: advanced ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-11-pivottables.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7"><strong>Create a pivot table</strong> from source data on a worksheet, place row and data (value) fields and aggregate. Also change the aggregation method (Sum → Average) and delete the pivot table.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Step</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Add a pivot table</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getPivotTables().addPivotTable(name, sourceRange, destCell)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Get a field</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.getField(fieldName)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Place fields</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.setFields(settings, layoutType, rowFields, colFields, filterFields, dataFields)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Run the aggregation</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.refreshData()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.refresh()</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Change the aggregation method</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">field.setDataFieldSubTotalsMethod(...)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields</code> again → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Delete the pivot table</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getPivotTables().removePivotTable(name)</code></td></tr>
  </tbody></table></div>
<h2 id="classes-involved" class="text-2xl font-bold mt-10 mb-4">Classes involved</h2>
<h3 id="pivottables-obtained-via-worksheet-getpivottables" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">PivotTables</code> (obtained via WorkSheet.getPivotTables())</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotTable(name, sourceRange, destCell)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Adds a pivot table. sourceRange is an A1-style cell range; destCell is the top-left output cell (A1 style).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTable(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a pivot table by name. Throws if it does not exist.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removePivotTable(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Deletes a pivot table.</td></tr>
  </tbody></table></div>
<h3 id="pivottable" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">PivotTable</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getField(fieldName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a field by its source-data header name (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.PivotFieldObject</code>).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFields()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns an array of all fields.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTableSetting()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns the settings object (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.PivotTableSettingObject</code>). Used as the first argument of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields</code>.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields(settings, layout, rows, cols, filters, data)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Places fields and fixes the layout.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Reloads the source data and re-aggregates.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Reflects field-layout changes and re-aggregates.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getLayoutType()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns the current layout type (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlPivotTableLayoutType</code>).</td></tr>
  </tbody></table></div>
<h3 id="dto-pivotfieldobject" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.PivotFieldObject</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFieldType(type)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets the field's role (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlPivotFieldType</code>).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataFieldSubTotalsMethod(method)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets the aggregation method (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlPivotFieldSubtotalsMethodType</code>). Only valid for data fields.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCaption()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns the field name.</td></tr>
  </tbody></table></div>
<h3 id="main-enums" class="text-xl font-bold mt-8 mb-3">Main enums</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">enum</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Value</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldType.PivotFieldTypeRow</code></td><td class="border border-surface-200 px-3 py-2 align-top">0</td><td class="border border-surface-200 px-3 py-2 align-top">Row field</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldType.PivotFieldTypeCol</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">Column field</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldType.PivotFieldTypeData</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">Data (value) field</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotTableLayoutType.TabularForm</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">Tabular form</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotTableLayoutType.CompactForm</code></td><td class="border border-surface-200 px-3 py-2 align-top">0</td><td class="border border-surface-200 px-3 py-2 align-top">Compact form</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotTableLayoutType.OutlineForm</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">Outline form</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">Sum</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">Average</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeCount</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">Count</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMax</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">Max</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMin</code></td><td class="border border-surface-200 px-3 py-2 align-top">5</td><td class="border border-surface-200 px-3 py-2 align-top">Min</td></tr>
  </tbody></table></div>
<h2 id="code" class="text-2xl font-bold mt-10 mb-4">Code</h2>
<p class="mb-4 text-slate-700 leading-7">Output:</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
  </div>
    <div class="tab-content" id="tab-g0-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// pivot-sample.js — pivot tables: create, aggregate, change aggregation method, delete
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
  ["North", "Apple",  5], ["West", "Apple", 10], ["East", "Orange", 30],
  ["North", "Apple", 20],
];
headers.forEach((h, i) =&gt; {
  const col = String.fromCharCode(65 + i);
  ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) =&gt; {
  row.forEach((v, ci) =&gt; {
    const col = String.fromCharCode(65 + ci);
    const r = ri + 2;
    if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
    else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
  });
});
const srcTotal = srcData.reduce((s, r) =&gt; s + r[2], 0);
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
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    pivotRows[label] = Number(val);
    console.log(`    ${label} = ${val}`);
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
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    console.log(`    ${label} = ${val}`);
  }
}

// ---------- 5. Get field information ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} fields → ${fields.map(f =&gt; f.getCaption()).join(", ")}`);
console.log(`    getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);

// ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary");
try {
  ws.getPivotTables().getPivotTable("SalesSummary");
  console.log("[6] Removal failed");
} catch (e) {
  console.log("[6] removePivotTable → removal confirmed OK");
}

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># pivot-sample.py — pivot tables: same verification as the Node.js sample (pyosbxl)
import pyosbxl

enums = pyosbxl.enums

app = pyosbxl.App()
wb = app.createWorkBook("pivot-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. Source data ----------
# 10 rows of Region / Product / Quantity
headers = ["Region", "Product", "Quantity"]
src_data = [
    ["East", "Apple", 10], ["East", "Orange", 20], ["West", "Apple", 30],
    ["West", "Orange", 15], ["East", "Apple", 25], ["North", "Orange", 40],
    ["North", "Apple",  5], ["West", "Apple", 10], ["East", "Orange", 30],
    ["North", "Apple", 20],
]
for i, h in enumerate(headers):
    col = chr(65 + i)
    ws.getRange(f"{col}1:{col}1").setValue(h)
for ri, row in enumerate(src_data):
    for ci, v in enumerate(row):
        col = chr(65 + ci)
        r = ri + 2
        cell = ws.getRange(f"{col}{r}:{col}{r}")
        if isinstance(v, int):
            cell.setNumberValue(v)
        else:
            cell.setValue(v)
src_total = sum(row[2] for row in src_data)
print(f"[1] Source data: {len(src_data)} rows, quantity total = {src_total}")

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
pivot_rows = {}
for r in range(2, 11):
    label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
    val   = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
    if label is not None and label != "":
        pivot_rows[label] = float(val)
        print(f"    {label} = {val}")
grand_total = pivot_rows["総計"]
# To match node's Number display (205), print as int when the value is integral
gt = int(grand_total) if grand_total == int(grand_total) else grand_total
print(f"[3] Check: grand total {gt} === source total {src_total} → {'OK' if grand_total == src_total else 'NG'}")

# ---------- 4. Change the aggregation method to Average ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[4] Changed subtotal method from Sum to Average")
for r in range(2, 11):
    label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
    val   = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
    if label is not None and label != "":
        print(f"    {label} = {val}")

# ---------- 5. Get field information ----------
fields = pt.getFields()
print(f"[5] getFields: {len(fields)} fields → " + ", ".join(f.getCaption() for f in fields))
print(f"    getLayoutType = {int(pt.getLayoutType())} (TabularForm={int(enums.XlPivotTableLayoutType.TabularForm)})")

# ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary")
try:
    ws.getPivotTables().getPivotTable("SalesSummary")
    print("[6] Removal failed")
except Exception:
    print("[6] removePivotTable → removal confirmed OK")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// pivot-sample.java — pivot tables: same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
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
            {"North", "Apple",  "5"}, {"West", "Apple", "10"}, {"East", "Orange", "30"},
            {"North", "Apple", "20"},
        };
        for (int i = 0; i &lt; headers.length; i++) {
            String col = String.valueOf((char) ('A' + i));
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        int srcTotal = 0;
        for (int ri = 0; ri &lt; srcData.length; ri++) {
            for (int ci = 0; ci &lt; srcData[ri].length; ci++) {
                String col = String.valueOf((char) ('A' + ci));
                int r = ri + 2;
                String addr = col + r + ":" + col + r;
                if (ci == 2) {   // the Quantity column is entered as numbers
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

        List&lt;PivotFieldObjectWrapper&gt; empty = new ArrayList&lt;&gt;();
        pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
                     Arrays.asList(rf), empty, empty, Arrays.asList(df));
        pt.refreshData();
        pt.refresh();
        System.out.println("[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)");

        // ---------- 3. Read back the results and cross-check ----------
        // Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
        // The grand-total label ("総計") is written by the library itself, not by our data.
        Map&lt;String, Double&gt; pivotRows = new LinkedHashMap&lt;&gt;();
        for (int r = 2; r &lt;= 10; r++) {
            String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
            String val   = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
            if (label != null &amp;&amp; !label.isEmpty()) {
                pivotRows.put(label, Double.parseDouble(val));
                System.out.println("    " + label + " = " + val);
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
        for (int r = 2; r &lt;= 10; r++) {
            String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
            String val   = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
            if (label != null &amp;&amp; !label.isEmpty()) {
                System.out.println("    " + label + " = " + val);
            }
        }

        // ---------- 5. Get field information ----------
        List&lt;PivotFieldObjectWrapper&gt; fields = pt.getFields();
        StringBuilder caps = new StringBuilder();
        for (int i = 0; i &lt; fields.size(); i++) {
            if (i &gt; 0) caps.append(", ");
            caps.append(fields.get(i).getCaption());
        }
        System.out.println("[5] getFields: " + fields.size() + " fields → " + caps);
        System.out.println("    getLayoutType = " + pt.getLayoutType().getCode()
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
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// pivot-sample.php — pivot tables: same verification as the Node.js sample (phposbxl). Body only (&lt;?php / autoload provided by the harness)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("pivot-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// ---------- 1. Source data ----------
// 10 rows of Region / Product / Quantity
$headers = ["Region", "Product", "Quantity"];
$srcData = [
  ["East", "Apple", 10], ["East", "Orange", 20], ["West", "Apple", 30],
  ["West", "Orange", 15], ["East", "Apple", 25], ["North", "Orange", 40],
  ["North", "Apple",  5], ["West", "Apple", 10], ["East", "Orange", 30],
  ["North", "Apple", 20],
];
foreach ($headers as $i =&gt; $h) {
  $col = chr(65 + $i);
  $ws-&gt;getRange("{$col}1:{$col}1")-&gt;setValue($h);
}
$srcTotal = 0;
foreach ($srcData as $ri =&gt; $row) {
  foreach ($row as $ci =&gt; $v) {
    $col = chr(65 + $ci);
    $r = $ri + 2;
    if (is_int($v)) $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;setNumberValue($v);
    else $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;setValue($v);
  }
  $srcTotal += $row[2];
}
echo "[1] Source data: " . count($srcData) . " rows, quantity total = {$srcTotal}\n";

// ---------- 2. Create the pivot table (rows=Region, values=Quantity sum) ----------
$pt = $ws-&gt;getPivotTables()-&gt;addPivotTable("SalesSummary", "A1:C11", "E1");

$rf = $pt-&gt;getField("Region");
$rf-&gt;setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeRow);
$df = $pt-&gt;getField("Quantity");
$df-&gt;setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeData);
$df-&gt;setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeSum);

$pt-&gt;setFields($pt-&gt;getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
               [$rf], [], [], [$df]);
$pt-&gt;refreshData();
$pt-&gt;refresh();
echo "[2] Created PivotTable (rows=Region, values=Quantity sum, layout=tabular form)\n";

// ---------- 3. Read back the results and cross-check ----------
// Pivot output: E1=header, E2 onward=row labels, F2 onward=values, last row=grand total.
// The grand-total label ("総計") is written by the library itself, not by our data.
// getValue() returns an associative array (key=cell address, value=displayed value as string).
$pivotRows = [];
for ($r = 2; $r &lt;= 10; $r++) {
  $label = $ws-&gt;getRange("E{$r}:E{$r}")-&gt;getValue()["E{$r}"] ?? null;
  $val   = $ws-&gt;getRange("F{$r}:F{$r}")-&gt;getValue()["F{$r}"] ?? null;
  if ($label !== null &amp;&amp; $label !== "") {
    $pivotRows[$label] = (float)$val;
    echo "    {$label} = {$val}\n";
  }
}
$grandTotal = $pivotRows["総計"];
// print like a JS Number (205): use int representation when the value is integral
$gt = ($grandTotal == (int)$grandTotal) ? (string)(int)$grandTotal : (string)$grandTotal;
echo "[3] Check: grand total {$gt} === source total {$srcTotal} → " . ($grandTotal == $srcTotal ? "OK" : "NG") . "\n";

// ---------- 4. Change the aggregation method to Average ----------
$df-&gt;setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeAverage);
$pt-&gt;setFields($pt-&gt;getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
               [$rf], [], [], [$df]);
$pt-&gt;refreshData();
$pt-&gt;refresh();
echo "[4] Changed subtotal method from Sum to Average\n";
for ($r = 2; $r &lt;= 10; $r++) {
  $label = $ws-&gt;getRange("E{$r}:E{$r}")-&gt;getValue()["E{$r}"] ?? null;
  $val   = $ws-&gt;getRange("F{$r}:F{$r}")-&gt;getValue()["F{$r}"] ?? null;
  if ($label !== null &amp;&amp; $label !== "") {
    echo "    {$label} = {$val}\n";
  }
}

// ---------- 5. Get field information ----------
$fields = $pt-&gt;getFields();
$caps = array_map(fn($f) =&gt; $f-&gt;getCaption(), $fields);
echo "[5] getFields: " . count($fields) . " fields → " . implode(", ", $caps) . "\n";
$tabular = \phposbxl\enums\XlPivotTableLayoutType::TabularForm-&gt;value;
echo "    getLayoutType = {$pt-&gt;getLayoutType()} (TabularForm={$tabular})\n";

// ---------- 6. Delete the pivot table ----------
$ws-&gt;getPivotTables()-&gt;removePivotTable("SalesSummary");
try {
  $ws-&gt;getPivotTables()-&gt;getPivotTable("SalesSummary");
  echo "[6] Removal failed\n";
} catch (\Throwable $e) {
  echo "[6] removePivotTable → removal confirmed OK\n";
}

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// pivot-sample.wasm.js — pivot tables: same verification as the Node.js sample (wasmosbxl / embind)
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
  ["North", "Apple",  5], ["West", "Apple", 10], ["East", "Orange", 30],
  ["North", "Apple", 20],
];
headers.forEach((h, i) =&gt; {
  const col = String.fromCharCode(65 + i);
  ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) =&gt; {
  row.forEach((v, ci) =&gt; {
    const col = String.fromCharCode(65 + ci);
    const r = ri + 2;
    if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
    else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
  });
});
const srcTotal = srcData.reduce((s, r) =&gt; s + r[2], 0);
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
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    pivotRows[label] = Number(val);
    console.log(`    ${label} = ${val}`);
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
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    console.log(`    ${label} = ${val}`);
  }
}

// ---------- 5. Get field information ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} fields → ${fields.map(f =&gt; f.getCaption()).join(", ")}`);
console.log(`    getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);

// ---------- 6. Delete the pivot table ----------
ws.getPivotTables().removePivotTable("SalesSummary");
try {
  ws.getPivotTables().getPivotTable("SalesSummary");
  console.log("[6] Removal failed");
} catch (e) {
  console.log("[6] removePivotTable → removal confirmed OK");
}

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] Source data: 10 rows, quantity total = 205
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
[6] removePivotTable → removal confirmed OK</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<h3 id="flow-for-creating-a-pivot-table" class="text-xl font-bold mt-8 mb-3">Flow for creating a pivot table</h3>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>Prepare the source data</strong> — headers (field names) in row 1, data from row 2 onward.</li>
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotTable(name, sourceRange, destCell)</code></strong> — specify the source range (A1 style) and the top-left output cell.</li>
</ol>
<p class="mb-4 text-slate-700 leading-7">When placing the pivot on the same sheet, choose a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">destCell</code> that does not overlap the source. When placing it on another sheet, prefix <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">sourceRange</code> with the sheet name, e.g. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet1!A1:C11</code>.</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>Get fields and set their roles</strong> — get a field with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.getField("header name")</code> and assign</li>
</ol>
<p class="mb-4 text-slate-700 leading-7">row / column / data / filter with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFieldType()</code>.</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>Place them with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code></strong> — pass the settings object, the layout type and the field arrays.</li>
</ol>
<p class="mb-4 text-slate-700 leading-7">The argument order is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(settings, layoutType, rowFields, colFields, filterFields, dataFields)</code>.</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code></strong> — reloads the source data and writes the aggregation results to the sheet.</li>
</ol>
<p class="mb-4 text-slate-700 leading-7"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> alone does not produce results.</strong> Always call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code>.</p>
<h3 id="changing-the-aggregation-method" class="text-xl font-bold mt-8 mb-3">Changing the aggregation method</h3>
<p class="mb-4 text-slate-700 leading-7">The aggregation method of a data field is changed with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataFieldSubTotalsMethod()</code>. After changing it you must run <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code></strong> again.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">enum value</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Number</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Aggregation method</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeSum</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">Sum (default)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeCount</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">Count</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeAverage</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">Average</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeMax</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">Max</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeMin</code></td><td class="border border-surface-200 px-3 py-2 align-top">5</td><td class="border border-surface-200 px-3 py-2 align-top">Min</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeProduct</code></td><td class="border border-surface-200 px-3 py-2 align-top">6</td><td class="border border-surface-200 px-3 py-2 align-top">Product</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeCountNums</code></td><td class="border border-surface-200 px-3 py-2 align-top">7</td><td class="border border-surface-200 px-3 py-2 align-top">Count numbers</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeStdDev</code></td><td class="border border-surface-200 px-3 py-2 align-top">8</td><td class="border border-surface-200 px-3 py-2 align-top">Sample standard deviation</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeStdDevP</code></td><td class="border border-surface-200 px-3 py-2 align-top">9</td><td class="border border-surface-200 px-3 py-2 align-top">Standard deviation</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeVar</code></td><td class="border border-surface-200 px-3 py-2 align-top">10</td><td class="border border-surface-200 px-3 py-2 align-top">Sample variance</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeVarP</code></td><td class="border border-surface-200 px-3 py-2 align-top">11</td><td class="border border-surface-200 px-3 py-2 align-top">Variance</td></tr>
  </tbody></table></div>
<h3 id="row-column-cross-tabulation" class="text-xl font-bold mt-8 mb-3">Row × column cross-tabulation</h3>
<p class="mb-4 text-slate-700 leading-7">To also use a column field, pass an array as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colFields</code> of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code>.</p>
<p class="mb-4 text-slate-700 leading-7">The output becomes a matrix of row labels (column E) × column labels (column F onward).</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
  </div>
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">const cf = pt.getField("Product");
cf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeCol);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [cf], [], [df]);
pt.refreshData();
pt.refresh();</code></pre>
      </div>
    </div>
</div>
<h3 id="reading-back-the-aggregated-results" class="text-xl font-bold mt-8 mb-3">Reading back the aggregated results</h3>
<p class="mb-4 text-slate-700 leading-7">Pivot table output can be read as ordinary cell values. Get aggregated values with e.g. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("F3:F3").getValue()</code> and cross-check them by comparing against the sum of the source data.</p>
<h3 id="deleting-the-pivot-table" class="text-xl font-bold mt-8 mb-3">Deleting the pivot table</h3>
<p class="mb-4 text-slate-700 leading-7">Delete with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removePivotTable(name)</code>. Calling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTable(name)</code> afterwards throws.</p>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Always call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code> after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code>.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> only registers the layout; writing to the sheet is done by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code>. If you changed the source data, call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code> first.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>A pivot without a column field (rows + values only) also works.</strong> Pass an empty array <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[]</code> for <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colFields</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFields()</code> returns every header of the source data.</strong> It includes fields you never placed with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code>, so if you want only the placed fields, check each one with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFieldType()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Include the header row in the source range of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotTable()</code>.</strong> Specify a range whose first row holds the headers, e.g. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A1:C11</code>. The header names become the field names as-is.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The "grand total" row of the results is labeled by the library — currently <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"総計"</code> (Japanese for "Grand Total").</strong> Since the label string may depend on the environment/build, when reading results back either access rows by row number or search for the label instead of hardcoding a translated word (this sample looks up the exact label the library writes).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>To build multiple pivot tables over the same source range,</strong> you can place several on one sheet as long as the names (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pivotTableName</code>) differ.<br>Make sure their output cells do not overlap.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removePivotTable()</code> does not erase the cell values the pivot wrote.</strong> To clear them too, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCell()</code> or similar separately.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-10-charts.php" class="text-primary-600 hover:underline">Charts Basics</a></li>
<li><div class="mt-2 mb-2">
  <div class="text-sm font-bold mb-2">API reference</div>
  <div class="tab-group">
    <div class="flex gap-2 mb-3 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
    </div>
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/en/nodejs/classes/dto.PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/en/nodejs/enums/enums.XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/en/nodejs/enums/enums.XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/en/python/pyosbxl.PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/en/python/pyosbxl.dto.PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/PivotTableWrapper.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/PivotTablesWrapper.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/PivotFieldObjectWrapper.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/en/php/classes/phposbxl-PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/en/wasm/classes/osbxl.PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/en/wasm/classes/dto.PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/en/wasm/enums/enums.XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/en/wasm/enums/enums.XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Pivot Tables — create, aggregate, delete",
   "description": "Pivot Tables — create, aggregate, delete. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-11-pivottables.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-11-pivottables.php",
   "publisher": {
    "@type": "Organization",
    "name": "osb-office.inc",
    "url": "https://osb-office.com/",
    "logo": "https://osb-office.com/og-image.png"
   },
   "articleSection": "general"
  },
  {
   "@type": "BreadcrumbList",
   "itemListElement": [
    {
     "@type": "ListItem",
     "position": 1,
     "name": "Home",
     "item": "https://osb-office.com/en/"
    },
    {
     "@type": "ListItem",
     "position": 2,
     "name": "Sample Collection",
     "item": "https://osb-office.com/en/sample/"
    },
    {
     "@type": "ListItem",
     "position": 3,
     "name": "Pivot Tables — create, aggregate, delete",
     "item": "https://osb-office.com/en/sample/general-11-pivottables.php"
    }
   ]
  }
 ]
}
</script>

</main>

<script>
document.addEventListener('DOMContentLoaded', function() {
  hljs.highlightAll();
  var saved = null;
  try {
    saved = new URLSearchParams(location.search).get('lang') || localStorage.getItem('osbxl-sample-lang');
  } catch (e) {}
  document.querySelectorAll('.tab-group').forEach(function(g) {
    var btns = g.querySelectorAll('.tab-btn');
    var conts = g.querySelectorAll('.tab-content');
    function activate(lang) {
      var has = false;
      conts.forEach(function(c) { if (c.dataset.lang === lang) has = true; });
      if (!has && btns.length) lang = btns[0].dataset.tab;
      btns.forEach(function(b) {
        var on = b.dataset.tab === lang;
        b.classList.toggle('active', on);
        b.classList.toggle('border-primary-600', on);
        b.classList.toggle('text-primary-600', on);
        b.classList.toggle('border-transparent', !on);
        b.classList.toggle('text-text-light', !on);
      });
      conts.forEach(function(c) { c.classList.toggle('hidden', c.dataset.lang !== lang); });
    }
    btns.forEach(function(b) {
      b.addEventListener('click', function() {
        activate(b.dataset.tab);
        try { localStorage.setItem('osbxl-sample-lang', b.dataset.tab); } catch (e) {}
      });
    });
    activate(saved || (btns.length ? btns[0].dataset.tab : 'node'));
  });
});
</script>

<?php include __DIR__ . '/../../includes/footer.php'; ?>
