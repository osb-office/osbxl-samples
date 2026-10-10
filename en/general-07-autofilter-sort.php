<?php
$lang = 'en';
$page_title = 'AutoFilter and Sorting — filter, sort, read back | osbxl Samples';
$no_tagline = true;
$meta_description = 'AutoFilter and Sorting — filter, sort, read back. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,autofilter,sort,setcustomfilter,settop10valuefilter,executeSortAscending,executeSortDescending,getSortConditions,hidden,filter';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">AutoFilter and Sorting — filter, sort, read back</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-07-autofilter-sort.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Set an <strong>AutoFilter</strong> on a data table so that only rows matching a condition are shown, and sort (ascending / descending) with <strong>Sort</strong>. A filter works by <strong>hiding the rows</strong> that do not match the condition; whether a row is hidden can be read with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row.getHidden()</code>.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Approach</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">What to use</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Set filter / conditions</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getAutoFilter(range, firstRowAsHeader)</code> → the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">set*Filter</code> methods on <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">AutoFilter</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Clear conditions</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.resetAllFilter()</code> (conditions only) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.removeFilter()</code> (removes the filter itself)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Sort</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort(range)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortDescending</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Check hidden rows</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRow(r).getHidden()</code></td></tr>
  </tbody></table></div>
<h2 id="classes-covered" class="text-2xl font-bold mt-10 mb-4">Classes covered</h2>
<h3 id="nodeosbxl-autofilter-obtained-via-ws-getautofilter-a1c1-firstrowasheader" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl.AutoFilter</code> — obtained via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getAutoFilter(A1C1, firstRowAsHeader?)</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAddress()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns the target cell range.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(op, criteria, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Filter with an operator + criteria. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">criteria</code> supports the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">*</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">?</code> wildcards.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAndCustomFilter(op1, c1, op2, c2, colId?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setOrCustomFilter(...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">AND / OR combination of two conditions.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTop10ValueFilter(n, colId?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBottom10ValueFilter(n, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Top / bottom n rows by value.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTop10PercentFilter(p, colId?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBottom10PercentFilter(p, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Top / bottom p percent by value.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAverageFilter(aboveAverage, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Above average (true) / below average (false).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesNumberFilter(vals, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Show only rows matching one of the numbers in the list.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesStringFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesDateTimeFilter</code></td><td class="border border-surface-200 px-3 py-2 align-top">String / date-time versions (see <a href="#caveats" class="text-primary-600 hover:underline">Caveats</a>).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeGroupingFilter</code></td><td class="border border-surface-200 px-3 py-2 align-top">Date period / grouping conditions.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFontColorFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCellColorFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setIconFilter</code></td><td class="border border-surface-200 px-3 py-2 align-top">Conditions based on color / icon.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSort()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sort</code> that protects the header row.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetFilter(colId)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllFilter()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Clear conditions (the filter itself remains).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeFilter()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Removes the AutoFilter itself and restores hidden rows.</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">op</code> is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlAutoFilterOperator</code>:</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Value</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Number</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Meaning</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Equal</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">equal to</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">LessThan</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">less than</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">LessThanOrEqual</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">less than or equal to</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">NotEqual</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">not equal to</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">GreaterThanOrEqual</code></td><td class="border border-surface-200 px-3 py-2 align-top">5</td><td class="border border-surface-200 px-3 py-2 align-top">greater than or equal to</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">GreaterThan</code></td><td class="border border-surface-200 px-3 py-2 align-top">6</td><td class="border border-surface-200 px-3 py-2 align-top">greater than</td></tr>
  </tbody></table></div>
<h3 id="nodeosbxl-sort-obtained-via-ws-getsort-a1c1-firstrowasheader-or-af-getsort" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl.Sort</code> — obtained via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort(A1C1, firstRowAsHeader?)</code> or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending(target, direction?, matchCase?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sorts ascending. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">target</code> is the column number within the range (first column = 1).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortDescending(target, direction?, matchCase?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sorts descending.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">execute(sortFieldObject)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sorts by a single <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.SortFieldObject</code> (configure it with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setSortOnValues(target, ascending)</code>).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeMultiple([...])</code></td><td class="border border-surface-200 px-3 py-2 align-top">Multi-key sort (works correctly since the 2026-10-01 fix. See <a href="#caveats" class="text-primary-600 hover:underline">Caveats</a>).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortConditions()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Reads back the configured conditions as an array of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.SortFieldObject</code>.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetSort(target)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllSort()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Clear conditions. <strong>Only effective on a Sort obtained via AutoFilter / Table.</strong></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlRowCol</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Rows</code> = 1 / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Columns</code> = 2). The default is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Rows</code> (sort rows).</p>
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
        <pre><code class="language-javascript">// af-sort.js — AutoFilter and sorting
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
["Product", "Region", "Qty", "Unit Price"].forEach((h, i) =&gt;
    ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) =&gt; {
    const r = i + 2;
    ws.getRange(`A${r}:A${r}`).setValue(row[0]);
    ws.getRange(`B${r}:B${r}`).setValue(row[1]);
    ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
    ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] data ready: A1:D11 (header + 10 rows)");

// display helpers: row number → "Product(value)" / list of rows that are not hidden
const text = (r, col) =&gt; ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) =&gt; `${text(r, "A")}(${text(r, col)})`;
const visible = (col) =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
    return out.join(" ");
};
const hiddenList = () =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
    return out.length ? out.join(",") : "none";
};

// ---------- 2. set up the AutoFilter ----------
const af = ws.getAutoFilter("A1:D11", true);   // first row is the header
console.log(`[2] getAutoFilter → range = ${af.getAddress()} / hidden rows = ${hiddenList()}`);

// ---------- 3. custom filter: Qty &gt; 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] visible rows with Qty&gt;100: ${visible("C")}`);
console.log(`    hidden rows = ${hiddenList()} / getValue("A2:A11") count = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter();                            // clear conditions only (the filter itself remains)

// ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3);                   // top 3 by Qty
console.log(`[4] visible rows in top 3 by Qty: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3);   // Qty is 120 or 40
console.log(`[5] visible rows with Qty in {120,40}: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2);   // Region is Tokyo or Fukuoka (string list)
console.log(`[6] visible rows with Region in {Tokyo,Fukuoka}: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "Osaka", 2);   // Region = Osaka (string match)
console.log(`[7] visible rows with Region="Osaka": ${visible("B")}`);
af.resetAllFilter();

// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
ws.getSort("A2:D11").executeSortAscending(3);   // ascending by Qty (3rd column within the range)
console.log(`[8] ascending by Qty: ${visible("C")}`);

// a Sort obtained via the AutoFilter protects the header row automatically
const sort = af.getSort();
sort.executeSortDescending(4);                  // descending by Unit Price (4th column)
console.log(`[9] descending by Unit Price: ${visible("D")}`);

// ---------- 6. read the sort conditions back ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] condition read back: target column = ${cond.getTarget()} / ascending = ${cond.getSortAscending()}`);

// ---------- 7. remove the AutoFilter itself ----------
af.removeFilter();
console.log(`[11] after removeFilter: hidden rows = ${hiddenList()}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># af-sort.py — same verification as the Node.js sample (pyosbxl)
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

def hidden_list():
    out = [str(r) for r in range(2, 12) if ws.getRow(r).getHidden()]
    return ",".join(out) if out else "none"

# ---------- 2. set up the AutoFilter ----------
af = ws.getAutoFilter("A1:D11", True)   # first row is the header
print(f"[2] getAutoFilter → range = {af.getAddress()} / hidden rows = {hidden_list()}")

# ---------- 3. custom filter: Qty &gt; 100 ----------
# colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.GreaterThan, "100", 3)
print(f"[3] visible rows with Qty&gt;100: {visible('C')}")
print(f'    hidden rows = {hidden_list()} / getValue("A2:A11") count = {len(ws.getRange("A2:A11").getValue())}')
af.resetAllFilter()                     # clear conditions only (the filter itself remains)

# ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3)            # top 3 by Qty
print(f"[4] visible rows in top 3 by Qty: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesNumberFilter([120, 40], 3)   # Qty is 120 or 40
print(f"[5] visible rows with Qty in {{120,40}}: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2)   # Region is Tokyo or Fukuoka (string list)
print(f"[6] visible rows with Region in {{Tokyo,Fukuoka}}: {visible('B')}")
af.resetAllFilter()
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.Equal, "Osaka", 2)   # Region = Osaka (string match)
print(f'[7] visible rows with Region="Osaka": {visible("B")}')
af.resetAllFilter()

# ---------- 5. sorting (single key) ----------
# target only the data rows without the header (firstRowAsHeader omitted = False)
ws.getSort("A2:D11").executeSortAscending(3)   # ascending by Qty (3rd column within the range)
print(f"[8] ascending by Qty: {visible('C')}")

# a Sort obtained via the AutoFilter protects the header row automatically
sort = af.getSort()
sort.executeSortDescending(4)                  # descending by Unit Price (4th column)
print(f"[9] descending by Unit Price: {visible('D')}")

# ---------- 6. read the sort conditions back ----------
cond = sort.getSortConditions()[0]
# Python prints booleans as True/False, so use lower() to match the JS style
print(f"[10] condition read back: target column = {cond.getTarget()} / ascending = {str(cond.getSortAscending()).lower()}")

# ---------- 7. remove the AutoFilter itself ----------
af.removeFilter()
print(f"[11] after removeFilter: hidden rows = {hidden_list()}")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// af-sort.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
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
        List&lt;String&gt; out = new ArrayList&lt;&gt;();
        for (int r = 2; r &lt;= 11; r++) {
            if (!ws.getRow(r).getHidden()) out.add(label(r, col));
        }
        return String.join(" ", out);
    }

    static String hiddenList() {
        List&lt;String&gt; out = new ArrayList&lt;&gt;();
        for (int r = 2; r &lt;= 11; r++) {
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
        for (int i = 0; i &lt; headers.length; i++) {
            String col = String.valueOf("ABCD".charAt(i));
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        for (int i = 0; i &lt; DATA.length; i++) {
            int r = i + 2;
            ws.getRange("A" + r + ":A" + r).setValue((String) DATA[i][0]);
            ws.getRange("B" + r + ":B" + r).setValue((String) DATA[i][1]);
            ws.getRange("C" + r + ":C" + r).setNumberValue((Double) DATA[i][2]);
            ws.getRange("D" + r + ":D" + r).setNumberValue((Double) DATA[i][3]);
        }
        System.out.println("[1] data ready: A1:D11 (header + 10 rows)");

        // ---------- 2. set up the AutoFilter ----------
        AutoFilterWrapper af = ws.getAutoFilter("A1:D11", true);   // first row is the header
        System.out.println("[2] getAutoFilter → range = " + af.getAddress() + " / hidden rows = " + hiddenList());

        // ---------- 3. custom filter: Qty &gt; 100 ----------
        // colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
        af.setCustomFilter(XlAutoFilterOperator.GreaterThan, "100", 3);
        System.out.println("[3] visible rows with Qty&gt;100: " + visible("C"));
        System.out.println("    hidden rows = " + hiddenList()
                + " / getValue(\"A2:A11\") count = " + ws.getRange("A2:A11").getValue().size());
        af.resetAllFilter();                            // clear conditions only (the filter itself remains)

        // ---------- 4. Top-N / value-list filters ----------
        af.setTop10ValueFilter(3, 3);                   // top 3 by Qty
        System.out.println("[4] visible rows in top 3 by Qty: " + visible("C"));
        af.resetAllFilter();
        af.setUniqueValuesNumberFilter(Arrays.asList(120.0, 40.0), 3);   // Qty is 120 or 40
        System.out.println("[5] visible rows with Qty in {120,40}: " + visible("C"));
        af.resetAllFilter();
        af.setUniqueValuesStringFilter(Arrays.asList("Tokyo", "Fukuoka"), 2);   // Region is Tokyo or Fukuoka (string list)
        System.out.println("[6] visible rows with Region in {Tokyo,Fukuoka}: " + visible("B"));
        af.resetAllFilter();
        af.setCustomFilter(XlAutoFilterOperator.Equal, "Osaka", 2);   // Region = Osaka (string match)
        System.out.println("[7] visible rows with Region=\"Osaka\": " + visible("B"));
        af.resetAllFilter();

        // ---------- 5. sorting (single key) ----------
        // target only the data rows without the header (firstRowAsHeader omitted = false)
        ws.getSort("A2:D11").executeSortAscending(3);   // ascending by Qty (3rd column within the range)
        System.out.println("[8] ascending by Qty: " + visible("C"));

        // a Sort obtained via the AutoFilter protects the header row automatically
        SortWrapper sort = af.getSort();
        sort.executeSortDescending(4);                  // descending by Unit Price (4th column)
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
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// af-sort.php — AutoFilter and sorting (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("af-sort-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

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
foreach (["Product", "Region", "Qty", "Unit Price"] as $i =&gt; $h) {
    $col = "ABCD"[$i];
    $ws-&gt;getRange("{$col}1:{$col}1")-&gt;setValue($h);
}
foreach ($DATA as $i =&gt; $row) {
    $r = $i + 2;
    $ws-&gt;getRange("A{$r}:A{$r}")-&gt;setValue($row[0]);
    $ws-&gt;getRange("B{$r}:B{$r}")-&gt;setValue($row[1]);
    $ws-&gt;getRange("C{$r}:C{$r}")-&gt;setNumberValue($row[2]);
    $ws-&gt;getRange("D{$r}:D{$r}")-&gt;setNumberValue($row[3]);
}
echo "[1] data ready: A1:D11 (header + 10 rows)\n";

// display helpers: row number → "Product(value)" / list of rows that are not hidden
$text = fn(int $r, string $col): string =&gt; $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;getValue()["{$col}{$r}"];
$label = fn(int $r, string $col): string =&gt; $text($r, "A") . "(" . $text($r, $col) . ")";
$visible = function (string $col) use ($ws, $label): string {
    $out = [];
    for ($r = 2; $r &lt;= 11; $r++) {
        if (!$ws-&gt;getRow($r)-&gt;getHidden()) {
            $out[] = $label($r, $col);
        }
    }
    return implode(" ", $out);
};
$hiddenList = function () use ($ws): string {
    $out = [];
    for ($r = 2; $r &lt;= 11; $r++) {
        if ($ws-&gt;getRow($r)-&gt;getHidden()) {
            $out[] = (string)$r;
        }
    }
    return count($out) &gt; 0 ? implode(",", $out) : "none";
};

// ---------- 2. set up the AutoFilter ----------
$af = $ws-&gt;getAutoFilter("A1:D11", true);   // first row is the header
echo "[2] getAutoFilter → range = ", $af-&gt;getAddress(), " / hidden rows = ", $hiddenList(), "\n";

// ---------- 3. custom filter: Qty &gt; 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
$af-&gt;setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::GreaterThan, "100", 3);
echo "[3] visible rows with Qty&gt;100: ", $visible("C"), "\n";
echo "    hidden rows = ", $hiddenList(),
    " / getValue(\"A2:A11\") count = ", count($ws-&gt;getRange("A2:A11")-&gt;getValue()), "\n";
$af-&gt;resetAllFilter();                      // clear conditions only (the filter itself remains)

// ---------- 4. Top-N / value-list filters ----------
$af-&gt;setTop10ValueFilter(3, 3);             // top 3 by Qty
echo "[4] visible rows in top 3 by Qty: ", $visible("C"), "\n";
$af-&gt;resetAllFilter();
$af-&gt;setUniqueValuesNumberFilter([120, 40], 3);   // Qty is 120 or 40
echo "[5] visible rows with Qty in {120,40}: ", $visible("C"), "\n";
$af-&gt;resetAllFilter();
$af-&gt;setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2);   // Region is Tokyo or Fukuoka (string list)
echo "[6] visible rows with Region in {Tokyo,Fukuoka}: ", $visible("B"), "\n";
$af-&gt;resetAllFilter();
$af-&gt;setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::Equal, "Osaka", 2);   // Region = Osaka (string match)
echo "[7] visible rows with Region=\"Osaka\": ", $visible("B"), "\n";
$af-&gt;resetAllFilter();

// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
$ws-&gt;getSort("A2:D11")-&gt;executeSortAscending(3);   // ascending by Qty (3rd column within the range)
echo "[8] ascending by Qty: ", $visible("C"), "\n";

// a Sort obtained via the AutoFilter protects the header row automatically
$sort = $af-&gt;getSort();
$sort-&gt;executeSortDescending(4);                  // descending by Unit Price (4th column)
echo "[9] descending by Unit Price: ", $visible("D"), "\n";

// ---------- 6. read the sort conditions back ----------
$cond = $sort-&gt;getSortConditions()[0];
// echo of a PHP bool prints 1/empty string, so render JS-style true/false
echo "[10] condition read back: target column = ", $cond-&gt;getTarget(),
    " / ascending = ", $cond-&gt;getSortAscending() ? "true" : "false", "\n";

// ---------- 7. remove the AutoFilter itself ----------
$af-&gt;removeFilter();
echo "[11] after removeFilter: hidden rows = ", $hiddenList(), "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// af-sort.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
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
["Product", "Region", "Qty", "Unit Price"].forEach((h, i) =&gt;
    ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) =&gt; {
    const r = i + 2;
    ws.getRange(`A${r}:A${r}`).setValue(row[0]);
    ws.getRange(`B${r}:B${r}`).setValue(row[1]);
    ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
    ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] data ready: A1:D11 (header + 10 rows)");

// display helpers: row number → "Product(value)" / list of rows that are not hidden
const text = (r, col) =&gt; ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) =&gt; `${text(r, "A")}(${text(r, col)})`;
const visible = (col) =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
    return out.join(" ");
};
const hiddenList = () =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
    return out.length ? out.join(",") : "none";
};

// ---------- 2. set up the AutoFilter ----------
const af = ws.getAutoFilter("A1:D11", true);   // first row is the header
console.log(`[2] getAutoFilter → range = ${af.getAddress()} / hidden rows = ${hiddenList()}`);

// ---------- 3. custom filter: Qty &gt; 100 ----------
// colId is the column number within the range (A=1 … D=4). Do not put operator symbols in criteria
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] visible rows with Qty&gt;100: ${visible("C")}`);
console.log(`    hidden rows = ${hiddenList()} / getValue("A2:A11") count = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter();                            // clear conditions only (the filter itself remains)

// ---------- 4. Top-N / value-list filters ----------
af.setTop10ValueFilter(3, 3);                   // top 3 by Qty
console.log(`[4] visible rows in top 3 by Qty: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3);   // Qty is 120 or 40
console.log(`[5] visible rows with Qty in {120,40}: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["Tokyo", "Fukuoka"], 2);   // Region is Tokyo or Fukuoka (string list)
console.log(`[6] visible rows with Region in {Tokyo,Fukuoka}: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "Osaka", 2);   // Region = Osaka (string match)
console.log(`[7] visible rows with Region="Osaka": ${visible("B")}`);
af.resetAllFilter();

// ---------- 5. sorting (single key) ----------
// target only the data rows without the header (firstRowAsHeader omitted = false)
ws.getSort("A2:D11").executeSortAscending(3);   // ascending by Qty (3rd column within the range)
console.log(`[8] ascending by Qty: ${visible("C")}`);

// a Sort obtained via the AutoFilter protects the header row automatically
const sort = af.getSort();
sort.executeSortDescending(4);                  // descending by Unit Price (4th column)
console.log(`[9] descending by Unit Price: ${visible("D")}`);

// ---------- 6. read the sort conditions back ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] condition read back: target column = ${cond.getTarget()} / ascending = ${cond.getSortAscending()}`);

// ---------- 7. remove the AutoFilter itself ----------
af.removeFilter();
console.log(`[11] after removeFilter: hidden rows = ${hiddenList()}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] data ready: A1:D11 (header + 10 rows)
[2] getAutoFilter → range = A1:D11 / hidden rows = none
[3] visible rows with Qty&gt;100: Apple(120) Banana(200) Peach(150) Strawberry(110)
    hidden rows = 3,5,6,8,9,11 / getValue("A2:A11") count = 10
[4] visible rows in top 3 by Qty: Apple(120) Banana(200) Peach(150)
[5] visible rows with Qty in {120,40}: Apple(120) Orange(40)
[6] visible rows with Region in {Tokyo,Fukuoka}: Apple(Tokyo) Banana(Tokyo) Grape(Fukuoka) Peach(Fukuoka) Kiwi(Tokyo) Strawberry(Fukuoka) Melon(Tokyo)
[7] visible rows with Region="Osaka": Orange(Osaka) Pear(Osaka) Mango(Osaka)
[8] ascending by Qty: Melon(25) Kiwi(30) Orange(40) Grape(60) Mango(70) Pear(90) Strawberry(110) Apple(120) Peach(150) Banana(200)
[9] descending by Unit Price: Melon(800) Mango(500) Peach(400) Strawberry(350) Grape(300) Banana(150) Pear(120) Apple(100) Kiwi(90) Orange(80)
[10] condition read back: target column = 4 / ascending = false
[11] after removeFilter: hidden rows = none</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<h3 id="filters-hiding-rows" class="text-xl font-bold mt-8 mb-3">Filters = hiding rows</h3>
<p class="mb-4 text-slate-700 leading-7">An AutoFilter only <strong>hides the rows that do not match the condition</strong>; no data is deleted.</p>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Check the hidden state with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRow(r).getHidden()</code> (see <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">Rows, columns and ranges</a> for <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row</code> usage).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> returns the values of hidden rows as well</strong> (output <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[3]</code>: the count is 10 even though 6 rows are hidden). If you want "only the visible rows", narrow them down with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHidden()</code> before reading.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetFilter(colId)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllFilter()</code> clear <strong>only the conditions</strong>; the filter (frame) remains.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeFilter()</code> removes the filter itself and makes all hidden rows visible again (output <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[11]</code>).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">The filter range, conditions, and hidden state are <strong>restored after save → reopen</strong> (measured).</li>
</ul>
<h3 id="writing-filter-conditions" class="text-xl font-bold mt-8 mb-3">Writing filter conditions</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Both <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colId</code> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(op, criteria, colId)</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">target</code> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sort</code> are <strong>the column number within the specified range (first column = 1)</strong> — not the sheet column number itself, so watch out for the offset when the range starts at column <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">B</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Write <strong>only the value</strong> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">criteria</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"100"</code>). The operator is specified via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">op</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlAutoFilterOperator</code>).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Numeric comparisons (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">GreaterThan</code> etc.), Top-N, average, and number-list matching (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesNumberFilter</code>) all work as expected (output <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[3]</code>–<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[5]</code>).</li>
</ul>
<h3 id="sorting" class="text-xl font-bold mt-8 mb-3">Sorting</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">The basic form is to <strong>pass only the data rows as the range</strong>, as in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort("A2:D11")</code>. If you pass a range that includes the header, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort("A1:D11", true)</code> (2nd argument <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">firstRowAsHeader</code>) or <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code> (the header is protected automatically)</strong>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Sorting moves <strong>whole rows</strong>. Columns other than the key column (Product, Region, etc.) move along with their row.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">After running, read the conditions back with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortConditions()</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTarget()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortAscending()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDirection()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMatchCase()</code>). The read-back conditions also survive save → reopen.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">The <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending(target, direction, matchCase)</code> is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlRowCol.Rows</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Columns</code>(2). The default is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Rows</code>.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Pass only the value in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">criteria</code> (no operator symbols, no wildcards).</strong> The operator is specified via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlAutoFilterOperator</code> by design.<br>Mixing symbols in, as in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(GreaterThan, "&gt;100", 3)</code>, makes the condition be treated as a string; on a numeric column it matches nothing and all rows are hidden (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"100"</code> works fine).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>String match filters work correctly</strong> (fixed on 2026-10-01. In the old version 1.5.0 there was a bug where string cells newly created within the session did not match and all rows were hidden). <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(Equal, "Tokyo", ...)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesStringFilter(["Tokyo","Fukuoka"], ...)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setOrCustomFilter(...)</code> all work as shown in output <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[6]</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[7]</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeMultiple()</code> (multi-key sort) works correctly</strong> (fixed on 2026-10-01). In the old version 1.5.0 rows were duplicated or lost.<br>Build each condition with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SortFieldObject.setSortOnValues(column, ascending?)</code> and pass them as an array in priority order.<br>Acceptance test: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">test/sort_af_fc_ext.test.js</code> "Sort executeMultiple (multiple keys, A-2 fix)".</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Calling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending()</code> multiple times does not produce a multi-key sort.</strong> The whole table is re-sorted by the key called last, while conditions keep accumulating in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortConditions()</code> (measured).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Sorting a header-inclusive range without <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">firstRowAsHeader</code> (default false) sorts the header row as data.</strong> In a measured run, an ascending sort sank the row-1 header (a string) to the very bottom.<br>For ranges that include the header, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSort(range, true)</code> or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Sorting while a filter is applied reorders only the visible rows.</strong> Hidden rows stay at their original positions, so it is safer to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllFilter()</code> before sorting (measured).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetSort()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllSort()</code> work only on a Sort obtained via AutoFilter / Table.</strong> We confirmed by measurement that conditions are not cleared on a Sort obtained via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort()</code> (the specification as documented).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAutoFilter(A1C1, firstRowAsHeader?)</code> is a "getter" that also sets the range.</strong> There is only one AutoFilter range per sheet; calling it with a different range replaces the configured range.<br>The 2nd argument is not "whether to get" but <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">firstRowAsHeader</code> (whether the first row is treated as the header).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Put numbers in numeric columns with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue()</code> (or values interpreted as numbers).</strong> We measured that comparison filters do work on numbers entered as strings such as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue("120")</code>, but numeric-typed input is recommended for stable behavior of Top-N and similar filters.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">Bulk Input of Large Data</a> / Next: <a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">Tables and Comments</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/en/nodejs/classes/dto.SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/nodejs/enums/enums.XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/en/nodejs/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/en/python/pyosbxl.Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/en/python/pyosbxl.dto.SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/en/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/AutoFilterWrapper.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/SortWrapper.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/SortFieldObjectWrapper.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/en/php/classes/phposbxl-Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/en/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/en/wasm/classes/osbxl.Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/en/wasm/classes/dto.SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/en/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/wasm/enums/enums.XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/en/wasm/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "AutoFilter and Sorting — filter, sort, read back",
   "description": "AutoFilter and Sorting — filter, sort, read back. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-07-autofilter-sort.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-07-autofilter-sort.php",
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
     "name": "AutoFilter and Sorting — filter, sort, read back",
     "item": "https://osb-office.com/en/sample/general-07-autofilter-sort.php"
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
