<?php
$lang = 'en';
$page_title = 'Tables and Comments — create, totals, structured refs | osbxl Samples';
$no_tagline = true;
$meta_description = 'Tables and Comments — create, totals, structured refs. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,listobjects,table,addlist,totalsrow,structured-reference,comments,setcomment,setmemo';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Tables and Comments — create, totals, structured refs</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-08-tables-comments.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Turn a data range into an Excel "table" (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ListObject</code>) and work with its name, column names, style and totals row. Use "structured references" to refer to table columns from formulas, and add, read and remove cell comments (modern threaded comments) and memos (legacy notes).</p>
<h2 id="classes-used" class="text-2xl font-bold mt-10 mb-4">Classes used</h2>
<p class="mb-4 text-slate-700 leading-7">Get tables from <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getListObjects()</code> and comments from <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getComments()</code>.</p>
<h3 id="listobjects-collection-of-tables" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ListObjects</code> (collection of tables)</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList(name, A1C1, useFirstRowAsHeader, insertTotals)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Creates a table from an existing cell range and returns the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Table</code>.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addListFromRange(name, topA1, copyFromA1C1, useFirstRowAsHeader, insertTotals)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Creates a table by <strong>copying</strong> the data to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">topA1</code>. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyFromA1C1</code> may point to another sheet, like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet2!A1:C1</code>.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getList(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a table. Throws a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">table is not found</code> exception if it does not exist.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList(name, deleteData?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Removes a table. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteData</code> defaults to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> (<strong>the data is deleted too</strong>). Pass <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code> to keep it.</td></tr>
  </tbody></table></div>
<h3 id="table-a-single-table" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Table</code> (a single table)</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Group</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Methods</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Ranges</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAllRange()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHeaderRowRange()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDataBodyRange()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowRange()</code> (all return a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code>; <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowRange()</code> throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">totalRow is not set</code> when there is no totals row)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Name</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">changeTableName(name)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Column names</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColumnName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnName([...])</code> (also changes what the header cells display)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Style</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTableStyleName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBuiltinStyleName(enums.XlDefaultTableStyle.Xxx, clearFormat)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomTableStyleName(name, clearFormat)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Totals row</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowTotals()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setShowTotals(bool)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowFunction(col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowFunction(col, enums.XlTotalsCalculation.Xxx)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomRowFunction(col, formula, isArray?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowLabel(col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowLabel(col, text)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Filter / display</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAutoFilter()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeAutoFilter()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowAutoFilter()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowHeaders()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setShowHeaders(bool)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowTableStyleRowStripes()</code> and more</td></tr>
  </tbody></table></div>
<h3 id="comments-cell-comments" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Comments</code> (cell comments)</h3>
<p class="mb-4 text-slate-700 leading-7">Excel has two comment systems: <strong>modern</strong> comments (threads, replies, done flag) and <strong>legacy memos</strong> (callout notes).</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">System</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment(A1, commentObject)</code></td><td class="border border-surface-200 px-3 py-2 align-top">modern</td><td class="border border-surface-200 px-3 py-2 align-top">Adds one comment. On a cell that already has comments this <strong>appends</strong> (automatically as a reply).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread(A1, [commentObject...])</code></td><td class="border border-surface-200 px-3 py-2 align-top">modern</td><td class="border border-surface-200 px-3 py-2 align-top">Sets a whole thread (<strong>overwrites</strong>). The first entry is the body, the rest are replies.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComment(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">modern</td><td class="border border-surface-200 px-3 py-2 align-top">Returns <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Array&lt;dto.CommentObject&gt;</code>. Throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">comment not found</code> if there is none.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeComment(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">modern</td><td class="border border-surface-200 px-3 py-2 align-top">Removes modern comments.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMemo(A1, author, text, fontObject, visible?, rowColumnsNum?, colColumnsNum?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">legacy</td><td class="border border-surface-200 px-3 py-2 align-top">Sets a callout memo. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">fontObject</code> can be <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new dto.FontObject()</code>.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoText(A1)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoAuthor(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">legacy</td><td class="border border-surface-200 px-3 py-2 align-top">Gets the memo text / author. Throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">memo not found</code> if there is none.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeMemo(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">legacy</td><td class="border border-surface-200 px-3 py-2 align-top">Removes a memo.</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CommentObject</code> has <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAuthor</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setContent</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setParentId</code> (parent ID of a reply) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDone</code> (done flag) plus the corresponding getters. Once set, an ID is automatically assigned as a UUID in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{...}</code> form.</p>
<h3 id="totals-functions-enums-xltotalscalculation" class="text-xl font-bold mt-8 mb-3">Totals functions <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlTotalsCalculation</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Member</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Value</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Meaning</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationNone</code></td><td class="border border-surface-200 px-3 py-2 align-top">0</td><td class="border border-surface-200 px-3 py-2 align-top">No totals</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationSum</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">Sum</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationAverage</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">Average</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCount</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">Count</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCountNums</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">Count of numbers</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationMin</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationMax</code></td><td class="border border-surface-200 px-3 py-2 align-top">5 / 6</td><td class="border border-surface-200 px-3 py-2 align-top">Min / Max</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationStdDev</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationVar</code></td><td class="border border-surface-200 px-3 py-2 align-top">7 / 8</td><td class="border border-surface-200 px-3 py-2 align-top">StdDev / Var</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCustom</code></td><td class="border border-surface-200 px-3 py-2 align-top">9</td><td class="border border-surface-200 px-3 py-2 align-top">Custom formula (set with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomRowFunction</code>)</td></tr>
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
        <pre><code class="language-javascript">// tables-comments.js — table (ListObject) creation, totals row, structured references, cell comments
const { nodeosbxl, enums, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("tables-comments.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- fill in the data (1 header row + 3 data rows) ----------
const rows = [
    ["Product", "Unit Price", "Qty"],
    ["Apple", 120, 3],
    ["Orange", 80, 5],
    ["Grape", 300, 2],
];
const values = [];
rows.forEach((row, ri) =&gt; row.forEach((cell, ci) =&gt; {
    const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
    const o = new dto.InputValueObject();
    if (typeof cell === "number") o.setNumberValue(a1, cell);
    else o.setStringValue(a1, cell);
    values.push(o);
}));
ws.setValueArray(values);

// ---------- 1. create the table ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] Table created");
console.log(`    name       = ${table.getName()}`);
console.log(`    columns    = ${JSON.stringify(table.getColumnName())}`);
console.log(`    style      = ${table.getTableStyleName()}`);
console.log(`    full range = ${table.getAllRange().getAddress()}`);
console.log(`    header     = ${table.getHeaderRowRange().getAddress()}`);
console.log(`    data body  = ${table.getDataBodyRange().getAddress()}`);
console.log(`    totals row = ${table.getTotalRowRange().getAddress()}`);

// ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] Totals row");
console.log(`    label (col 1) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(`    formula       = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(`    value         = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);

// ---------- 3. change the table style ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] style after change = ${table.getTableStyleName()}`);

// ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])"); // no leading "="
console.log("[4] Structured reference");
console.log(`    E1 formula = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(`    E1 value   = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);

// ---------- 5. comments (modern) and memos (legacy) ----------
const cm = ws.getComments();
const mkComment = (author, content) =&gt; {
    const c = new dto.CommentObject();
    c.setAuthor(author);
    c.setContent(content);
    return c;
};

// modern comment (thread): body + reply = 2 entries -&gt; B2 (kept in the file after save)
cm.setCommentThread("B2", [mkComment("alice", "Please check this number"), mkComment("bob", "Checked")]);
// single modern comment -&gt; D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -&gt; A2
cm.setMemo("A2", "Carol", "Top-priority product", new dto.FontObject());

const thread = cm.getComment("B2");
console.log("[5] Comments");
console.log(`    B2 thread count = ${thread.length}`);
thread.forEach((c, i) =&gt; console.log(`      [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (reply)" : ""}`));
console.log(`    D2 comment count = ${cm.getComment("D2").length}`);
console.log(`    A2 memo = ${JSON.stringify(cm.getMemoText("A2"))} / author ${JSON.stringify(cm.getMemoAuthor("A2"))}`);

// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) =&gt; {
    try { fn(); console.log(`    after removal ${label}: no exception`); }
    catch (e) { console.log(`    after removal ${label}: exception ${e.message}`); }
};
tryRead("D2 comment", () =&gt; cm.getComment("D2"));
tryRead("A2 memo", () =&gt; cm.getMemoText("A2"));
console.log(`    B2 is kept, so ${cm.getComment("B2").length} comments remain after save`);

const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[Check] table=${table.getName()} total(Unit Price)=${total} structuredRef(E1)=${e1}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># tables-comments.py — same verification as the Node.js sample (pyosbxl)
import json
import pyosbxl


def jdump(obj):  # JSON.stringify-equivalent output (no separator spaces, non-ASCII kept as-is)
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


app = pyosbxl.App()
wb = app.createWorkBook("tables-comments-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- fill in the data (1 header row + 3 data rows) ----------
rows = [
    ["Product", "Unit Price", "Qty"],
    ["Apple", 120, 3],
    ["Orange", 80, 5],
    ["Grape", 300, 2],
]
values = []
for ri, row in enumerate(rows):
    for ci, cell in enumerate(row):
        a1 = app.convertFromRowColNumber(ri + 1, ci + 1)
        o = pyosbxl.dto.InputValueObject()
        if isinstance(cell, (int, float)):
            o.setNumberValue(a1, cell)
        else:
            o.setStringValue(a1, cell)
        values.append(o)
ws.setValueArray(values)

# ---------- 1. create the table ----------
lo = ws.getListObjects()
table = lo.addList("Sales", "A1:C4", True, True)  # useFirstRowAsHeader, insertTotals
print("[1] Table created")
print(f"    name       = {table.getName()}")
print(f"    columns    = {jdump(table.getColumnName())}")
print(f"    style      = {table.getTableStyleName()}")
print(f"    full range = {table.getAllRange().getAddress()}")
print(f"    header     = {table.getHeaderRowRange().getAddress()}")
print(f"    data body  = {table.getDataBodyRange().getAddress()}")
print(f"    totals row = {table.getTotalRowRange().getAddress()}")

# ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total")
table.setTotalRowFunction(2, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
table.setTotalRowFunction(3, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
print("[2] Totals row")
print(f"    label (col 1) = {jdump(table.getTotalRowLabel(1))}")
print(f"    formula       = {jdump(ws.getRange('B5:C5').getFormula())}")
print(f"    value         = {jdump(ws.getRange('B5:C5').getValue(True))}")

# ---------- 3. change the table style ----------
table.setBuiltinStyleName(pyosbxl.enums.XlDefaultTableStyle.TableStyleMedium9, True)
print(f"[3] style after change = {table.getTableStyleName()}")

# ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])")  # no leading "="
print("[4] Structured reference")
print(f"    E1 formula = {jdump(ws.getRange('E1:E1').getFormula())}")
print(f"    E1 value   = {jdump(ws.getRange('E1:E1').getValue(True))}")

# ---------- 5. comments (modern) and memos (legacy) ----------
cm = ws.getComments()


def mkComment(author, content):
    c = pyosbxl.dto.CommentObject()
    c.setAuthor(author)
    c.setContent(content)
    return c


# modern comment (thread): body + reply = 2 entries -&gt; B2 (kept in the file after save)
cm.setCommentThread("B2", [mkComment("alice", "Please check this number"), mkComment("bob", "Checked")])
# single modern comment -&gt; D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"))
# legacy memo (callout) -&gt; A2
cm.setMemo("A2", "Carol", "Top-priority product", pyosbxl.dto.FontObject())

thread = cm.getComment("B2")
print("[5] Comments")
print(f"    B2 thread count = {len(thread)}")
for i, c in enumerate(thread):
    print(f"      [{i}] {c.getAuthor()}: {c.getContent()}{' (reply)' if c.getParentId() else ''}")
print(f"    D2 comment count = {len(cm.getComment('D2'))}")
print(f"    A2 memo = {jdump(cm.getMemoText('A2'))} / author {jdump(cm.getMemoAuthor('A2'))}")

# removal: removeComment for modern, removeMemo for legacy
cm.removeComment("D2")
cm.removeMemo("A2")


# reading after removal throws "comment not found" / "memo not found" just like the Node.js version
def tryRead(label, fn):
    try:
        fn()
        print(f"    after removal {label}: no exception")
    except Exception as e:
        print(f"    after removal {label}: exception {e}")


tryRead("D2 comment", lambda: cm.getComment("D2"))
tryRead("A2 memo", lambda: cm.getMemoText("A2"))
print(f"    B2 is kept, so {len(cm.getComment('B2'))} comments remain after save")

total = ws.getRange("B5:B5").getValue(True)["B5"]
e1 = ws.getRange("E1:E1").getValue(True)["E1"]
print(f"[Check] table={table.getName()} total(Unit Price)={total} structuredRef(E1)={e1}")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// tables-comments.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class Main {
    // JSON.stringify-equivalent output helpers (Maps keep insertion order)
    static String jsonMap(Map&lt;String, String&gt; m) {
        StringBuilder sb = new StringBuilder("{");
        boolean first = true;
        for (Map.Entry&lt;String, String&gt; e : m.entrySet()) {
            if (!first) sb.append(",");
            first = false;
            sb.append("\"").append(e.getKey()).append("\":\"").append(e.getValue()).append("\"");
        }
        return sb.append("}").toString();
    }

    static String jsonList(List&lt;String&gt; l) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i &lt; l.size(); i++) {
            if (i &gt; 0) sb.append(",");
            sb.append("\"").append(l.get(i)).append("\"");
        }
        return sb.append("]").toString();
    }

    static CommentObjectWrapper mkComment(String author, String content) {
        CommentObjectWrapper c = new CommentObjectWrapper();
        c.setAuthor(author);
        c.setContent(content);
        return c;
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("tables-comments-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- fill in the data (1 header row + 3 data rows) ----------
        Object[][] rows = {
            {"Product", "Unit Price", "Qty"},
            {"Apple", 120, 3},
            {"Orange", 80, 5},
            {"Grape", 300, 2},
        };
        List&lt;InputValueObjectWrapper&gt; values = new ArrayList&lt;&gt;();
        for (int ri = 0; ri &lt; rows.length; ri++) {
            for (int ci = 0; ci &lt; rows[ri].length; ci++) {
                String a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
                InputValueObjectWrapper o = new InputValueObjectWrapper();
                if (rows[ri][ci] instanceof Number) {
                    o.setNumberValue(a1, ((Number) rows[ri][ci]).doubleValue());
                } else {
                    o.setStringValue(a1, (String) rows[ri][ci]);
                }
                values.add(o);
            }
        }
        ws.setValueArray(values);

        // ---------- 1. create the table ----------
        ListObjectsWrapper lo = ws.getListObjects();
        TableWrapper table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
        System.out.println("[1] Table created");
        System.out.println("    name       = " + table.getName());
        System.out.println("    columns    = " + jsonList(table.getColumnName()));
        System.out.println("    style      = " + table.getTableStyleName());
        System.out.println("    full range = " + table.getAllRange().getAddress());
        System.out.println("    header     = " + table.getHeaderRowRange().getAddress());
        System.out.println("    data body  = " + table.getDataBodyRange().getAddress());
        System.out.println("    totals row = " + table.getTotalRowRange().getAddress());

        // ---------- 2. totals row ----------
        table.setTotalRowLabel(1, "Total");
        table.setTotalRowFunction(2, XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(3, XlTotalsCalculation.TotalsCalculationSum);
        System.out.println("[2] Totals row");
        System.out.println("    label (col 1) = \"" + table.getTotalRowLabel(1) + "\"");
        System.out.println("    formula       = " + jsonMap(ws.getRange("B5:C5").getFormula()));
        System.out.println("    value         = " + jsonMap(ws.getRange("B5:C5").getValue(true)));

        // ---------- 3. change the table style ----------
        table.setBuiltinStyleName(XlDefaultTableStyle.TableStyleMedium9, true);
        System.out.println("[3] style after change = " + table.getTableStyleName());

        // ---------- 4. structured reference (formula referring to a table column) ----------
        ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])"); // no leading "="
        System.out.println("[4] Structured reference");
        System.out.println("    E1 formula = " + jsonMap(ws.getRange("E1:E1").getFormula()));
        System.out.println("    E1 value   = " + jsonMap(ws.getRange("E1:E1").getValue(true)));

        // ---------- 5. comments (modern) and memos (legacy) ----------
        CommentsWrapper cm = ws.getComments();
        // modern comment (thread): body + reply = 2 entries -&gt; B2 (kept in the file after save)
        cm.setCommentThread("B2", Arrays.asList(
            mkComment("alice", "Please check this number"), mkComment("bob", "Checked")));
        // single modern comment -&gt; D2 (used to demonstrate removal)
        cm.setComment("D2", mkComment("dave", "Comment to delete later"));
        // legacy memo (callout) -&gt; A2
        cm.setMemo("A2", "Carol", "Top-priority product", new FontObjectWrapper());

        List&lt;CommentObjectWrapper&gt; thread = cm.getComment("B2");
        System.out.println("[5] Comments");
        System.out.println("    B2 thread count = " + thread.size());
        for (int i = 0; i &lt; thread.size(); i++) {
            CommentObjectWrapper c = thread.get(i);
            String pid = c.getParentId();
            boolean isReply = pid != null &amp;&amp; !pid.isEmpty();
            System.out.println("      [" + i + "] " + c.getAuthor() + ": " + c.getContent()
                + (isReply ? " (reply)" : ""));
        }
        System.out.println("    D2 comment count = " + cm.getComment("D2").size());
        System.out.println("    A2 memo = \"" + cm.getMemoText("A2") + "\" / author \"" + cm.getMemoAuthor("A2") + "\"");

        // removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
        cm.removeComment("D2");
        cm.removeMemo("A2");
        try {
            cm.getComment("D2");
            System.out.println("    after removal D2 comment: no exception");
        } catch (RuntimeException e) {
            System.out.println("    after removal D2 comment: exception " + e.getMessage());
        }
        try {
            cm.getMemoText("A2");
            System.out.println("    after removal A2 memo: no exception");
        } catch (RuntimeException e) {
            System.out.println("    after removal A2 memo: exception " + e.getMessage());
        }
        System.out.println("    B2 is kept, so " + cm.getComment("B2").size() + " comments remain after save");

        String total = ws.getRange("B5:B5").getValue(true).get("B5");
        String e1 = ws.getRange("E1:E1").getValue(true).get("E1");
        System.out.println("[Check] table=" + table.getName() + " total(Unit Price)=" + total + " structuredRef(E1)=" + e1);

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// tables-comments.php — table (ListObject) creation, totals row, structured references, cell comments (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("tables-comments-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// JSON.stringify-compatible output (non-ASCII kept as-is)
$jq = fn($x): string =&gt; json_encode($x, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);

// ---------- fill in the data (1 header row + 3 data rows) ----------
$rows = [
    ["Product", "Unit Price", "Qty"],
    ["Apple", 120, 3],
    ["Orange", 80, 5],
    ["Grape", 300, 2],
];
$values = [];
foreach ($rows as $ri =&gt; $row) {
    foreach ($row as $ci =&gt; $cell) {
        $a1 = $app-&gt;convertFromRowColNumber($ri + 1, $ci + 1);
        $o = new \phposbxl\dto\InputValueObject();
        if (is_int($cell) || is_float($cell)) {
            $o-&gt;setNumberValue($a1, $cell);
        } else {
            $o-&gt;setStringValue($a1, $cell);
        }
        $values[] = $o;
    }
}
$ws-&gt;setValueArray($values);

// ---------- 1. create the table ----------
$lo = $ws-&gt;getListObjects();
$table = $lo-&gt;addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
echo "[1] Table created\n";
echo "    name       = ", $table-&gt;getName(), "\n";
echo "    columns    = ", $jq($table-&gt;getColumnName()), "\n";
echo "    style      = ", $table-&gt;getTableStyleName(), "\n";
echo "    full range = ", $table-&gt;getAllRange()-&gt;getAddress(), "\n";
echo "    header     = ", $table-&gt;getHeaderRowRange()-&gt;getAddress(), "\n";
echo "    data body  = ", $table-&gt;getDataBodyRange()-&gt;getAddress(), "\n";
echo "    totals row = ", $table-&gt;getTotalRowRange()-&gt;getAddress(), "\n";

// ---------- 2. totals row ----------
$table-&gt;setTotalRowLabel(1, "Total");
$table-&gt;setTotalRowFunction(2, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
$table-&gt;setTotalRowFunction(3, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
echo "[2] Totals row\n";
echo "    label (col 1) = ", $jq($table-&gt;getTotalRowLabel(1)), "\n";
echo "    formula       = ", $jq($ws-&gt;getRange("B5:C5")-&gt;getFormula()), "\n";
echo "    value         = ", $jq($ws-&gt;getRange("B5:C5")-&gt;getValue(true)), "\n";

// ---------- 3. change the table style ----------
$table-&gt;setBuiltinStyleName(\phposbxl\enums\XlDefaultTableStyle::TableStyleMedium9, true);
echo "[3] style after change = ", $table-&gt;getTableStyleName(), "\n";

// ---------- 4. structured reference (formula referring to a table column) ----------
$ws-&gt;getRange("E1:E1")-&gt;setFormula("SUM(Sales[Unit Price])"); // no leading "="
echo "[4] Structured reference\n";
echo "    E1 formula = ", $jq($ws-&gt;getRange("E1:E1")-&gt;getFormula()), "\n";
echo "    E1 value   = ", $jq($ws-&gt;getRange("E1:E1")-&gt;getValue(true)), "\n";

// ---------- 5. comments (modern) and memos (legacy) ----------
$cm = $ws-&gt;getComments();
$mkComment = function (string $author, string $content): \phposbxl\dto\CommentObject {
    $c = new \phposbxl\dto\CommentObject();
    $c-&gt;setAuthor($author);
    $c-&gt;setContent($content);
    return $c;
};

// modern comment (thread): body + reply = 2 entries -&gt; B2 (kept in the file after save)
$cm-&gt;setCommentThread("B2", [$mkComment("alice", "Please check this number"), $mkComment("bob", "Checked")]);
// single modern comment -&gt; D2 (used to demonstrate removal)
$cm-&gt;setComment("D2", $mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -&gt; A2
$cm-&gt;setMemo("A2", "Carol", "Top-priority product", new \phposbxl\dto\FontObject());

$thread = $cm-&gt;getComment("B2");
echo "[5] Comments\n";
echo "    B2 thread count = ", count($thread), "\n";
foreach ($thread as $i =&gt; $c) {
    echo "      [{$i}] ", $c-&gt;getAuthor(), ": ", $c-&gt;getContent(), $c-&gt;getParentId() ? " (reply)" : "", "\n";
}
echo "    D2 comment count = ", count($cm-&gt;getComment("D2")), "\n";
echo "    A2 memo = ", $jq($cm-&gt;getMemoText("A2")), " / author ", $jq($cm-&gt;getMemoAuthor("A2")), "\n";

// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
$cm-&gt;removeComment("D2");
$cm-&gt;removeMemo("A2");
$tryRead = function (string $label, callable $fn): void {
    try {
        $fn();
        echo "    after removal {$label}: no exception\n";
    } catch (\Exception $e) {
        echo "    after removal {$label}: exception ", $e-&gt;getMessage(), "\n";
    }
};
$tryRead("D2 comment", fn() =&gt; $cm-&gt;getComment("D2"));
$tryRead("A2 memo", fn() =&gt; $cm-&gt;getMemoText("A2"));
echo "    B2 is kept, so ", count($cm-&gt;getComment("B2")), " comments remain after save\n";

$total = $ws-&gt;getRange("B5:B5")-&gt;getValue(true)["B5"];
$e1 = $ws-&gt;getRange("E1:E1")-&gt;getValue(true)["E1"];
echo "[Check] table=", $table-&gt;getName(), " total(Unit Price)=", $total, " structuredRef(E1)=", $e1, "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// tables-comments.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("tables-comments-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- fill in the data (1 header row + 3 data rows) ----------
const rows = [
    ["Product", "Unit Price", "Qty"],
    ["Apple", 120, 3],
    ["Orange", 80, 5],
    ["Grape", 300, 2],
];
const values = [];
rows.forEach((row, ri) =&gt; row.forEach((cell, ci) =&gt; {
    const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
    const o = new dto.InputValueObject();
    if (typeof cell === "number") o.setNumberValue(a1, cell);
    else o.setStringValue(a1, cell);
    values.push(o);
}));
ws.setValueArray(values);

// ---------- 1. create the table ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] Table created");
console.log(`    name       = ${table.getName()}`);
console.log(`    columns    = ${JSON.stringify(table.getColumnName())}`);
console.log(`    style      = ${table.getTableStyleName()}`);
console.log(`    full range = ${table.getAllRange().getAddress()}`);
console.log(`    header     = ${table.getHeaderRowRange().getAddress()}`);
console.log(`    data body  = ${table.getDataBodyRange().getAddress()}`);
console.log(`    totals row = ${table.getTotalRowRange().getAddress()}`);

// ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] Totals row");
console.log(`    label (col 1) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(`    formula       = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(`    value         = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);

// ---------- 3. change the table style ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] style after change = ${table.getTableStyleName()}`);

// ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])"); // no leading "="
console.log("[4] Structured reference");
console.log(`    E1 formula = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(`    E1 value   = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);

// ---------- 5. comments (modern) and memos (legacy) ----------
const cm = ws.getComments();
const mkComment = (author, content) =&gt; {
    const c = new dto.CommentObject();
    c.setAuthor(author);
    c.setContent(content);
    return c;
};

// modern comment (thread): body + reply = 2 entries -&gt; B2 (kept in the file after save)
cm.setCommentThread("B2", [mkComment("alice", "Please check this number"), mkComment("bob", "Checked")]);
// single modern comment -&gt; D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -&gt; A2
cm.setMemo("A2", "Carol", "Top-priority product", new dto.FontObject());

const thread = cm.getComment("B2");
console.log("[5] Comments");
console.log(`    B2 thread count = ${thread.length}`);
thread.forEach((c, i) =&gt; console.log(`      [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (reply)" : ""}`));
console.log(`    D2 comment count = ${cm.getComment("D2").length}`);
console.log(`    A2 memo = ${JSON.stringify(cm.getMemoText("A2"))} / author ${JSON.stringify(cm.getMemoAuthor("A2"))}`);

// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) =&gt; {
    try { fn(); console.log(`    after removal ${label}: no exception`); }
    catch (e) {
        // embind exception messages carry a prefix like "Comments.getComment: ";
        // strip it so the output matches the Node.js version exactly
        const msg = String(e.message || e).replace(/^[A-Za-z_$][\w$]*\.[\w$]+: /, "");
        console.log(`    after removal ${label}: exception ${msg}`);
    }
};
tryRead("D2 comment", () =&gt; cm.getComment("D2"));
tryRead("A2 memo", () =&gt; cm.getMemoText("A2"));
console.log(`    B2 is kept, so ${cm.getComment("B2").length} comments remain after save`);

const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[Check] table=${table.getName()} total(Unit Price)=${total} structuredRef(E1)=${e1}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] Table created
    name       = Sales
    columns    = ["Product","Unit Price","Qty"]
    style      = TableStyleMedium2
    full range = A1:C5
    header     = A1:C1
    data body  = A2:C4
    totals row = A5:C5
[2] Totals row
    label (col 1) = "Total"
    formula       = {"B5":"SUBTOTAL(109,Sales[Unit Price])","C5":"SUBTOTAL(109,Sales[Qty])"}
    value         = {"B5":"500","C5":"10"}
[3] style after change = TableStyleMedium9
[4] Structured reference
    E1 formula = {"E1":"SUM(Sales[Unit Price])"}
    E1 value   = {"E1":"500"}
[5] Comments
    B2 thread count = 2
      [0] alice: Please check this number
      [1] bob: Checked (reply)
    D2 comment count = 1
    A2 memo = "Top-priority product" / author "Carol"
    after removal D2 comment: exception comment not found
    after removal A2 memo: exception memo not found
    B2 is kept, so 2 comments remain after save
[Check] table=Sales total(Unit Price)=500 structuredRef(E1)=500</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Tables are handled by name, not by range.</strong> From the return value of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList()</code> (a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Table</code>) or from <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getList("Sales")</code> you can pull out the header, data body and totals row ranges as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> objects.<br>Because a totals row was inserted, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAllRange()</code> is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A1:C5</code> (data <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A1:C4</code> + totals row 5).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The totals row holds <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL</code> formulas.</strong> Setting <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowFunction(2, ...Sum)</code> auto-generates <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(109,Sales[Unit Price])</code> in B5 (109 = sum, nothing ignored). Average maps to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(101,...)</code> and Count to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(103,...)</code> (measured). To put an arbitrary formula there, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomRowFunction(col, "MAX(C2:C4)")</code>; in that case <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowFunction()</code> returns <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCustom</code>(9).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Structured references are written as "TableName[ColumnName]".</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUM(Sales[Unit Price])</code> evaluates to 500. A bare column name (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUM(Unit Price)</code>) gives a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#NAME?</code> error.<br>Formulas do not take a leading <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Renaming a column also updates structured references inside the table.</strong> When you change column names with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnName([...])</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sales[OldName]</code> in the totals row and in existing formulas is automatically replaced with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sales[NewName]</code> (measured).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>There are two comment systems.</strong> The <strong>modern</strong> one with threads/replies/done flag (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComment</code>) and the legacy Excel callout <strong>memo</strong> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMemo</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoText</code>). Pick whichever fits the use case.<br>Both are persisted to the file by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code>.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Table names that look like cell references are not allowed.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList("A1", ...)</code> or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList("S2", ...)</code> throws a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">table name is invalid format</code> exception.<br>Re-creating a table with an existing name throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">table name is already exists</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>With <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList(insertTotals=true)</code> the default totals are "None" for every column (Excel-compatible, fixed 2026-10-01).</strong> Only the hardcoded label <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">集計</code> ("Total" in Japanese) is placed automatically in the first column — set your own with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowLabel()</code>. No totals function is set for the data columns (older versions defaulted the last column to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Count</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(103)</code>)). Set the columns you want totaled explicitly with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowFunction()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList()</code> deletes the data by default.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList("Sales")</code> behaves as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteData=true</code> and wipes the cell values too.<br>To remove only the table outline and keep the values, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList("Sales", false)</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Modern comments require a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CommentObject</code> instance.</strong> A <strong>plain literal like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment("A1", { text: "..." })</code> from the docs' <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">@example</code> throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Invalid argument</code></strong> (measured). Build the object with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new dto.CommentObject()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAuthor()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setContent()</code> and pass that.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment</code> appends, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread</code> overwrites.</strong> Calling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment</code> on a cell that already has comments increases the count, and the new one automatically becomes a reply (with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">parentId</code>) in the existing thread.<br>Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread</code> when you want to replace the whole thread (measured).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Reading a nonexistent comment/memo throws.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComment()</code> throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">comment not found</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoText()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoAuthor()</code> throw <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">memo not found</code> (fixed 2026-10-01). Verify state after removal by catching the exception in a try/catch.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Averages etc. in the totals row come back as raw floating point.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> returns pre-rounding values as strings, so an average <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(101,...)</code> can be something like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">3.3333333333333335</code> (round it via the display format).</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">AutoFilter and Sorting</a> / Next: <a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">Conditional Formatting</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/en/nodejs/classes/dto.CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/en/nodejs/enums/enums.XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/en/nodejs/enums/enums.XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/en/python/pyosbxl.Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/en/python/pyosbxl.Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/en/python/pyosbxl.dto.CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/ListObjectsWrapper.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/TableWrapper.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/CommentsWrapper.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/CommentObjectWrapper.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/en/php/classes/phposbxl-Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/en/php/classes/phposbxl-Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/en/wasm/classes/osbxl.Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/en/wasm/classes/osbxl.Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/en/wasm/classes/dto.CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/en/wasm/enums/enums.XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/en/wasm/enums/enums.XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Tables and Comments — create, totals, structured refs",
   "description": "Tables and Comments — create, totals, structured refs. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-08-tables-comments.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-08-tables-comments.php",
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
     "name": "Tables and Comments — create, totals, structured refs",
     "item": "https://osb-office.com/en/sample/general-08-tables-comments.php"
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
