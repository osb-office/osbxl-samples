<?php
$lang = 'en';
$page_title = 'Bulk Input — setValueArray / setFormulaArray | osbxl Samples';
$no_tagline = true;
$meta_description = 'Bulk Input — setValueArray / setFormulaArray. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,setvaluearray,setformulaarray,inputvalueobject,inputformulaobject,bulk,performance,dto';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Bulk Input — setValueArray / setFormulaArray</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-06-bulk-input.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Write values and formulas into thousands or tens of thousands of cells <strong>in one shot as an array</strong> instead of calling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange()</code> per cell. Also, bulk-input formulas whose relative references shift row by row.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Approach</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">What to use</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Per-cell input</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getCells(r, c).setNumberValue(...)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>Bulk value input</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.setValueArray(Array&lt;dto.InputValueObject&gt;)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>Bulk formula input</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.setFormulaArray(Array&lt;dto.InputFormulaObject&gt;)</code></td></tr>
  </tbody></table></div>
<h2 id="classes-covered" class="text-2xl font-bold mt-10 mb-4">Classes covered</h2>
<h3 id="dto-inputvalueobject" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputValueObject</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setStringValue(A1C1, val, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a string value.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(A1C1, val, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a number.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(A1C1, val, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a boolean.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeValue(A1C1, val, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a date/time value.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setEmptyVal(A1C1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Clears the cell range.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getA1C1()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getInputType()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getStringValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumberValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBooleanValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDateTimeValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumberFormat()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getForceString()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Read back the configured content.</td></tr>
  </tbody></table></div>
<h3 id="dto-inputformulaobject" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputFormulaObject</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1(A1C1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets the target cell.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula(formula)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets the formula (no leading <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code>).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getA1C1()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Read back the configured content.</td></tr>
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
        <pre><code class="language-javascript">// bulk-input.js — bulk input of values and formulas, speed comparison against per-cell input
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. write mixed types in one shot ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);

const values = [];
const addString = (a1, val) =&gt; {
    const o = new dto.InputValueObject();
    o.setStringValue(a1, val);
    values.push(o);
};
// Trailing optional arguments can be passed as undefined (treated as omitted)
const addNumber = (a1, val, fmt) =&gt; {
    const o = new dto.InputValueObject();
    o.setNumberValue(a1, val, false, fmt);
    values.push(o);
};

addString("A1", "Product");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// Date: omitting the format auto-sets a default format (same as Range.setDateValue)
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd");   // explicit format
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date);                       // omitted → auto
values.push(dateVal2);

ws.setValueArray(values);
console.log("[1] bulk input with mixed types");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`    ${a1}: display=${JSON.stringify(r.getValue()[a1])} format=${JSON.stringify(r.getNumberFormat()[a1])}`);
}

// ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] after setEmptyVal, A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));

// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for (let r = 1; r &lt;= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r * 10);
const formulas = [];
for (let r = 1; r &lt;= 3; r++) {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(r, 7));   // G1, G2, G3
    f.setFormula(`F${r}*2`);                        // different relative reference per row
    formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] after setFormulaArray");
console.log("    formulas =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log("    values   =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));

// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
const ROWS = 2000, COLS = 5;

const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r &lt;= ROWS; r++) {
    for (let c = 1; c &lt;= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

const wsBulk = wb.addWorkSheet("Bulk");
const bulk = [];
t = Date.now();
for (let r = 1; r &lt;= ROWS; r++) {
    for (let c = 1; c &lt;= COLS; c++) {
        const o = new dto.InputValueObject();
        o.setNumberValue(app.convertFromRowColNumber(r, c), r * 100 + c);
        bulk.push(o);
    }
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setValueArray(bulk);
const msApply = Date.now() - t;

console.log(`[4] writing ${ROWS * COLS} cells`);
console.log(`    per-cell getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setValueArray (build ${msBuild} + write ${msApply}) : ${msBuild + msApply} ms`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[check] both methods agree: E2000 = ${a === b ? a : `mismatch ${a} / ${b}`}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># bulk-input.py — same verification as the Node.js sample (pyosbxl)
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

def add_string(a1, val):
    o = pyosbxl.dto.InputValueObject()
    o.setStringValue(a1, val)
    values.append(o)

# Trailing optional arguments are treated as omitted when not passed
def add_number(a1, val, fmt=None):
    o = pyosbxl.dto.InputValueObject()
    if fmt is None:
        o.setNumberValue(a1, val)
    else:
        o.setNumberValue(a1, val, False, fmt)
    values.append(o)

add_string("A1", "Product")
add_number("B1", 1234.5)
add_number("C1", 0.25, "0.0%")
# Date: omitting the format auto-sets a default format (same as Range.setDateValue)
date_val = pyosbxl.dto.InputValueObject()
date_val.setDateTimeValue("D1", date, False, "yyyy-mm-dd")   # explicit format
values.append(date_val)
date_val2 = pyosbxl.dto.InputValueObject()
date_val2.setDateTimeValue("D2", date)                       # omitted → auto
values.append(date_val2)

ws.setValueArray(values)
print("[1] bulk input with mixed types")
for a1 in ["A1", "B1", "C1", "D1", "D2"]:
    r = ws.getRange(f"{a1}:{a1}")
    v = json.dumps(r.getValue()[a1], ensure_ascii=False)
    fmt = json.dumps(r.getNumberFormat()[a1], ensure_ascii=False)
    print(f"    {a1}: display={v} format={fmt}")

# ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry")
empty = pyosbxl.dto.InputValueObject()
empty.setEmptyVal("A2")
ws.setValueArray([empty])
print("[2] after setEmptyVal, A2 =", json.dumps(ws.getRange("A2:A2").getValue()["A2"], ensure_ascii=False))

# ---------- 3. bulk-input formulas with relative references shifted per row ----------
# Range.setFormula() does not expand relative references, so use this instead
for r in range(1, 4):
    ws.getRange(f"F{r}:F{r}").setNumberValue(r * 10)
formulas = []
for r in range(1, 4):
    f = pyosbxl.dto.InputFormulaObject()
    f.setA1C1(app.convertFromRowColNumber(r, 7))   # G1, G2, G3
    f.setFormula(f"F{r}*2")                        # different relative reference per row (no leading "=")
    formulas.append(f)
ws.setFormulaArray(formulas)
print("[3] after setFormulaArray")
fm = ws.getRange("G1:G3").getFormula()
vm = ws.getRange("G1:G3").getValue(True)
keys = ["G1", "G2", "G3"]
print("    formulas =", json.dumps({k: fm[k] for k in keys}, ensure_ascii=False, separators=(",", ":")))
print("    values   =", json.dumps({k: vm[k] for k in keys}, ensure_ascii=False, separators=(",", ":")))

# ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
ROWS, COLS = 2000, 5

ws_cell = wb.addWorkSheet("CellByCell")
t = time.perf_counter()
for r in range(1, ROWS + 1):
    for c in range(1, COLS + 1):
        ws_cell.getCells(r, c).setNumberValue(r * 100 + c)
ms_cell = int((time.perf_counter() - t) * 1000)

ws_bulk = wb.addWorkSheet("Bulk")
bulk = []
t = time.perf_counter()
for r in range(1, ROWS + 1):
    for c in range(1, COLS + 1):
        o = pyosbxl.dto.InputValueObject()
        o.setNumberValue(app.convertFromRowColNumber(r, c), r * 100 + c)
        bulk.append(o)
ms_build = int((time.perf_counter() - t) * 1000)
t = time.perf_counter()
ws_bulk.setValueArray(bulk)
ms_apply = int((time.perf_counter() - t) * 1000)

print(f"[4] writing {ROWS * COLS} cells")
print(f"    per-cell getCells().setNumberValue() : {ms_cell} ms")
print(f"    setValueArray (build {ms_build} + write {ms_apply}) : {ms_build + ms_apply} ms")

a = ws_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[check] both methods agree: E2000 = {a if a == b else f'mismatch {a} / {b}'}")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// bulk-input.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    // Quoted string, equivalent to JSON.stringify
    static String q(String s) { return "\"" + s + "\""; }

    static void addString(List&lt;InputValueObjectWrapper&gt; values, String a1, String val) {
        InputValueObjectWrapper o = new InputValueObjectWrapper();
        o.setStringValue(a1, val);
        values.add(o);
    }

    // If fmt == null, do not pass the optional arguments (treated as omitted)
    static void addNumber(List&lt;InputValueObjectWrapper&gt; values, String a1, double val, String fmt) {
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

        List&lt;InputValueObjectWrapper&gt; values = new ArrayList&lt;&gt;();
        addString(values, "A1", "Product");
        addNumber(values, "B1", 1234.5, null);
        addNumber(values, "C1", 0.25, "0.0%");
        // Date: omitting the format auto-sets a default format (same as Range.setDateValue)
        InputValueObjectWrapper dateVal = new InputValueObjectWrapper();
        dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd");   // explicit format
        values.add(dateVal);
        InputValueObjectWrapper dateVal2 = new InputValueObjectWrapper();
        dateVal2.setDateTimeValue("D2", date);                       // omitted → auto
        values.add(dateVal2);

        ws.setValueArray(values);
        System.out.println("[1] bulk input with mixed types");
        for (String a1 : new String[]{"A1", "B1", "C1", "D1", "D2"}) {
            RangeWrapper r = ws.getRange(a1 + ":" + a1);
            System.out.println("    " + a1 + ": display=" + q(r.getValue().get(a1))
                    + " format=" + q(r.getNumberFormat().get(a1)));
        }

        // ---------- 2. clearing cells is bulk input too ----------
        ws.getRange("A2:A2").setValue("temporary entry");
        InputValueObjectWrapper empty = new InputValueObjectWrapper();
        empty.setEmptyVal("A2");
        List&lt;InputValueObjectWrapper&gt; emptyList = new ArrayList&lt;&gt;();
        emptyList.add(empty);
        ws.setValueArray(emptyList);
        System.out.println("[2] after setEmptyVal, A2 = " + q(ws.getRange("A2:A2").getValue().get("A2")));

        // ---------- 3. bulk-input formulas with relative references shifted per row ----------
        // Range.setFormula() does not expand relative references, so use this instead
        for (int r = 1; r &lt;= 3; r++) ws.getRange("F" + r + ":F" + r).setNumberValue(r * 10);
        List&lt;InputFormulaObjectWrapper&gt; formulas = new ArrayList&lt;&gt;();
        for (int r = 1; r &lt;= 3; r++) {
            InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
            f.setA1C1(app.convertFromRowColNumber(r, 7));   // G1, G2, G3
            f.setFormula("F" + r + "*2");                   // different relative reference per row (no leading "=")
            formulas.add(f);
        }
        ws.setFormulaArray(formulas);
        System.out.println("[3] after setFormulaArray");
        Map&lt;String, String&gt; fm = ws.getRange("G1:G3").getFormula();
        Map&lt;String, String&gt; vm = ws.getRange("G1:G3").getValue(true);
        System.out.println("    formulas = {\"G1\":" + q(fm.get("G1")) + ",\"G2\":" + q(fm.get("G2"))
                + ",\"G3\":" + q(fm.get("G3")) + "}");
        System.out.println("    values   = {\"G1\":" + q(vm.get("G1")) + ",\"G2\":" + q(vm.get("G2"))
                + ",\"G3\":" + q(vm.get("G3")) + "}");

        // ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
        final int ROWS = 2000, COLS = 5;

        WorkSheetWrapper wsCell = wb.addWorkSheet("CellByCell");
        long t = System.currentTimeMillis();
        for (int r = 1; r &lt;= ROWS; r++) {
            for (int c = 1; c &lt;= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
        }
        long msCell = System.currentTimeMillis() - t;

        WorkSheetWrapper wsBulk = wb.addWorkSheet("Bulk");
        List&lt;InputValueObjectWrapper&gt; bulk = new ArrayList&lt;&gt;();
        t = System.currentTimeMillis();
        for (int r = 1; r &lt;= ROWS; r++) {
            for (int c = 1; c &lt;= COLS; c++) {
                InputValueObjectWrapper o = new InputValueObjectWrapper();
                o.setNumberValue(app.convertFromRowColNumber(r, c), r * 100 + c);
                bulk.add(o);
            }
        }
        long msBuild = System.currentTimeMillis() - t;
        t = System.currentTimeMillis();
        wsBulk.setValueArray(bulk);
        long msApply = System.currentTimeMillis() - t;

        System.out.println("[4] writing " + (ROWS * COLS) + " cells");
        System.out.println("    per-cell getCells().setNumberValue() : " + msCell + " ms");
        System.out.println("    setValueArray (build " + msBuild + " + write " + msApply + ") : "
                + (msBuild + msApply) + " ms");

        String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
        String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
        System.out.println("[check] both methods agree: E2000 = " + (a.equals(b) ? a : "mismatch " + a + " / " + b));

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// bulk-input.php — bulk input of values and formulas, speed comparison against per-cell input (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("bulk-input-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// JSON.stringify-compatible output (non-ASCII kept as-is)
$jq = fn($x): string =&gt; json_encode($x, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
$now = fn(): int =&gt; (int)(hrtime(true) / 1000000);   // milliseconds

// ---------- 1. write mixed types in one shot ----------
$date = new \phposbxl\dto\DateTimeObject();
$date-&gt;setYMD(2024, 3, 20);

$values = [];
$addString = function (string $a1, string $val) use (&amp;$values): void {
    $o = new \phposbxl\dto\InputValueObject();
    $o-&gt;setStringValue($a1, $val);
    $values[] = $o;
};
// When an optional argument is omitted, call with fewer arguments (treated as omitted)
$addNumber = function (string $a1, float $val, ?string $fmt = null) use (&amp;$values): void {
    $o = new \phposbxl\dto\InputValueObject();
    if ($fmt === null) {
        $o-&gt;setNumberValue($a1, $val);
    } else {
        $o-&gt;setNumberValue($a1, $val, false, $fmt);
    }
    $values[] = $o;
};

$addString("A1", "Product");
$addNumber("B1", 1234.5);
$addNumber("C1", 0.25, "0.0%");
// Date: omitting the format auto-sets a default format (same as Range.setDateValue)
$dateVal = new \phposbxl\dto\InputValueObject();
$dateVal-&gt;setDateTimeValue("D1", $date, false, "yyyy-mm-dd");   // explicit format
$values[] = $dateVal;
$dateVal2 = new \phposbxl\dto\InputValueObject();
$dateVal2-&gt;setDateTimeValue("D2", $date);                       // omitted → auto
$values[] = $dateVal2;

$ws-&gt;setValueArray($values);
echo "[1] bulk input with mixed types\n";
foreach (["A1", "B1", "C1", "D1", "D2"] as $a1) {
    $r = $ws-&gt;getRange("{$a1}:{$a1}");
    echo "    {$a1}: display=", $jq($r-&gt;getValue()[$a1]), " format=", $jq($r-&gt;getNumberFormat()[$a1]), "\n";
}

// ---------- 2. clearing cells is bulk input too ----------
$ws-&gt;getRange("A2:A2")-&gt;setValue("temporary entry");
$empty = new \phposbxl\dto\InputValueObject();
$empty-&gt;setEmptyVal("A2");
$ws-&gt;setValueArray([$empty]);
echo "[2] after setEmptyVal, A2 = ", $jq($ws-&gt;getRange("A2:A2")-&gt;getValue()["A2"]), "\n";

// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for ($r = 1; $r &lt;= 3; $r++) {
    $ws-&gt;getRange("F{$r}:F{$r}")-&gt;setNumberValue($r * 10);
}
$formulas = [];
for ($r = 1; $r &lt;= 3; $r++) {
    $f = new \phposbxl\dto\InputFormulaObject();
    $f-&gt;setA1C1($app-&gt;convertFromRowColNumber($r, 7));   // G1, G2, G3
    $f-&gt;setFormula("F{$r}*2");                           // different relative reference per row (no leading "=")
    $formulas[] = $f;
}
$ws-&gt;setFormulaArray($formulas);
echo "[3] after setFormulaArray\n";
echo "    formulas = ", $jq($ws-&gt;getRange("G1:G3")-&gt;getFormula()), "\n";
echo "    values   = ", $jq($ws-&gt;getRange("G1:G3")-&gt;getValue(true)), "\n";

// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
$ROWS = 2000;
$COLS = 5;

$wsCell = $wb-&gt;addWorkSheet("CellByCell");
$t = $now();
for ($r = 1; $r &lt;= $ROWS; $r++) {
    for ($c = 1; $c &lt;= $COLS; $c++) {
        $wsCell-&gt;getCells($r, $c)-&gt;setNumberValue($r * 100 + $c);
    }
}
$msCell = $now() - $t;

$wsBulk = $wb-&gt;addWorkSheet("Bulk");
$bulk = [];
$t = $now();
for ($r = 1; $r &lt;= $ROWS; $r++) {
    for ($c = 1; $c &lt;= $COLS; $c++) {
        $o = new \phposbxl\dto\InputValueObject();
        $o-&gt;setNumberValue($app-&gt;convertFromRowColNumber($r, $c), $r * 100 + $c);
        $bulk[] = $o;
    }
}
$msBuild = $now() - $t;
$t = $now();
$wsBulk-&gt;setValueArray($bulk);
$msApply = $now() - $t;

echo "[4] writing ", $ROWS * $COLS, " cells\n";
echo "    per-cell getCells().setNumberValue() : {$msCell} ms\n";
echo "    setValueArray (build {$msBuild} + write {$msApply}) : ", $msBuild + $msApply, " ms\n";

$a = $wsCell-&gt;getRange("E2000:E2000")-&gt;getValue(true)["E2000"];
$b = $wsBulk-&gt;getRange("E2000:E2000")-&gt;getValue(true)["E2000"];
echo "[check] both methods agree: E2000 = ", $a === $b ? $a : "mismatch {$a} / {$b}", "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// bulk-input.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
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
const addString = (a1, val) =&gt; {
    const o = new dto.InputValueObject();
    o.setStringValue(a1, val);
    values.push(o);
};
// The wasm build registers a separate overload per argument count; unlike node,
// passing undefined for a trailing optional argument throws a "Cannot pass non-string"
// exception. To omit an argument, call the overload with fewer arguments.
const addNumber = (a1, val, fmt) =&gt; {
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
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd");   // explicit format
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date);                       // omitted → auto
values.push(dateVal2);

ws.setValueArray(values);
console.log("[1] bulk input with mixed types");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`    ${a1}: display=${JSON.stringify(r.getValue()[a1])} format=${JSON.stringify(r.getNumberFormat()[a1])}`);
}

// ---------- 2. clearing cells is bulk input too ----------
ws.getRange("A2:A2").setValue("temporary entry");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] after setEmptyVal, A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));

// ---------- 3. bulk-input formulas with relative references shifted per row ----------
// Range.setFormula() does not expand relative references, so use this instead
for (let r = 1; r &lt;= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r * 10);
const formulas = [];
for (let r = 1; r &lt;= 3; r++) {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(r, 7));   // G1, G2, G3
    f.setFormula(`F${r}*2`);                        // different relative reference per row
    formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] after setFormulaArray");
console.log("    formulas =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log("    values   =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));

// ---------- 4. speed comparison (2,000 rows x 5 cols = 10,000 cells) ----------
const ROWS = 2000, COLS = 5;

const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r &lt;= ROWS; r++) {
    for (let c = 1; c &lt;= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

const wsBulk = wb.addWorkSheet("Bulk");
const bulk = [];
t = Date.now();
for (let r = 1; r &lt;= ROWS; r++) {
    for (let c = 1; c &lt;= COLS; c++) {
        const o = new dto.InputValueObject();
        o.setNumberValue(app.convertFromRowColNumber(r, c), r * 100 + c);
        bulk.push(o);
    }
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setValueArray(bulk);
const msApply = Date.now() - t;

console.log(`[4] writing ${ROWS * COLS} cells`);
console.log(`    per-cell getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setValueArray (build ${msBuild} + write ${msApply}) : ${msBuild + msApply} ms`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[check] both methods agree: E2000 = ${a === b ? a : `mismatch ${a} / ${b}`}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] bulk input with mixed types
    A1: display="Product" format="General"
    B1: display="1234.5" format="General"
    C1: display="25.0%" format="0.0%"
    D1: display="2024-03-20" format="yyyy-mm-dd"
    D2: display="2024/3/20" format="yyyy/m/d;@"
[2] after setEmptyVal, A2 = ""
[3] after setFormulaArray
    formulas = {"G1":"F1*2","G2":"F2*2","G3":"F3*2"}
    values   = {"G1":"20","G2":"40","G3":"60"}
[4] writing 10000 cells
    per-cell getCells().setNumberValue() : 102 ms
    setValueArray (build 18 + write 58) : 76 ms
[check] both methods agree: E2000 = 200005</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7">The ms values in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[4]</code> change with every run (the above is one example run).</p>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<h3 id="speed" class="text-xl font-bold mt-8 mb-3">Speed</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Cells</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Per-cell <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells().setNumberValue()</code></th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray()</code></th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Difference</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">10,000 (2,000 rows x 5 cols)</td><td class="border border-surface-200 px-3 py-2 align-top">~110 ms</td><td class="border border-surface-200 px-3 py-2 align-top">76 ms (build 18 + write 58)</td><td class="border border-surface-200 px-3 py-2 align-top">~1.4x</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">100,000 (20,000 rows x 5 cols)</td><td class="border border-surface-200 px-3 py-2 align-top">1,390 ms</td><td class="border border-surface-200 px-3 py-2 align-top">770 ms (build 130 + write 640)</td><td class="border border-surface-200 px-3 py-2 align-top">~1.8x</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Both rows are single measured runs on osbxl 1.5.0 / Node.js v21; results vary with the environment and between runs. As a trend, <strong>the gap widens as the cell count grows</strong>. **The time spent assembling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InputValueObject</code>s on the JavaScript side is not negligible either**, which is why the measurements above separate "build" from "write". If you need to write tens of thousands of cells or more, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray()</code>.</p>
<h3 id="when-to-use-which" class="text-xl font-bold mt-8 mb-3">When to use which</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray()</code></strong> — when you have "values to insert in bulk", such as an array, CSV, or DB query results.<br>You can mix types (string / number / boolean / date) in a single array.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray()</code></strong> — when inserting <strong>formulas with different relative references per row</strong>. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range.setFormula()</code> does not expand relative references (see <a href="general-01-overview.php#notes" class="text-primary-600 hover:underline">overview caveats</a>), so this is the right tool for building calculated columns.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">For a few dozen cells, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells(r, c)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange("A1:A1")</code> is perfectly fine — favor readability.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Plain object literals are not accepted.</strong> You must create instances of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputValueObject</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputFormulaObject</code>.<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">// ✗ throws: Invalid argument
ws.setValueArray([{ cell: "A1", value: 100 }]);

// ✓
const o = new dto.InputValueObject();
o.setNumberValue("A1", 100);
ws.setValueArray([o]);</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeValue()</code> auto-sets the display format</strong> (symmetric with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range.setDateValue()</code>). Omitting <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">numberFormat</code> automatically applies <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d;@</code> (date only) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">h:mm:ss;@</code> (time only) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d h:mm:ss;@</code> (date + time) depending on the content.<br>Pass the 4th argument only when you want a specific format.<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-plaintext">setDateTimeValue("D2", date)                      → display "2024/3/20"   format "yyyy/m/d;@"
setDateTimeValue("D1", date, false, "yyyy-mm-dd") → display "2024-03-20" format "yyyy-mm-dd"</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>When the same cell is specified multiple times, the element later in the array wins</strong> (measured). To avoid unintended overwrites, make sure there are no duplicate A1C1 values before writing.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InputFormulaObject</code> requires <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1()</code>.</strong> Setting only the formula has no effect if the target cell is not specified.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>No leading <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray()</code> formulas either.</strong> Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">f.setFormula("F1*2")</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code> may only be passed for *trailing* optional arguments** (treated as omitted). <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code> or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">null</code> in a middle position throws an <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">unsupported arguments</code> exception.<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">// ✓ trailing undefined = omitted
o.setNumberValue("A1", 100, false, undefined);

// ✗ throws: undefined in the middle
o.setNumberValue("A1", undefined, false, "#,##0");

// ✗ throws: null is not treated as omitted
o.setNumberValue("A1", 100, null);</code></pre></div>This is a common specification across osbxl (including the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> setters).</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">Rows, columns and ranges</a> / Next: <a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">AutoFilter and Sorting</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/dto.InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/en/nodejs/classes/dto.InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.dto.InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/en/python/pyosbxl.dto.InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/en/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/dto/InputValueObjectWrapper.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/InputFormulaObjectWrapper.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-dto-InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/en/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/dto.InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/en/wasm/classes/dto.InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/en/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Bulk Input — setValueArray / setFormulaArray",
   "description": "Bulk Input — setValueArray / setFormulaArray. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-06-bulk-input.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-06-bulk-input.php",
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
     "name": "Bulk Input — setValueArray / setFormulaArray",
     "item": "https://osb-office.com/en/sample/general-06-bulk-input.php"
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
