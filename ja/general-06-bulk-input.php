<?php
$page_title = '大量データの一括投入 — setValueArray / setFormulaArray | osbxl サンプル';
$no_tagline = true;
$meta_description = '大量データの一括投入 — setValueArray / setFormulaArray。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,setvaluearray,setformulaarray,inputvalueobject,inputformulaobject,bulk,performance,dto';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">大量データの一括投入 — setValueArray / setFormulaArray</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-06-bulk-input.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">何千・何万セルもの値や数式を、セルごとに <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange()</code> するのではなく <strong>配列にまとめて一度に投入する</strong>。あわせて相対参照を行ごとにずらした数式を一括入力する。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">やり方</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">使うもの</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">セルごとに投入</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getCells(r, c).setNumberValue(...)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>値を一括投入</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.setValueArray(Array&lt;dto.InputValueObject&gt;)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>数式を一括投入</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.setFormulaArray(Array&lt;dto.InputFormulaObject&gt;)</code></td></tr>
  </tbody></table></div>
<h2 id="対象クラス" class="text-2xl font-bold mt-10 mb-4">対象クラス</h2>
<h3 id="dto-inputvalueobject" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputValueObject</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setStringValue(A1C1, val, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">文字列値を設定します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(A1C1, val, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">数値を設定します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(A1C1, val, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">真偽値を設定します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeValue(A1C1, val, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">日付時刻値を設定します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setEmptyVal(A1C1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル範囲を空に設定します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getA1C1()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getInputType()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getStringValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumberValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBooleanValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDateTimeValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumberFormat()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getForceString()</code></td><td class="border border-surface-200 px-3 py-2 align-top">設定内容の取得。</td></tr>
  </tbody></table></div>
<h3 id="dto-inputformulaobject" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputFormulaObject</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1(A1C1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">対象セルを設定します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula(formula)</code></td><td class="border border-surface-200 px-3 py-2 align-top">数式を設定します（先頭に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> は付けない）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getA1C1()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code></td><td class="border border-surface-200 px-3 py-2 align-top">設定内容の取得。</td></tr>
  </tbody></table></div>
<h2 id="コード" class="text-2xl font-bold mt-10 mb-4">コード</h2>
<p class="mb-4 text-slate-700 leading-7">実行結果:</p>
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
        <pre><code class="language-javascript">// bulk-input.js — 値と数式の一括投入、セルごと投入との速度比較
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. いろいろな型を一度に投入する ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);

const values = [];
const addString = (a1, val) =&gt; {
    const o = new dto.InputValueObject();
    o.setStringValue(a1, val);
    values.push(o);
};
// 末尾の省略可能引数は undefined のまま渡せます（省略として扱われます）
const addNumber = (a1, val, fmt) =&gt; {
    const o = new dto.InputValueObject();
    o.setNumberValue(a1, val, false, fmt);
    values.push(o);
};

addString("A1", "商品名");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd");   // 書式を明示
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date);                       // 省略 → 自動
values.push(dateVal2);

ws.setValueArray(values);
console.log("[1] 型混在の一括投入");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`    ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])} 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
}

// ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] setEmptyVal 後 A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));

// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for (let r = 1; r &lt;= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r * 10);
const formulas = [];
for (let r = 1; r &lt;= 3; r++) {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(r, 7));   // G1, G2, G3
    f.setFormula(`F${r}*2`);                        // 行ごとに異なる相対参照
    formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] setFormulaArray 後");
console.log("    数式 =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log("    値   =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));

// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
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

console.log(`[4] ${ROWS * COLS} セルの投入`);
console.log(`    セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setValueArray (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === b ? a : `不一致 ${a} / ${b}`}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># bulk-input.py — node 版サンプルと同一検証 (pyosbxl)
import json
import time

import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("bulk-input-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. いろいろな型を一度に投入する ----------
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 3, 20)

values = []

def add_string(a1, val):
    o = pyosbxl.dto.InputValueObject()
    o.setStringValue(a1, val)
    values.append(o)

# 末尾の省略可能引数は渡さなければ省略として扱われます
def add_number(a1, val, fmt=None):
    o = pyosbxl.dto.InputValueObject()
    if fmt is None:
        o.setNumberValue(a1, val)
    else:
        o.setNumberValue(a1, val, False, fmt)
    values.append(o)

add_string("A1", "商品名")
add_number("B1", 1234.5)
add_number("C1", 0.25, "0.0%")
# 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
date_val = pyosbxl.dto.InputValueObject()
date_val.setDateTimeValue("D1", date, False, "yyyy-mm-dd")   # 書式を明示
values.append(date_val)
date_val2 = pyosbxl.dto.InputValueObject()
date_val2.setDateTimeValue("D2", date)                       # 省略 → 自動
values.append(date_val2)

ws.setValueArray(values)
print("[1] 型混在の一括投入")
for a1 in ["A1", "B1", "C1", "D1", "D2"]:
    r = ws.getRange(f"{a1}:{a1}")
    v = json.dumps(r.getValue()[a1], ensure_ascii=False)
    fmt = json.dumps(r.getNumberFormat()[a1], ensure_ascii=False)
    print(f"    {a1}: 表示値={v} 書式={fmt}")

# ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力")
empty = pyosbxl.dto.InputValueObject()
empty.setEmptyVal("A2")
ws.setValueArray([empty])
print("[2] setEmptyVal 後 A2 =", json.dumps(ws.getRange("A2:A2").getValue()["A2"], ensure_ascii=False))

# ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
# Range.setFormula() では相対参照が展開されないので、こちらを使う
for r in range(1, 4):
    ws.getRange(f"F{r}:F{r}").setNumberValue(r * 10)
formulas = []
for r in range(1, 4):
    f = pyosbxl.dto.InputFormulaObject()
    f.setA1C1(app.convertFromRowColNumber(r, 7))   # G1, G2, G3
    f.setFormula(f"F{r}*2")                        # 行ごとに異なる相対参照（先頭の "=" は付けない）
    formulas.append(f)
ws.setFormulaArray(formulas)
print("[3] setFormulaArray 後")
fm = ws.getRange("G1:G3").getFormula()
vm = ws.getRange("G1:G3").getValue(True)
keys = ["G1", "G2", "G3"]
print("    数式 =", json.dumps({k: fm[k] for k in keys}, ensure_ascii=False, separators=(",", ":")))
print("    値   =", json.dumps({k: vm[k] for k in keys}, ensure_ascii=False, separators=(",", ":")))

# ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
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

print(f"[4] {ROWS * COLS} セルの投入")
print(f"    セルごと getCells().setNumberValue() : {ms_cell} ms")
print(f"    setValueArray (構築{ms_build} + 投入{ms_apply}) : {ms_build + ms_apply} ms")

a = ws_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[検証] 両方式とも E2000 = {a if a == b else f'不一致 {a} / {b}'}")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// bulk-input.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    // JSON.stringify 相当の引用符付き文字列
    static String q(String s) { return "\"" + s + "\""; }

    static void addString(List&lt;InputValueObjectWrapper&gt; values, String a1, String val) {
        InputValueObjectWrapper o = new InputValueObjectWrapper();
        o.setStringValue(a1, val);
        values.add(o);
    }

    // fmt == null なら省略可能引数を渡さない（省略として扱われる）
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

        // ---------- 1. いろいろな型を一度に投入する ----------
        DateTimeObjectWrapper date = new DateTimeObjectWrapper();
        date.setYMD(2024, 3, 20);

        List&lt;InputValueObjectWrapper&gt; values = new ArrayList&lt;&gt;();
        addString(values, "A1", "商品名");
        addNumber(values, "B1", 1234.5, null);
        addNumber(values, "C1", 0.25, "0.0%");
        // 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
        InputValueObjectWrapper dateVal = new InputValueObjectWrapper();
        dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd");   // 書式を明示
        values.add(dateVal);
        InputValueObjectWrapper dateVal2 = new InputValueObjectWrapper();
        dateVal2.setDateTimeValue("D2", date);                       // 省略 → 自動
        values.add(dateVal2);

        ws.setValueArray(values);
        System.out.println("[1] 型混在の一括投入");
        for (String a1 : new String[]{"A1", "B1", "C1", "D1", "D2"}) {
            RangeWrapper r = ws.getRange(a1 + ":" + a1);
            System.out.println("    " + a1 + ": 表示値=" + q(r.getValue().get(a1))
                    + " 書式=" + q(r.getNumberFormat().get(a1)));
        }

        // ---------- 2. 空にするのも一括投入 ----------
        ws.getRange("A2:A2").setValue("いったん入力");
        InputValueObjectWrapper empty = new InputValueObjectWrapper();
        empty.setEmptyVal("A2");
        List&lt;InputValueObjectWrapper&gt; emptyList = new ArrayList&lt;&gt;();
        emptyList.add(empty);
        ws.setValueArray(emptyList);
        System.out.println("[2] setEmptyVal 後 A2 = " + q(ws.getRange("A2:A2").getValue().get("A2")));

        // ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
        // Range.setFormula() では相対参照が展開されないので、こちらを使う
        for (int r = 1; r &lt;= 3; r++) ws.getRange("F" + r + ":F" + r).setNumberValue(r * 10);
        List&lt;InputFormulaObjectWrapper&gt; formulas = new ArrayList&lt;&gt;();
        for (int r = 1; r &lt;= 3; r++) {
            InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
            f.setA1C1(app.convertFromRowColNumber(r, 7));   // G1, G2, G3
            f.setFormula("F" + r + "*2");                   // 行ごとに異なる相対参照（先頭の "=" は付けない）
            formulas.add(f);
        }
        ws.setFormulaArray(formulas);
        System.out.println("[3] setFormulaArray 後");
        Map&lt;String, String&gt; fm = ws.getRange("G1:G3").getFormula();
        Map&lt;String, String&gt; vm = ws.getRange("G1:G3").getValue(true);
        System.out.println("    数式 = {\"G1\":" + q(fm.get("G1")) + ",\"G2\":" + q(fm.get("G2"))
                + ",\"G3\":" + q(fm.get("G3")) + "}");
        System.out.println("    値   = {\"G1\":" + q(vm.get("G1")) + ",\"G2\":" + q(vm.get("G2"))
                + ",\"G3\":" + q(vm.get("G3")) + "}");

        // ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
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

        System.out.println("[4] " + (ROWS * COLS) + " セルの投入");
        System.out.println("    セルごと getCells().setNumberValue() : " + msCell + " ms");
        System.out.println("    setValueArray (構築" + msBuild + " + 投入" + msApply + ") : "
                + (msBuild + msApply) + " ms");

        String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
        String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
        System.out.println("[検証] 両方式とも E2000 = " + (a.equals(b) ? a : "不一致 " + a + " / " + b));

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// bulk-input.php — 値と数式の一括投入、セルごと投入との速度比較 (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("bulk-input-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// JSON.stringify 相当の出力（非 ASCII はそのまま）
$jq = fn($x): string =&gt; json_encode($x, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
$now = fn(): int =&gt; (int)(hrtime(true) / 1000000);   // ミリ秒

// ---------- 1. いろいろな型を一度に投入する ----------
$date = new \phposbxl\dto\DateTimeObject();
$date-&gt;setYMD(2024, 3, 20);

$values = [];
$addString = function (string $a1, string $val) use (&amp;$values): void {
    $o = new \phposbxl\dto\InputValueObject();
    $o-&gt;setStringValue($a1, $val);
    $values[] = $o;
};
// 省略可能引数は null のとき引数の少ない呼び出しを使う（省略として扱われる）
$addNumber = function (string $a1, float $val, ?string $fmt = null) use (&amp;$values): void {
    $o = new \phposbxl\dto\InputValueObject();
    if ($fmt === null) {
        $o-&gt;setNumberValue($a1, $val);
    } else {
        $o-&gt;setNumberValue($a1, $val, false, $fmt);
    }
    $values[] = $o;
};

$addString("A1", "商品名");
$addNumber("B1", 1234.5);
$addNumber("C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
$dateVal = new \phposbxl\dto\InputValueObject();
$dateVal-&gt;setDateTimeValue("D1", $date, false, "yyyy-mm-dd");   // 書式を明示
$values[] = $dateVal;
$dateVal2 = new \phposbxl\dto\InputValueObject();
$dateVal2-&gt;setDateTimeValue("D2", $date);                       // 省略 → 自動
$values[] = $dateVal2;

$ws-&gt;setValueArray($values);
echo "[1] 型混在の一括投入\n";
foreach (["A1", "B1", "C1", "D1", "D2"] as $a1) {
    $r = $ws-&gt;getRange("{$a1}:{$a1}");
    echo "    {$a1}: 表示値=", $jq($r-&gt;getValue()[$a1]), " 書式=", $jq($r-&gt;getNumberFormat()[$a1]), "\n";
}

// ---------- 2. 空にするのも一括投入 ----------
$ws-&gt;getRange("A2:A2")-&gt;setValue("いったん入力");
$empty = new \phposbxl\dto\InputValueObject();
$empty-&gt;setEmptyVal("A2");
$ws-&gt;setValueArray([$empty]);
echo "[2] setEmptyVal 後 A2 = ", $jq($ws-&gt;getRange("A2:A2")-&gt;getValue()["A2"]), "\n";

// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for ($r = 1; $r &lt;= 3; $r++) {
    $ws-&gt;getRange("F{$r}:F{$r}")-&gt;setNumberValue($r * 10);
}
$formulas = [];
for ($r = 1; $r &lt;= 3; $r++) {
    $f = new \phposbxl\dto\InputFormulaObject();
    $f-&gt;setA1C1($app-&gt;convertFromRowColNumber($r, 7));   // G1, G2, G3
    $f-&gt;setFormula("F{$r}*2");                           // 行ごとに異なる相対参照（先頭の "=" は付けない）
    $formulas[] = $f;
}
$ws-&gt;setFormulaArray($formulas);
echo "[3] setFormulaArray 後\n";
echo "    数式 = ", $jq($ws-&gt;getRange("G1:G3")-&gt;getFormula()), "\n";
echo "    値   = ", $jq($ws-&gt;getRange("G1:G3")-&gt;getValue(true)), "\n";

// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
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

echo "[4] ", $ROWS * $COLS, " セルの投入\n";
echo "    セルごと getCells().setNumberValue() : {$msCell} ms\n";
echo "    setValueArray (構築{$msBuild} + 投入{$msApply}) : ", $msBuild + $msApply, " ms\n";

$a = $wsCell-&gt;getRange("E2000:E2000")-&gt;getValue(true)["E2000"];
$b = $wsBulk-&gt;getRange("E2000:E2000")-&gt;getValue(true)["E2000"];
echo "[検証] 両方式とも E2000 = ", $a === $b ? $a : "不一致 {$a} / {$b}", "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// bulk-input.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("bulk-input-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. いろいろな型を一度に投入する ----------
const date = new dto.DateTimeObject();
date.setYMD(2024, 3, 20);

const values = [];
const addString = (a1, val) =&gt; {
    const o = new dto.InputValueObject();
    o.setStringValue(a1, val);
    values.push(o);
};
// wasm 版は引数ごとに別オーバーロードが登録されており、node と違い
// 末尾の省略可能引数へ undefined を渡すと "Cannot pass non-string" 例外になる。
// 省略したいときは引数の少ないオーバーロードを呼ぶ。
const addNumber = (a1, val, fmt) =&gt; {
    const o = new dto.InputValueObject();
    if (fmt === undefined) o.setNumberValue(a1, val);
    else o.setNumberValue(a1, val, false, fmt);
    values.push(o);
};

addString("A1", "商品名");
addNumber("B1", 1234.5);
addNumber("C1", 0.25, "0.0%");
// 日付: 書式を省略すると既定書式が自動設定されます（Range.setDateValue と同じ）
const dateVal = new dto.InputValueObject();
dateVal.setDateTimeValue("D1", date, false, "yyyy-mm-dd");   // 書式を明示
values.push(dateVal);
const dateVal2 = new dto.InputValueObject();
dateVal2.setDateTimeValue("D2", date);                       // 省略 → 自動
values.push(dateVal2);

ws.setValueArray(values);
console.log("[1] 型混在の一括投入");
for (const a1 of ["A1", "B1", "C1", "D1", "D2"]) {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`    ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])} 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
}

// ---------- 2. 空にするのも一括投入 ----------
ws.getRange("A2:A2").setValue("いったん入力");
const empty = new dto.InputValueObject();
empty.setEmptyVal("A2");
ws.setValueArray([empty]);
console.log("[2] setEmptyVal 後 A2 =", JSON.stringify(ws.getRange("A2:A2").getValue()["A2"]));

// ---------- 3. 相対参照を行ごとにずらした数式を一括入力 ----------
// Range.setFormula() では相対参照が展開されないので、こちらを使う
for (let r = 1; r &lt;= 3; r++) ws.getRange(`F${r}:F${r}`).setNumberValue(r * 10);
const formulas = [];
for (let r = 1; r &lt;= 3; r++) {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(r, 7));   // G1, G2, G3
    f.setFormula(`F${r}*2`);                        // 行ごとに異なる相対参照
    formulas.push(f);
}
ws.setFormulaArray(formulas);
console.log("[3] setFormulaArray 後");
console.log("    数式 =", JSON.stringify(ws.getRange("G1:G3").getFormula()));
console.log("    値   =", JSON.stringify(ws.getRange("G1:G3").getValue(true)));

// ---------- 4. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
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

console.log(`[4] ${ROWS * COLS} セルの投入`);
console.log(`    セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setValueArray (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === b ? a : `不一致 ${a} / ${b}`}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] 型混在の一括投入
    A1: 表示値="商品名" 書式="General"
    B1: 表示値="1234.5" 書式="General"
    C1: 表示値="25.0%" 書式="0.0%"
    D1: 表示値="2024-03-20" 書式="yyyy-mm-dd"
    D2: 表示値="2024/3/20" 書式="yyyy/m/d;@"
[2] setEmptyVal 後 A2 = ""
[3] setFormulaArray 後
    数式 = {"G1":"F1*2","G2":"F2*2","G3":"F3*2"}
    値   = {"G1":"20","G2":"40","G3":"60"}
[4] 10000 セルの投入
    セルごと getCells().setNumberValue() : 102 ms
    setValueArray (構築18 + 投入58) : 76 ms
[検証] 両方式とも E2000 = 200005</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[4]</code> の ms は実行のたびに変わります（上記は 1 回の実行例）。</p>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<h3 id="速度" class="text-xl font-bold mt-8 mb-3">速度</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">セル数</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">セルごと <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells().setNumberValue()</code></th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray()</code></th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">差</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">10,000（2,000 行 x 5 列）</td><td class="border border-surface-200 px-3 py-2 align-top">約 110 ms</td><td class="border border-surface-200 px-3 py-2 align-top">76 ms（構築 18 + 投入 58）</td><td class="border border-surface-200 px-3 py-2 align-top">約 1.4 倍</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">100,000（20,000 行 x 5 列）</td><td class="border border-surface-200 px-3 py-2 align-top">1,390 ms</td><td class="border border-surface-200 px-3 py-2 align-top">770 ms（構築 130 + 投入 640）</td><td class="border border-surface-200 px-3 py-2 align-top">約 1.8 倍</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">いずれも osbxl 1.5.0 / Node.js v21 での 1 回の実測例で、環境と実行のたびに変わります。 傾向として<strong>セル数が増えるほど差が開きます</strong>。 <strong>JavaScript 側で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InputValueObject</code> を組み立てる時間も無視できない</strong>ので、 上記では「構築」と「投入」を分けて計測しています。 数万セルを超えるなら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray()</code> を使ってください。</p>
<h3 id="使いどころ" class="text-xl font-bold mt-8 mb-3">使いどころ</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray()</code></strong> — 配列や CSV、DB の結果など「まとめて入れたい値」があるとき。<br> 型（文字列 / 数値 / 真偽値 / 日付）を 1 つの配列に混在させられます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray()</code></strong> — <strong>行ごとに異なる相対参照の数式</strong>を入れるとき。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range.setFormula()</code> では相対参照が展開されない（<a href="general-01-overview.php#注意点" class="text-primary-600 hover:underline">概要の注意点</a>参照）ため、 計算列を作るならこちらが正解です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">セル数が数十程度なら、可読性を優先して <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells(r, c)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange("A1:A1")</code> で十分です。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>素のオブジェクトリテラルは渡せません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputValueObject</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputFormulaObject</code> の インスタンスを作る必要があります。<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">// ✗ 例外: Invalid argument
ws.setValueArray([{ cell: "A1", value: 100 }]);

// ✓
const o = new dto.InputValueObject();
o.setNumberValue("A1", 100);
ws.setValueArray([o]);</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeValue()</code> は表示書式を自動設定します</strong>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range.setDateValue()</code> と対称）。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">numberFormat</code> を省略すると内容に応じて <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d;@</code>（日付のみ）/ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">h:mm:ss;@</code>（時刻のみ）/ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d h:mm:ss;@</code>（日付+時刻）が自動で付きます。<br> 特定の書式にしたいときだけ第4引数で渡します。<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-plaintext">setDateTimeValue("D2", date)                      → 表示値 "2024/3/20"   書式 "yyyy/m/d;@"
setDateTimeValue("D1", date, false, "yyyy-mm-dd") → 表示値 "2024-03-20" 書式 "yyyy-mm-dd"</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>同じセルを複数回指定すると、配列の後ろの要素が勝ちます</strong>（実測）。<br> 意図しない上書きが起きないよう、投入前に A1C1 の重複が無いことを確認してください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InputFormulaObject</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1()</code> が必須。</strong> 数式だけ設定してもセルが指定されていなければ反映されません。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray()</code> の数式にも先頭の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> は付けません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">f.setFormula("F1*2")</code> です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code> は「末尾の」省略可能引数にだけ渡せます</strong>（省略として扱われます）。<br> 途中の位置の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code> や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">null</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">unsupported arguments</code> 例外になります。<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">// ✓ 末尾 undefined = 省略
o.setNumberValue("A1", 100, false, undefined);

// ✗ 例外: 途中の undefined
o.setNumberValue("A1", undefined, false, "#,##0");

// ✗ 例外: null は省略として扱われない
o.setNumberValue("A1", 100, null);</code></pre></div>これは osbxl 全体（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> の setter など含む）に共通の仕様です。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">行・列・セル範囲の操作</a> / 次: <a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">オートフィルターと並べ替え</a></li>
<li><div class="mt-2 mb-2">
  <div class="text-sm font-bold mb-2">API リファレンス</div>
  <div class="tab-group">
    <div class="flex gap-2 mb-3 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
    </div>
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/dto.InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/nodejs/classes/dto.InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.dto.InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/python/pyosbxl.dto.InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/dto/InputValueObjectWrapper.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/InputFormulaObjectWrapper.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-dto-InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/php/classes/phposbxl-dto-InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/dto.InputValueObject.html" class="text-primary-600 hover:underline">InputValueObject</a> / <a href="/api-docs/wasm/classes/dto.InputFormulaObject.html" class="text-primary-600 hover:underline">InputFormulaObject</a> / <a href="/api-docs/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "大量データの一括投入 — setValueArray / setFormulaArray",
   "description": "大量データの一括投入 — setValueArray / setFormulaArray。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-06-bulk-input.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-06-bulk-input.php",
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
     "name": "ホーム",
     "item": "https://osb-office.com/"
    },
    {
     "@type": "ListItem",
     "position": 2,
     "name": "サンプル集",
     "item": "https://osb-office.com/sample/"
    },
    {
     "@type": "ListItem",
     "position": 3,
     "name": "大量データの一括投入 — setValueArray / setFormulaArray",
     "item": "https://osb-office.com/sample/general-06-bulk-input.php"
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

<?php include __DIR__ . '/../includes/footer.php'; ?>
