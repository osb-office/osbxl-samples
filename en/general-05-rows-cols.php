<?php
$lang = 'en';
$page_title = 'Rows, Columns and Ranges — insert, delete, copy | osbxl Samples';
$no_tagline = true;
$meta_description = 'Rows, Columns and Ranges — insert, delete, copy. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,insertrow,deleterow,clearrow,copyrow,insertcol,deletecol,insertcell,deletecell,clearcell,copycell,row,col,setheight,setcolumnwidth';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Rows, Columns and Ranges — insert, delete, copy</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-05-rows-cols.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Insert / delete / clear / copy rows, columns and cell ranges on a worksheet. Also set row height, column width, hidden state and more with the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col</code> objects.</p>
<h2 id="the-four-operations-compared" class="text-2xl font-bold mt-10 mb-4">The four operations compared</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Operation</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Behavior</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">For rows / columns</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">For cell ranges</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>insert</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Inserts blanks and <strong>shifts</strong> existing data</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertRow(start, n)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCol(start, n)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell(A1C1, direction)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>delete</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Deletes and <strong>shifts</strong> existing data</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteRow(start, n)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCol(start, n)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell(A1C1, direction)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>clear</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Resets contents and formatting to the initial state. <strong>Does not shift</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow(start, n)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCol(start, n)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCell(A1C1)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>copy</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><strong>Overwrite-copies</strong> from another sheet (or the same sheet)</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow(fromSheet, fromRow, n, toRow, copyType?)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell(fromSheet, fromA1C1, toA1, copyType?)</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">The <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell</code> uses <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlInsertDirection</code>.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Value</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Number</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Purpose</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionRight</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">Shift cells right when inserting</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionDown</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">Shift cells down when inserting</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionLeft</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">Shift cells left when deleting</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionUp</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">Shift cells up when deleting</td></tr>
  </tbody></table></div>
<h2 id="code-1-rows-columns-and-cell-ranges" class="text-2xl font-bold mt-10 mb-4">Code 1: Rows, columns and cell ranges</h2>
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
        <pre><code class="language-javascript">// rows-cols.js — insert/delete/clear/copy for rows, columns and cell ranges
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("rows-cols.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// Print a range on one line, ordered by row → column (· marks an empty cell)
const dump = (label, range) =&gt; {
    const v = ws.getRange(range).getValue();
    const keys = Object.keys(v).map((k) =&gt; /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =&gt;
        Number(a[2]) - Number(b[2]) ||
        app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
    console.log(`${label} ${keys.map((m) =&gt; `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// Fill A1:C5 with r{row}c{col}
const fill = () =&gt; {
    for (let r = 1; r &lt;= 5; r++) for (let c = 1; c &lt;= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};

fill();
dump("[initial]     col A:", "A1:A5");

// ---------- rows ----------
ws.insertRow(3, 2);                    // insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow]   col A:", "A1:A7");
ws.deleteRow(3, 2);                    // delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow]   col A:", "A1:A5");
ws.clearRow(2, 1);                     // clear row 2 (positions do not shift)
dump("[clearRow]    col A:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4);         // overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow]     col A:", "A1:A5");

fill();
// ---------- columns ----------
ws.insertCol(2, 1);                    // insert 1 column before column B
dump("[insertCol]   row 1:", "A1:D1");
ws.deleteCol(2, 1);                    // delete the column B just inserted
dump("[deleteCol]   row 1:", "A1:C1");
ws.clearCol(3, 1);                     // clear column C
dump("[clearCol]    row 1:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5);         // copy columns A,B onto columns E,F
dump("[copyCol]     row 1:", "A1:F1");

// ---------- cell ranges ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell]  col B:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell]  col B:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell]   A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell]    F8:G9:", "F8:G9");

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># rows-cols.py — same verification as the Node.js sample (pyosbxl)
import re
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("rows-cols-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# Print a range on one line, ordered by row → column (· marks an empty cell)
def dump(label, rng):
    v = ws.getRange(rng).getValue()
    def keyf(k):
        m = re.match(r"^([A-Z]+)(\d+)$", k)
        return (int(m.group(2)), app.convertToColumnNumber(m.group(1)))
    keys = sorted(v.keys(), key=keyf)
    print(label, " ".join(f"{k}={v[k] or '·'}" for k in keys))

# Fill A1:C5 with r{row}c{col}
def fill():
    for r in range(1, 6):
        for c in range(1, 4):
            ws.getCells(r, c).setValue(f"r{r}c{c}")

fill()
dump("[initial]     col A:", "A1:A5")

# ---------- rows ----------
ws.insertRow(3, 2)                     # insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow]   col A:", "A1:A7")
ws.deleteRow(3, 2)                     # delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow]   col A:", "A1:A5")
ws.clearRow(2, 1)                      # clear row 2 (positions do not shift)
dump("[clearRow]    col A:", "A1:A5")
ws.copyRow("Sheet1", 1, 1, 4)          # overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow]     col A:", "A1:A5")

fill()
# ---------- columns ----------
ws.insertCol(2, 1)                     # insert 1 column before column B
dump("[insertCol]   row 1:", "A1:D1")
ws.deleteCol(2, 1)                     # delete the column B just inserted
dump("[deleteCol]   row 1:", "A1:C1")
ws.clearCol(3, 1)                      # clear column C
dump("[clearCol]    row 1:", "A1:C1")
fill()
ws.copyCol("Sheet1", 1, 2, 5)          # copy columns A,B onto columns E,F
dump("[copyCol]     row 1:", "A1:F1")

# ---------- cell ranges ----------
ws.insertCell("B2:B3", pyosbxl.enums.XlInsertDirection.InsertDirectionDown)
dump("[insertCell]  col B:", "B1:B5")
ws.deleteCell("B2:B3", pyosbxl.enums.XlInsertDirection.DeleteDirectionUp)
dump("[deleteCell]  col B:", "B1:B5")
ws.clearCell("A1:B2")
dump("[clearCell]   A1:B2:", "A1:B2")
ws.copyCell("Sheet1", "A3:B4", "F8")
dump("[copyCell]    F8:G9:", "F8:G9")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// rows-cols.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    static AppWrapper app;
    static WorkSheetWrapper ws;
    static final Pattern ADDR = Pattern.compile("^([A-Z]+)(\\d+)$");

    // Print a range on one line, ordered by row → column (· marks an empty cell)
    static void dump(String label, String range) {
        Map&lt;String, String&gt; v = ws.getRange(range).getValue();
        List&lt;String&gt; keys = new ArrayList&lt;&gt;(v.keySet());
        keys.sort((a, b) -&gt; {
            Matcher ma = ADDR.matcher(a);
            Matcher mb = ADDR.matcher(b);
            ma.matches();
            mb.matches();
            int cmp = Integer.compare(Integer.parseInt(ma.group(2)), Integer.parseInt(mb.group(2)));
            if (cmp != 0) return cmp;
            return Integer.compare(app.convertToColumnNumber(ma.group(1)),
                                   app.convertToColumnNumber(mb.group(1)));
        });
        StringBuilder sb = new StringBuilder();
        for (String k : keys) {
            if (sb.length() &gt; 0) sb.append(" ");
            String val = v.get(k);
            sb.append(k).append("=").append(val == null || val.isEmpty() ? "·" : val);
        }
        System.out.println(label + " " + sb);
    }

    // Fill A1:C5 with r{row}c{col}
    static void fill() {
        for (int r = 1; r &lt;= 5; r++)
            for (int c = 1; c &lt;= 3; c++)
                ws.getCells(r, c).setValue("r" + r + "c" + c);
    }

    public static void main(String[] args) {
        app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("rows-cols-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        fill();
        dump("[initial]     col A:", "A1:A5");

        // ---------- rows ----------
        ws.insertRow(3, 2);                    // insert 2 rows BEFORE row 3 → existing data shifts down
        dump("[insertRow]   col A:", "A1:A7");
        ws.deleteRow(3, 2);                    // delete 2 rows starting at row 3 → rows below move up
        dump("[deleteRow]   col A:", "A1:A5");
        ws.clearRow(2, 1);                     // clear row 2 (positions do not shift)
        dump("[clearRow]    col A:", "A1:A5");
        ws.copyRow("Sheet1", 1, 1, 4);         // overwrite-copy row 1 of the same sheet onto row 4
        dump("[copyRow]     col A:", "A1:A5");

        fill();
        // ---------- columns ----------
        ws.insertCol(2, 1);                    // insert 1 column before column B
        dump("[insertCol]   row 1:", "A1:D1");
        ws.deleteCol(2, 1);                    // delete the column B just inserted
        dump("[deleteCol]   row 1:", "A1:C1");
        ws.clearCol(3, 1);                     // clear column C
        dump("[clearCol]    row 1:", "A1:C1");
        fill();
        ws.copyCol("Sheet1", 1, 2, 5);         // copy columns A,B onto columns E,F
        dump("[copyCol]     row 1:", "A1:F1");

        // ---------- cell ranges ----------
        ws.insertCell("B2:B3", XlInsertDirection.InsertDirectionDown);
        dump("[insertCell]  col B:", "B1:B5");
        ws.deleteCell("B2:B3", XlInsertDirection.DeleteDirectionUp);
        dump("[deleteCell]  col B:", "B1:B5");
        ws.clearCell("A1:B2");
        dump("[clearCell]   A1:B2:", "A1:B2");
        ws.copyCell("Sheet1", "A3:B4", "F8");
        dump("[copyCell]    F8:G9:", "F8:G9");

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// rows-cols.php — insert/delete/clear/copy for rows, columns and cell ranges (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("rows-cols-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// Print a range on one line, ordered by row → column (· marks an empty cell)
$dump = function (string $label, string $range) use ($ws, $app): void {
    $v = $ws-&gt;getRange($range)-&gt;getValue();
    $keys = array_keys($v);
    usort($keys, function (string $a, string $b) use ($app): int {
        preg_match('/^([A-Z]+)(\d+)$/', $a, $ma);
        preg_match('/^([A-Z]+)(\d+)$/', $b, $mb);
        return ((int)$ma[2] &lt;=&gt; (int)$mb[2])
            ?: ($app-&gt;convertToColumnNumber($ma[1]) &lt;=&gt; $app-&gt;convertToColumnNumber($mb[1]));
    });
    $parts = [];
    foreach ($keys as $k) {
        $val = (string)$v[$k];
        $parts[] = $k . "=" . ($val === "" ? "·" : $val);
    }
    echo $label, " ", implode(" ", $parts), "\n";
};
// Fill A1:C5 with r{row}c{col}
$fill = function () use ($ws): void {
    for ($r = 1; $r &lt;= 5; $r++) {
        for ($c = 1; $c &lt;= 3; $c++) {
            $ws-&gt;getCells($r, $c)-&gt;setValue("r{$r}c{$c}");
        }
    }
};

$fill();
$dump("[initial]     col A:", "A1:A5");

// ---------- rows ----------
$ws-&gt;insertRow(3, 2);                    // insert 2 rows BEFORE row 3 → existing data shifts down
$dump("[insertRow]   col A:", "A1:A7");
$ws-&gt;deleteRow(3, 2);                    // delete 2 rows starting at row 3 → rows below move up
$dump("[deleteRow]   col A:", "A1:A5");
$ws-&gt;clearRow(2, 1);                     // clear row 2 (positions do not shift)
$dump("[clearRow]    col A:", "A1:A5");
$ws-&gt;copyRow("Sheet1", 1, 1, 4);         // overwrite-copy row 1 of the same sheet onto row 4
$dump("[copyRow]     col A:", "A1:A5");

$fill();
// ---------- columns ----------
$ws-&gt;insertCol(2, 1);                    // insert 1 column before column B
$dump("[insertCol]   row 1:", "A1:D1");
$ws-&gt;deleteCol(2, 1);                    // delete the column B just inserted
$dump("[deleteCol]   row 1:", "A1:C1");
$ws-&gt;clearCol(3, 1);                     // clear column C
$dump("[clearCol]    row 1:", "A1:C1");
$fill();
$ws-&gt;copyCol("Sheet1", 1, 2, 5);         // copy columns A,B onto columns E,F
$dump("[copyCol]     row 1:", "A1:F1");

// ---------- cell ranges ----------
$ws-&gt;insertCell("B2:B3", \phposbxl\enums\XlInsertDirection::InsertDirectionDown);
$dump("[insertCell]  col B:", "B1:B5");
$ws-&gt;deleteCell("B2:B3", \phposbxl\enums\XlInsertDirection::DeleteDirectionUp);
$dump("[deleteCell]  col B:", "B1:B5");
$ws-&gt;clearCell("A1:B2");
$dump("[clearCell]   A1:B2:", "A1:B2");
$ws-&gt;copyCell("Sheet1", "A3:B4", "F8");
$dump("[copyCell]    F8:G9:", "F8:G9");

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// rows-cols.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("rows-cols-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// Print a range on one line, ordered by row → column (· marks an empty cell)
const dump = (label, range) =&gt; {
    const v = ws.getRange(range).getValue();
    const keys = Object.keys(v).map((k) =&gt; /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =&gt;
        Number(a[2]) - Number(b[2]) ||
        app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
    console.log(`${label} ${keys.map((m) =&gt; `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// Fill A1:C5 with r{row}c{col}
const fill = () =&gt; {
    for (let r = 1; r &lt;= 5; r++) for (let c = 1; c &lt;= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};

fill();
dump("[initial]     col A:", "A1:A5");

// ---------- rows ----------
ws.insertRow(3, 2);                    // insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow]   col A:", "A1:A7");
ws.deleteRow(3, 2);                    // delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow]   col A:", "A1:A5");
ws.clearRow(2, 1);                     // clear row 2 (positions do not shift)
dump("[clearRow]    col A:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4);         // overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow]     col A:", "A1:A5");

fill();
// ---------- columns ----------
ws.insertCol(2, 1);                    // insert 1 column before column B
dump("[insertCol]   row 1:", "A1:D1");
ws.deleteCol(2, 1);                    // delete the column B just inserted
dump("[deleteCol]   row 1:", "A1:C1");
ws.clearCol(3, 1);                     // clear column C
dump("[clearCol]    row 1:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5);         // copy columns A,B onto columns E,F
dump("[copyCol]     row 1:", "A1:F1");

// ---------- cell ranges ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell]  col B:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell]  col B:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell]   A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell]    F8:G9:", "F8:G9");

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[initial]     col A: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[insertRow]   col A: A1=r1c1 A2=r2c1 A3=· A4=· A5=r3c1 A6=r4c1 A7=r5c1
[deleteRow]   col A: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[clearRow]    col A: A1=r1c1 A2=· A3=r3c1 A4=r4c1 A5=r5c1
[copyRow]     col A: A1=r1c1 A2=· A3=r3c1 A4=r1c1 A5=r5c1
[insertCol]   row 1: A1=r1c1 B1=· C1=r1c2 D1=r1c3
[deleteCol]   row 1: A1=r1c1 B1=r1c2 C1=r1c3
[clearCol]    row 1: A1=r1c1 B1=r1c2 C1=·
[copyCol]     row 1: A1=r1c1 B1=r1c2 C1=r1c3 D1=· E1=r1c1 F1=r1c2
[insertCell]  col B: B1=r1c2 B2=· B3=· B4=r2c2 B5=r3c2
[deleteCell]  col B: B1=r1c2 B2=r2c2 B3=r3c2 B4=r4c2 B5=r5c2
[clearCell]   A1:B2: A1=· B1=· A2=· B2=·
[copyCell]    F8:G9: F8=r3c1 G8=r3c2 F9=r4c1 G9=r4c2</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">·</code> marks an empty cell. You can see that <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertRow</code> left rows 3 and 4 blank and shifted the existing data down, that <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow</code> did not shift positions, and that <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> overwrote row 4 with the contents of row 1.</p>
<h2 id="code-2-row-col-objects" class="text-2xl font-bold mt-10 mb-4">Code 2: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col</code> objects</h2>
<p class="mb-4 text-slate-700 leading-7">Output:</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
  </div>
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// row-col-props.js — row height, column width, hidden state, autofit
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("row-col-props.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- Row ----------
const row = ws.getRow(1);                        // row 1
console.log("default height        =", row.getHeight());
row.setHeight(30);
console.log("after setHeight(30)   =", row.getHeight());
console.log("getHidden default     =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true);                    // hide row 2
console.log("row 2 hidden          =", ws.getRow(2).getHidden());

// ---------- Col ----------
const col = ws.getCol(1);                        // column 1 (column A). Column numbers are passed as numbers
console.log("default col width     = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -&gt; columnWidth=" + col.getColumnWidth(),
            "/ width=" + col.getWidth());
ws.getCol(2).setHidden(true);                    // hide column B
console.log("col B hidden          =", ws.getCol(2).getHidden());

ws.getRange("C1:C1").setValue("Sample string for autofit");
console.log("col C bestFit default =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true);                   // autofit the column width to the contents
console.log("setBestFit(true) -&gt; bestFit =", ws.getCol(3).getBestFit(),
            "/ columnWidth =", ws.getCol(3).getColumnWidth());

// Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1);                       // group rows 4-6 at level 1
ws.setColOutline(2, 3, 1);                       // group columns B-C at level 1
console.log("row 4 outlineLevel    =", ws.getRow(4).getOutlineLevel(),
            "/ col B outlineLevel =", ws.getCol(2).getOutlineLevel());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># row-col-props.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("row-col-props-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

def num(x):          # render integer-valued floats as ints (to match the JS output)
    return int(x) if isinstance(x, float) and x.is_integer() else x

def boolstr(x):      # render Python's True/False as JS-style true/false
    return str(x).lower()

# ---------- Row ----------
row = ws.getRow(1)                           # row 1
print("default height        =", num(row.getHeight()))
row.setHeight(30)
print("after setHeight(30)   =", num(row.getHeight()))
print("getHidden default     =", boolstr(row.getHidden()), "/ getOutlineLevel =", row.getOutlineLevel())
ws.getRow(2).setHidden(True)                 # hide row 2
print("row 2 hidden          =", boolstr(ws.getRow(2).getHidden()))

# ---------- Col ----------
col = ws.getCol(1)                           # column 1 (column A). Column numbers are passed as numbers
print("default col width     = columnWidth:", col.getColumnWidth(), "/ width:", num(col.getWidth()))
col.setColumnWidth(20)
print("setColumnWidth(20) -&gt; columnWidth=" + str(num(col.getColumnWidth())),
      "/ width=" + str(num(col.getWidth())))
ws.getCol(2).setHidden(True)                 # hide column B
print("col B hidden          =", boolstr(ws.getCol(2).getHidden()))

ws.getRange("C1:C1").setValue("Sample string for autofit")
print("col C bestFit default =", boolstr(ws.getCol(3).getBestFit()))
ws.getCol(3).setBestFit(True)                # autofit the column width to the contents
print("setBestFit(true) -&gt; bestFit =", boolstr(ws.getCol(3).getBestFit()),
      "/ columnWidth =", ws.getCol(3).getColumnWidth())

# Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1)                    # group rows 4-6 at level 1
ws.setColOutline(2, 3, 1)                    # group columns B-C at level 1
print("row 4 outlineLevel    =", ws.getRow(4).getOutlineLevel(),
      "/ col B outlineLevel =", ws.getCol(2).getOutlineLevel())

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// row-col-props.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;

public class Main {
    // Render integer-valued doubles as ints (to match the JS output)
    static String num(double d) {
        if (d == Math.rint(d) &amp;&amp; !Double.isInfinite(d)) return String.valueOf((long) d);
        return String.valueOf(d);
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("row-col-props-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- Row ----------
        RowWrapper row = ws.getRow(1);                   // row 1
        System.out.println("default height        = " + num(row.getHeight()));
        row.setHeight(30);
        System.out.println("after setHeight(30)   = " + num(row.getHeight()));
        System.out.println("getHidden default     = " + row.getHidden()
                + " / getOutlineLevel = " + row.getOutlineLevel());
        ws.getRow(2).setHidden(true);                    // hide row 2
        System.out.println("row 2 hidden          = " + ws.getRow(2).getHidden());

        // ---------- Col ----------
        ColWrapper col = ws.getCol(1);                   // column 1 (column A). Column numbers are passed as numbers
        System.out.println("default col width     = columnWidth: " + num(col.getColumnWidth())
                + " / width: " + num(col.getWidth()));
        col.setColumnWidth(20);
        System.out.println("setColumnWidth(20) -&gt; columnWidth=" + num(col.getColumnWidth())
                + " / width=" + num(col.getWidth()));
        ws.getCol(2).setHidden(true);                    // hide column B
        System.out.println("col B hidden          = " + ws.getCol(2).getHidden());

        ws.getRange("C1:C1").setValue("Sample string for autofit");
        System.out.println("col C bestFit default = " + ws.getCol(3).getBestFit());
        ws.getCol(3).setBestFit(true);                   // autofit the column width to the contents
        System.out.println("setBestFit(true) -&gt; bestFit = " + ws.getCol(3).getBestFit()
                + " / columnWidth = " + num(ws.getCol(3).getColumnWidth()));

        // Outlines (grouping) are set on the WorkSheet side
        ws.setRowOutline(4, 6, 1);                       // group rows 4-6 at level 1
        ws.setColOutline(2, 3, 1);                       // group columns B-C at level 1
        System.out.println("row 4 outlineLevel    = " + ws.getRow(4).getOutlineLevel()
                + " / col B outlineLevel = " + ws.getCol(2).getOutlineLevel());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// row-col-props.php — row height, column width, hidden state, autofit (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("row-col-props-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// echo of a PHP bool prints 1/empty string, so render JS-style true/false
$boolstr = fn(bool $b): string =&gt; $b ? "true" : "false";

// ---------- Row ----------
$row = $ws-&gt;getRow(1);                        // row 1
echo "default height        = ", $row-&gt;getHeight(), "\n";
$row-&gt;setHeight(30);
echo "after setHeight(30)   = ", $row-&gt;getHeight(), "\n";
echo "getHidden default     = ", $boolstr($row-&gt;getHidden()), " / getOutlineLevel = ", $row-&gt;getOutlineLevel(), "\n";
$ws-&gt;getRow(2)-&gt;setHidden(true);              // hide row 2
echo "row 2 hidden          = ", $boolstr($ws-&gt;getRow(2)-&gt;getHidden()), "\n";

// ---------- Col ----------
$col = $ws-&gt;getCol(1);                        // column 1 (column A). Column numbers are passed as numbers
echo "default col width     = columnWidth: ", $col-&gt;getColumnWidth(), " / width: ", $col-&gt;getWidth(), "\n";
$col-&gt;setColumnWidth(20);
echo "setColumnWidth(20) -&gt; columnWidth=", $col-&gt;getColumnWidth(),
    " / width=", $col-&gt;getWidth(), "\n";
$ws-&gt;getCol(2)-&gt;setHidden(true);              // hide column B
echo "col B hidden          = ", $boolstr($ws-&gt;getCol(2)-&gt;getHidden()), "\n";

$ws-&gt;getRange("C1:C1")-&gt;setValue("Sample string for autofit");
echo "col C bestFit default = ", $boolstr($ws-&gt;getCol(3)-&gt;getBestFit()), "\n";
$ws-&gt;getCol(3)-&gt;setBestFit(true);             // autofit the column width to the contents
echo "setBestFit(true) -&gt; bestFit = ", $boolstr($ws-&gt;getCol(3)-&gt;getBestFit()),
    " / columnWidth = ", $ws-&gt;getCol(3)-&gt;getColumnWidth(), "\n";

// Outlines (grouping) are set on the WorkSheet side
$ws-&gt;setRowOutline(4, 6, 1);                  // group rows 4-6 at level 1
$ws-&gt;setColOutline(2, 3, 1);                  // group columns B-C at level 1
echo "row 4 outlineLevel    = ", $ws-&gt;getRow(4)-&gt;getOutlineLevel(),
    " / col B outlineLevel = ", $ws-&gt;getCol(2)-&gt;getOutlineLevel(), "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// row-col-props.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("row-col-props-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- Row ----------
const row = ws.getRow(1);                        // row 1
console.log("default height        =", row.getHeight());
row.setHeight(30);
console.log("after setHeight(30)   =", row.getHeight());
console.log("getHidden default     =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true);                    // hide row 2
console.log("row 2 hidden          =", ws.getRow(2).getHidden());

// ---------- Col ----------
const col = ws.getCol(1);                        // column 1 (column A). Column numbers are passed as numbers
console.log("default col width     = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -&gt; columnWidth=" + col.getColumnWidth(),
            "/ width=" + col.getWidth());
ws.getCol(2).setHidden(true);                    // hide column B
console.log("col B hidden          =", ws.getCol(2).getHidden());

ws.getRange("C1:C1").setValue("Sample string for autofit");
console.log("col C bestFit default =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true);                   // autofit the column width to the contents
console.log("setBestFit(true) -&gt; bestFit =", ws.getCol(3).getBestFit(),
            "/ columnWidth =", ws.getCol(3).getColumnWidth());

// Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1);                       // group rows 4-6 at level 1
ws.setColOutline(2, 3, 1);                       // group columns B-C at level 1
console.log("row 4 outlineLevel    =", ws.getRow(4).getOutlineLevel(),
            "/ col B outlineLevel =", ws.getCol(2).getOutlineLevel());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">default height        = 18.75
after setHeight(30)   = 30
getHidden default     = false / getOutlineLevel = 0
row 2 hidden          = true
default col width     = columnWidth: 8.38 / width: 72
setColumnWidth(20) -&gt; columnWidth=20 / width=165
col B hidden          = true
col C bestFit default = false
setBestFit(true) -&gt; bestFit = true / columnWidth = 26.38
row 4 outlineLevel    = 1 / col B outlineLevel = 1</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7">The column width produced by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBestFit(true)</code> is computed from the cell contents and the book's default font, so the value varies by environment (above it is 26.38).</p>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Row and column numbers are all 1-based numeric values.</strong> Columns are passed as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1</code>, not <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A"</code>. To convert from a string, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertToColumnNumber("B")</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insert</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">delete</code> shift data; <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clear</code> does not.</strong> To just empty a row use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow()</code> (existing data keeps its position); to remove the row itself use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteRow()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copy</code> family takes the copy source as a "sheet name".</strong> Same-sheet copies also pass the sheet name.<br>Passing <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlCopyContentType</code> as the fifth argument <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyType</code> narrows what gets copied. | Value | Number | What gets copied | |---|---|---| | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeAll</code> | 0 | Values, cell formats, and formats of formula cells (default) | | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeValueOnly</code> | 1 | Values only | | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeValueAndFormat</code> | 2 | Values and cell formats | | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeFormatOnly</code> | 3 | Cell formats only |</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col.getWidth()</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col.getColumnWidth()</code> use different units.</strong> Measured: for a default column A, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWidth()</code> = 72 and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColumnWidth()</code> = 8.38. After <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnWidth(20)</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWidth()</code> = 165 and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColumnWidth()</code> = 20. <strong>Only <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnWidth()</code> (character width) can be set</strong>; <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWidth()</code> is read-only.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBestFit(true)</code> recomputes the column width from the cell contents at the time of the call.</strong> It does not follow later content changes automatically, so call it after all data is in place.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Outlines (grouping) are <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> methods.</strong> Set them with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setRowOutline(startRow, endRow, level)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColOutline(startCol, endCol, level)</code> and release them with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRowOutline(level)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearColOutline(level)</code>. Check the result with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row.getOutlineLevel()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col.getOutlineLevel()</code>.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Use the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell</code> according to the operation.</strong> For insertion the intended values are <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionRight</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionDown</code>(2); for deletion, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionLeft</code>(3) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionUp</code>(4).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Some things are not copied by the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copy</code> family.</strong> Even with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeAll</code>, <strong>shapes (including charts), pivot tables, queryTable-style tables and data tables are not copied.</strong></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The third argument of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell</code> is a single "top-left cell of the destination".</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell("Sheet1", "A3:B4", "F8")</code> expands to F8:G9 (the range size is determined by the source).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>For <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCol</code>, the source is a sheet name and positions are row/column numbers (numeric).</strong> Do not use A1-style range strings — they are easy to confuse, so be careful.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-04-value-types.php" class="text-primary-600 hover:underline">Cell value types and number formats</a> / Next: <a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">Bulk input of large data</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/en/nodejs/enums/enums.XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/en/nodejs/enums/enums.XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/python/pyosbxl.Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/en/python/pyosbxl.Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/RowWrapper.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/ColWrapper.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/php/classes/phposbxl-Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/en/php/classes/phposbxl-Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/wasm/classes/osbxl.Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/en/wasm/classes/osbxl.Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/en/wasm/enums/enums.XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/en/wasm/enums/enums.XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Rows, Columns and Ranges — insert, delete, copy",
   "description": "Rows, Columns and Ranges — insert, delete, copy. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-05-rows-cols.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-05-rows-cols.php",
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
     "name": "Rows, Columns and Ranges — insert, delete, copy",
     "item": "https://osb-office.com/en/sample/general-05-rows-cols.php"
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
