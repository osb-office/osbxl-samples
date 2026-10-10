<?php
$lang = 'en';
$page_title = 'Books and Worksheets — add, delete, move, rename | osbxl Samples';
$no_tagline = true;
$meta_description = 'Books and Worksheets — add, delete, move, rename. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,worksheet,addworksheet,deleteworksheet,movesheet,setname,openworksheetbyindex,getsheetcount,activate';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Books and Worksheets — add, delete, move, rename</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-03-sheets.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Add, delete, move and rename worksheets within a single book, walk sheets by index, and control which sheet is shown when the book opens (the active tab).</p>
<h2 id="methods-covered" class="text-2xl font-bold mt-10 mb-4">Methods covered</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Belongs to</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet(sheetName, position?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Adds a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Deletes a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet(sheetName, position)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Moves a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Opens a worksheet by name.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex(sheetIndex)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Opens a worksheet by index.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Gets the number of worksheets.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetNumber(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">Gets the sheet number of a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setName(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet</td><td class="border border-surface-200 px-3 py-2 align-top">Gets / sets the sheet name.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isActive()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(isActive)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet</td><td class="border border-surface-200 px-3 py-2 align-top">Gets / sets the active tab (the sheet shown when the book opens).</td></tr>
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
        <pre><code class="language-javascript">// sheets.js — add, move, rename and delete sheets, plus index access
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sheets.xlsx");

// Indexes are 1-based. Open 1..getSheetCount() in order and list the names
const show = (label) =&gt; {
    const names = [];
    for (let i = 1; i &lt;= wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheetByIndex(i).getName());
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
for (let i = 1; i &lt;= wb2.getSheetCount(); i++) names.push(wb2.openWorkSheetByIndex(i).getName());
console.log("[End] sheets:", names.join(", "));
console.log("     Orders!A1 =", wb2.openWorkSheet("Orders").getRange("A1:A1").getValue()["A1"]);
wb2.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># sheets.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sheets-py.xlsx")

def b(x):
    # render Python's True/False using node-style true/false notation
    return "true" if x else "false"

# Indexes are 1-based. Open 1..getSheetCount() in order and list the names
def show(label):
    names = [wb.openWorkSheetByIndex(i).getName() for i in range(1, wb.getSheetCount() + 1)]
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
names = [wb2.openWorkSheetByIndex(i).getName() for i in range(1, wb2.getSheetCount() + 1)]
print("[End] sheets:", ", ".join(names))
print("     Orders!A1 =", wb2.openWorkSheet("Orders").getRange("A1:A1").getValue()["A1"])
wb2.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// sheets.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    static WorkBookWrapper wb;

    // Indexes are 1-based. Open 1..getSheetCount() in order and list the names
    static void show(String label) {
        List&lt;String&gt; names = new ArrayList&lt;&gt;();
        for (int i = 1; i &lt;= wb.getSheetCount(); i++) {
            names.add(wb.openWorkSheetByIndex(i).getName());
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
        List&lt;String&gt; names = new ArrayList&lt;&gt;();
        for (int i = 1; i &lt;= wb2.getSheetCount(); i++) names.add(wb2.openWorkSheetByIndex(i).getName());
        System.out.println("[End] sheets: " + String.join(", ", names));
        Map&lt;String, String&gt; v = wb2.openWorkSheet("Orders").getRange("A1:A1").getValue();
        System.out.println("     Orders!A1 = " + v.get("A1"));
        wb2.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// sheets.php — same verification as the Node.js sample (phposbxl)
// Body only (&lt;?php and the autoload require are added by the harness)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("sheets-php.xlsx");

// Render PHP booleans using node-style true/false notation
function b($x) {
    return $x ? "true" : "false";
}

// Indexes are 1-based. Open 1..getSheetCount() in order and list the names
function show($label) {
    global $wb;
    $names = [];
    for ($i = 1; $i &lt;= $wb-&gt;getSheetCount(); $i++) {
        $names[] = $wb-&gt;openWorkSheetByIndex($i)-&gt;getName();
    }
    echo $label . " count=" . $wb-&gt;getSheetCount() . " [" . implode(", ", $names) . "]\n";
}

show("[1] after creation:");

$wb-&gt;addWorkSheet("Summary");                     // position omitted → appended at the end
show("[2] addWorkSheet(Summary):");

$wb-&gt;addWorkSheet("Orders", 1);                   // position=1 → inserted at the front
show("[3] addWorkSheet(Orders,1):");

$wb-&gt;moveSheet("Orders", 3);                      // move to position 3
show("[4] moveSheet(Orders,3):");

$ws = $wb-&gt;openWorkSheet("Orders");
echo "[5] getName() = " . $ws-&gt;getName() . " / getSheetNumber(Orders) = " . $wb-&gt;getSheetNumber("Orders") . "\n";
$ws-&gt;getRange("A1:A1")-&gt;setValue("order data");

$wb-&gt;openWorkSheet("Sheet1")-&gt;setName("Details"); // rename
show("[6] setName(Details):");

// --- active tab (the sheet shown in front when the book is opened) ---
echo "[7] isActive: Summary = " . b($wb-&gt;openWorkSheet("Summary")-&gt;isActive())
   . " / Details = " . b($wb-&gt;openWorkSheet("Details")-&gt;isActive()) . "\n";
$wb-&gt;openWorkSheet("Summary")-&gt;activate(true);    // the other sheets are deactivated automatically
echo "    after activate: Summary = " . b($wb-&gt;openWorkSheet("Summary")-&gt;isActive())
   . " / Details = " . b($wb-&gt;openWorkSheet("Details")-&gt;isActive()) . "\n";

// --- delete ---
$wb-&gt;addWorkSheet("Temp");
$wb-&gt;deleteWorkSheet("Temp");
show("[8] deleteWorkSheet(Temp):");

$wb-&gt;save();
$wb-&gt;close();

// --- reopen and verify ---
$wb2 = $app-&gt;openWorkBook("sheets-php.xlsx");
$names = [];
for ($i = 1; $i &lt;= $wb2-&gt;getSheetCount(); $i++) {
    $names[] = $wb2-&gt;openWorkSheetByIndex($i)-&gt;getName();
}
echo "[End] sheets: " . implode(", ", $names) . "\n";
$v = $wb2-&gt;openWorkSheet("Orders")-&gt;getRange("A1:A1")-&gt;getValue();
echo "     Orders!A1 = " . $v["A1"] . "\n";
$wb2-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// sheets.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("sheets-wasm.xlsx");

// Indexes are 1-based. Open 1..getSheetCount() in order and list the names
const show = (label) =&gt; {
    const names = [];
    for (let i = 1; i &lt;= wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheetByIndex(i).getName());
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
for (let i = 1; i &lt;= wb2.getSheetCount(); i++) names.push(wb2.openWorkSheetByIndex(i).getName());
console.log("[End] sheets:", names.join(", "));
console.log("     Orders!A1 =", wb2.openWorkSheet("Orders").getRange("A1:A1").getValue()["A1"]);
wb2.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] after creation: count=1 [Sheet1]
[2] addWorkSheet(Summary): count=2 [Sheet1, Summary]
[3] addWorkSheet(Orders,1): count=3 [Orders, Sheet1, Summary]
[4] moveSheet(Orders,3): count=3 [Sheet1, Summary, Orders]
[5] getName() = Orders / getSheetNumber(Orders) = 3
[6] setName(Details): count=3 [Details, Summary, Orders]
[7] isActive: Summary = false / Details = true
    after activate: Summary = true / Details = false
[8] deleteWorkSheet(Temp): count=3 [Details, Summary, Orders]
[End] sheets: Details, Summary, Orders
     Orders!A1 = order data</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Both <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> and sheet numbers are 1-based.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet(name, position)</code> <strong>inserts before <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code></strong> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> = 1 means the front). Omitting it (default -1) appends at the end.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet(name, position)</code> is the "position after the move"</strong>, with the front being 1.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex(i)</code> is also 1-based.</strong> Combined with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code> it lets you enumerate all sheets by name (the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">show()</code> above). When handling a book whose sheet names you don't know, walking it this way is the reliable approach.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetNumber(name)</code> returns the number in the current ordering.</strong> The value changes after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet()</code> or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteWorkSheet()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setName()</code> takes effect immediately.</strong> From then on, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet()</code> with the new name.<br>Trying to open it by the old name raises an exception.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(true)</code> works exclusively.</strong> When you <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(true)</code> the target sheet, the other sheets are <strong>deactivated automatically</strong> (measured). You do not need to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(false)</code> the others beforehand.<br>The active state is kept across <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> and reopening.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Passing a nonexistent sheet name to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet()</code> raises an exception.</strong><div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">wb.openWorkSheet("NoSuch");   // → exception: "NoSuch is not found"</code></pre></div>When sheet names come from user input or external configuration, first enumerate the names with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code> + <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex()</code> and match against them.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>An out-of-range index for <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex()</code> also raises an exception.</strong> Call it with a value in the range <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1</code> to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Right after creating a new book, only the default sheet (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet1</code>) has <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isActive() === true</code></strong> (measured). Sheets added with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet()</code> are <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code>. If you want a different sheet shown in front when the book opens, call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(true)</code> on that sheet.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCol</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell</code> to copy cells across sheets</strong> (the copy source is specified by sheet name). See <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">Working with rows, columns and cell ranges</a> for details.<br>Shapes (including charts), pivot tables, queryTable-style tables and data tables are not included in what gets copied.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">Building a table with values and formulas</a> / Next: <a href="general-04-value-types.php" class="text-primary-600 hover:underline">Cell value types and number formats</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/WorkBookWrapper.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Books and Worksheets — add, delete, move, rename",
   "description": "Books and Worksheets — add, delete, move, rename. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-03-sheets.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-03-sheets.php",
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
     "name": "Books and Worksheets — add, delete, move, rename",
     "item": "https://osb-office.com/en/sample/general-03-sheets.php"
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
