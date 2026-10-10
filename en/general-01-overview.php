<?php
$lang = 'en';
$page_title = 'Overview — Object Model and Basic Flow | osbxl Samples';
$no_tagline = true;
$meta_description = 'Overview — Object Model and Basic Flow. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,overview,object-model,lifecycle,app,workbook,worksheet,range,create,open,save,saveas,close';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Overview — Object Model and Basic Flow</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-01-overview.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Get the big picture of nodeosbxl.</p>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li>The object hierarchy (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code>)</li>
  <li>The <strong>basic flow</strong> (create new / open existing book → sheet → Range → values &amp; formulas → save / saveAs / close)</li>
  <li>The main methods of the four classes</li>
</ul>
<p class="mb-4 text-slate-700 leading-7">Details of individual features are covered by the samples that follow this page.</p>
<h2 id="prerequisites" class="text-2xl font-bold mt-10 mb-4">Prerequisites</h2>
<p class="mb-4 text-slate-700 leading-7">Download the module for your language from the <a href="/pages/download.php" class="text-primary-600 hover:underline">download page</a> and install it.</p>
<p class="mb-4 text-slate-700 leading-7">For detailed per-OS PHP installation steps (Linux / macOS / Windows, including when Composer is not installed), see the <a href="/en/pages/quickstart.php" class="text-primary-600 hover:underline">Quickstart</a>.</p>
<p class="mb-4 text-slate-700 leading-7">The WebAssembly build uses MEMORY64 (Wasm 3.0), so the runtime must support it (Chrome/Edge 133+, Firefox 134+, Node.js 24+. Safari is not supported yet).</p>
<p class="mb-4 text-slate-700 leading-7">The entry point differs per language; the namespace layout (operation classes / dto / chart / enums) is common to all four.</p>
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
        <pre><code class="language-bash">npm install ./nodeosbxl-1.5.0.tgz</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># (example) Python 3.12+ (a single abi3 wheel covers 3.12 and later)
pip install ./pyosbxl-1.5.0-cp312-abi3-linux_x86_64.whl</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># Add the JAR to your classpath (javacpp-1.5.14.jar is also required)
javac -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar Main.java
java  -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar:. Main</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># Download the tarball or Composer zip for your PHP version (8.2-8.5) and OS
# from osb-office.com/files/php/
# e.g. phposbxl-1.5.0-php8.5-nts-linux-x86_64.tar.gz (Linux, PHP 8.5)
# Extract it and install via composer with the extracted dir as a path repository
composer config repositories.phposbxl path /path/to/phposbxl-1.5.0-php8.5-nts-linux-x86_64
composer require osb-office/phposbxl
# For the Composer zip: use the dir containing the zip as an artifact repository
# composer config repositories.phposbxl artifact /path/to/zip-dir &amp;&amp; composer require osb-office/phposbxl</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># Unpack the package (contains osbxl_wasm.js / osbxl_em.js / license/)
tar xzf wasmosbxl-1.5.0.tgz</code></pre>
        <pre><code class="language-javascript">// Loading example (requires Node.js 24+: this wasm is a MEMORY64 = Wasm 3.0 build)
const fs = require("fs");
const createModule = require("./package/osbxl_wasm.js");   // emscripten factory
const { createOsbxl } = require("./package/osbxl_em.js");  // loader + enum definitions

createOsbxl(createModule).then(({ Module, osbxl, dto, chart, enums }) =&gt; {
  // Stage the license file into the wasm virtual FS (required before new App())
  const lic = process.env.OSB_LICENSE_PATH_WASM;          // e.g. /root/lic_node.osb
  if (lic &amp;&amp; fs.existsSync(lic)) {
    Module.FS.mkdirTree(lic.substring(0, lic.lastIndexOf("/")));
    Module.FS.writeFile(lic, fs.readFileSync(lic));
  }
  const app = new osbxl.App();   // from here on it matches the Node.js build (namespace name is osbxl)
  // ...
});
// In the browser, load osbxl_wasm.js via &lt;script&gt; and call createOsbxl the same way.
// Supported: Chrome/Edge 133+, Firefox 134+ (Safari not supported yet)</code></pre>
      </div>
    </div>
</div>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Language</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Entry point</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Main namespaces</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Node.js</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">require("nodeosbxl")</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Python</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">import pyosbxl</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl.dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl.chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl.enums</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Java</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">import com.osboffice.osbxl.*</code></td><td class="border border-surface-200 px-3 py-2 align-top">packages such as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">WebAssembly</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createOsbxl(createModule)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">osbxl</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">The table below uses Node.js names; the counts per namespace are common to all four languages.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Namespace</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Contents</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Count</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl</code></td><td class="border border-surface-200 px-3 py-2 align-top">Classes you operate on (App / WorkBook / WorkSheet / Range / Font / Chart entry points, etc.)</td><td class="border border-surface-200 px-3 py-2 align-top">29 classes</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code></td><td class="border border-surface-200 px-3 py-2 align-top">Objects used to pass data around (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InputValueObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FontObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ColorObject</code>, etc.)</td><td class="border border-surface-200 px-3 py-2 align-top">40 classes</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code></td><td class="border border-surface-200 px-3 py-2 align-top">Chart-specific objects (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Axis</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartTitle</code>, etc.)</td><td class="border border-surface-200 px-3 py-2 align-top">43 classes</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td><td class="border border-surface-200 px-3 py-2 align-top">Excel-compatible enumerations (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlFont</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlChartType</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlIndexColor</code>, etc.)</td><td class="border border-surface-200 px-3 py-2 align-top">127 kinds</td></tr>
  </tbody></table></div>
<h2 id="object-model" class="text-2xl font-bold mt-10 mb-4">Object model</h2>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">App                     ← the single entry point. Open/create books, convert cell addresses
 │
 ├─ createWorkBook()  ─┐
 ├─ openWorkBook()    ─┴→ WorkBook          ← one xlsx file
 │                        │
 │                        ├─ openWorkSheet()      → WorkSheet   ← one sheet
 │                        │                         │
 │                        │                         ├─ getRange("A1:B2")    → Range
 │                        │                         │    ※ getCells(1,1,2,2) returns the same Range
 │                        │                         │                       │
 │                        │                         │                       ├─ getValue() / setValue() / setFormula()
 │                        │                         │                       ├─ getFont()   → Font   → getColor() → Color
 │                        │                         │                       ├─ getBorders()→ Borders→ getBorder()  → Border
 │                        │                         │                       ├─ getFill()   → Fill
 │                        │                         │                       └─ getAlignment() → Alignment
 │                        │                         │
 │                        │                         ├─ getRow() / getCol()  → Row / Col
 │                        │                         ├─ getSort() / getAutoFilter()
 │                        │                         ├─ getPivotTables() / getListObjects()
 │                        │                         ├─ addChartObject()     → chart.ChartObject
 │                        │                         └─ getComments() / getHyperLinks() / getShapes() ...
 │                        │
 │                        ├─ addWorkSheet() / deleteWorkSheet() / moveSheet()
 │                        ├─ addCustomStyle() / addCustomTableStyle()
 │                        └─ save() / saveAs() / close()
 │
 └─ utilities such as convertFromRowColNumber()</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><strong>Always walking this hierarchy top-down</strong> is the osbxl basics. Below <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Font</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Color</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Borders</code>, …) the pattern is "get the object, then call its setters".</p>
<h2 id="basic-flow" class="text-2xl font-bold mt-10 mb-4">Basic flow</h2>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext"> [1] new App()                 create the App
 [2] createWorkBook(path)      create a new book
     openWorkBook(path)        open an existing book
 [3] openWorkSheet(name)       open a worksheet
 [4] getRange("A1:B2")         get a Range (getCells(1,1,2,2) is the same)
 [5] setValue() / setFormula() write values / formulas
 [6] getValue(true)            read the results back
 [7] save() / saveAs(path)     save
 [8] close()                   close the book</code></pre>
</div>
<h3 id="code" class="text-xl font-bold mt-8 mb-3">Code</h3>
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
        <pre><code class="language-javascript">// basic-flow.js — the full flow: create → edit → save → reopen → save as
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();                  // [1] create the App

// ---------- create a new book, edit it and save ----------
let wb = app.createWorkBook("flow.xlsx");         // [2] create new
console.log("[2] createWorkBook -&gt; getBookPath =", wb.getBookPath());

let ws = wb.openWorkSheet("Sheet1");              // [3] open a worksheet
console.log("[3] openWorkSheet  -&gt; sheet name =", ws.getName());

// [4] get a Range and [5] set values
//     ws.getCells(1, 1) (numeric row/col) returns the same Range as getRange("A1:A1")
ws.getRange("A1:A1").setValue("created new");
ws.getRange("B1:B1").setFormula("100*3");         // [5] set a formula (no leading "=")
console.log("[6] getValue(true) -&gt; B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);

wb.save();                                        // [7] save
console.log("[7] save() done");
wb.close();                                       // [8] close (no close needed for WorkSheet)
console.log("[8] close() done");

// ---------- open the existing book and save it under a new name ----------
wb = app.openWorkBook("flow.xlsx");               // [2'] open an existing book
ws = wb.openWorkSheet("Sheet1");                  // [3']
const read = ws.getRange("A1:B1").getValue(true);
console.log("[2'] openWorkBook -&gt; A1 =", read["A1"], "/ B1 =", read["B1"]);

ws.getRange("A2:A2").setValue("appended after reopen");
wb.saveAs("flow-copy.xlsx");                      // [7'] save under a new name
console.log("[7'] saveAs() done -&gt; getBookPath =", wb.getBookPath());
wb.close();                                       // [8']
console.log("[8] close() done");</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># basic-flow.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

app = pyosbxl.App()                               # [1] create the App

wb = app.createWorkBook("flow-py.xlsx")           # [2] create new
ws = wb.openWorkSheet("Sheet1")                   # [3] open a worksheet

ws.getRange("A1:A1").setValue("created new")
ws.getRange("B1:B1").setFormula("100*3")          # [5] formula (no leading "=")
print("[6] getValue(True) -&gt; B1 =", ws.getRange("B1:B1").getValue(True)["B1"])

wb.save()                                         # [7] save
wb.close()                                        # [8] close
print("[8] close() done")

wb = app.openWorkBook("flow-py.xlsx")             # [2'] open an existing book
ws = wb.openWorkSheet("Sheet1")
print("[2'] openWorkBook -&gt; A1 =", ws.getRange("A1:A1").getValue(True)["A1"])
wb.close()
print("[8] close() done")</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// basic-flow.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();                          // [1] create the App

        WorkBookWrapper wb = app.createWorkBook("flow-java.xlsx");  // [2] create new
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");           // [3] open a worksheet

        ws.getRange("A1:A1").setValue("created new");
        ws.getRange("B1:B1").setFormula("100*3");                   // [5] formula (no leading "=")
        Map&lt;String, String&gt; v = ws.getRange("B1:B1").getValue(true);
        System.out.println("[6] getValue(true) -&gt; B1 = " + v.get("B1"));

        wb.save();                                                  // [7] save
        wb.close();                                                 // [8] close
        System.out.println("[8] close() done");

        wb = app.openWorkBook("flow-java.xlsx");                    // [2'] open an existing book
        ws = wb.openWorkSheet("Sheet1");
        Map&lt;String, String&gt; r = ws.getRange("A1:A1").getValue(true);
        System.out.println("[2'] openWorkBook -&gt; A1 = " + r.get("A1"));
        wb.close();
        System.out.println("[8] close() done");
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// basic-flow.php — same verification as the Node.js sample (phposbxl)
// Body only (&lt;?php and the autoload require are added by the harness)
$app = new \phposbxl\App();                        // [1] create the App

// ---------- create a new book, edit it and save ----------
$wb = $app-&gt;createWorkBook("flow-php.xlsx");       // [2] create new
echo "[2] createWorkBook -&gt; getBookPath = " . $wb-&gt;getBookPath() . "\n";

$ws = $wb-&gt;openWorkSheet("Sheet1");                // [3] open a worksheet
echo "[3] openWorkSheet  -&gt; sheet name = " . $ws-&gt;getName() . "\n";

// [4] get a Range and [5] set values
//     $ws-&gt;getCells(1, 1) (numeric row/col) returns the same Range as getRange("A1:A1")
$ws-&gt;getRange("A1:A1")-&gt;setValue("created new");
$ws-&gt;getRange("B1:B1")-&gt;setFormula("100*3");       // [5] set a formula (no leading "=")
$v = $ws-&gt;getRange("B1:B1")-&gt;getValue(true);
echo "[6] getValue(true) -&gt; B1 = " . $v["B1"] . "\n";

$wb-&gt;save();                                       // [7] save
echo "[7] save() done\n";
$wb-&gt;close();                                      // [8] close (no close needed for WorkSheet)
echo "[8] close() done\n";

// ---------- open the existing book and save it under a new name ----------
$wb = $app-&gt;openWorkBook("flow-php.xlsx");         // [2'] open an existing book
$ws = $wb-&gt;openWorkSheet("Sheet1");                // [3']
$read = $ws-&gt;getRange("A1:B1")-&gt;getValue(true);
echo "[2'] openWorkBook -&gt; A1 = " . $read["A1"] . " / B1 = " . $read["B1"] . "\n";

$ws-&gt;getRange("A2:A2")-&gt;setValue("appended after reopen");
$wb-&gt;saveAs("flow-copy-php.xlsx");                 // [7'] save under a new name
echo "[7'] saveAs() done -&gt; getBookPath = " . $wb-&gt;getBookPath() . "\n";
$wb-&gt;close();                                      // [8']
echo "[8] close() done\n";</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// basic-flow.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();                        // [1] create the App

const wb = app.createWorkBook("flow-wasm.xlsx");    // [2] create new
const ws = wb.openWorkSheet("Sheet1");              // [3] open a worksheet

ws.getRange("A1:A1").setValue("created new");
ws.getRange("B1:B1").setFormula("100*3");           // [5] formula (no leading "=")
console.log("[6] getValue(true) -&gt; B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);

wb.save();                                          // [7] save
wb.close();                                         // [8] close
console.log("[8] close() done");

const wb2 = app.openWorkBook("flow-wasm.xlsx");     // [2'] open an existing book
const ws2 = wb2.openWorkSheet("Sheet1");
console.log("[2'] openWorkBook -&gt; A1 =", ws2.getRange("A1:A1").getValue(true)["A1"]);
wb2.close();
console.log("[8] close() done");</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[2] createWorkBook -&gt; getBookPath = &lt;working directory&gt;/flow.xlsx
[3] openWorkSheet  -&gt; sheet name = Sheet1
[6] getValue(true) -&gt; B1 = 300
[7] save() done
[8] close() done
[2'] openWorkBook -&gt; A1 = created new / B1 = 300
[7'] saveAs() done -&gt; getBookPath = &lt;working directory&gt;/flow.xlsx
[8] close() done</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7">Two files are produced: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">flow.xlsx</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">flow-copy.xlsx</code>.</p>
<h3 id="flow-highlights" class="text-xl font-bold mt-8 mb-3">Flow highlights</h3>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>One <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> per process is enough.</strong> Even when handling multiple books, call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkBook()</code> on the same <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app</code>.</li>
  <li><strong>Relative paths are resolved to absolute paths.</strong> After <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook("flow.xlsx")</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBookPath()</code> returns an absolute path including the current directory.</li>
  <li><strong>There is no operation to close a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code>.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">close()</code> exists only as a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> method. Sheets are all released together by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.close()</code>.</li>
  <li><strong>Note the difference between <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs()</code></strong> (see "Caveats" below).</li>
</ul>
<h2 id="main-methods" class="text-2xl font-bold mt-10 mb-4">Main methods</h2>
<h3 id="app-10-methods" class="text-xl font-bold mt-8 mb-3">App (10 methods)</h3>
<p class="mb-4 text-slate-700 leading-7">Created with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new nodeosbxl.App()</code>. Handles opening/closing books and address conversion utilities.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getVersion()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns the version number.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook(path, defaultFont?, defaultFontSize?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Creates a workbook.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkBook(path)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Opens a workbook.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openPasswordWorkBook(path, password)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Opens a password-protected workbook.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber(row, col)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Converts numeric cell coordinates to A1C1 notation. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(1,1)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1"</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber2(startRow, startCol, endRow, endCol)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Converts a numeric cell range to A1C1 notation. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(1,1,2,3)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:C2"</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertToColumnNumber("B")</code></td><td class="border border-surface-200 px-3 py-2 align-top">Converts a column letter to its numeric index. → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromColumnNumber(3)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Converts a numeric column index to its letter. → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"C"</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumericValue(dateTimeObject, is1904?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the Excel internal serial value from a date-time object.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">calculateHexColor(bookPath, colorObj, isForeGroundColor)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the HEX value taking workbook-specific color information into account.</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber2</code> also have overloads that return absolute references (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"$A$1"</code>).</p>
<h3 id="workbook-27-methods-main-12" class="text-xl font-bold mt-8 mb-3">WorkBook (27 methods) — main 12</h3>
<p class="mb-4 text-slate-700 leading-7">Corresponds to one xlsx file.</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBookPath()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Returns the workbook's file path.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Opens a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex(sheetIndex)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Opens a worksheet by index.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet(sheetName, position?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Adds a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Deletes a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet(sheetName, position)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Moves a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the number of worksheets.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetNumber(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the sheet number of a worksheet.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save(password)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Saves the workbook to the same file.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs(path)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs(path, password)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Saves the workbook to a different file.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">close()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Closes the workbook.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isDate1904()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets whether the workbook uses the Excel 1904 date system.</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Others (15): styles — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNames</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomTableStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotCustomTableStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCustomTableStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCustomTableStyle</code> / external books — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addExternalWorkBook</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">updateExternalWorkBook</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getExternalWorkBookPath</code> / file properties — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCompanyName</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setManagerName</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCreateAuthor</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setLastAuthor</code></p>
<h3 id="worksheet-44-methods-main-14" class="text-xl font-bold mt-8 mb-3">WorkSheet (44 methods) — main 14</h3>
<p class="mb-4 text-slate-700 leading-7">Corresponds to one sheet. <strong>This is the gateway to the objects below it (Range / Sort / PivotTable / Chart …).</strong></p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange("A1:B2")</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a cell range class instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells(row, col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells(sr, sc, er, ec)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a cell range class instance from numeric coordinates.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRow(rowNum)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a row class instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCol(colNum)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets a column class instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setName(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets / sets the sheet name.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isActive()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets / sets whether the sheet tab is the one shown when the book opens.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertRow(startRow, numOfRows)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Inserts rows.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteRow(startRow, numOfRows)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Deletes rows.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCol(startCol, numOfCols)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Inserts columns.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCol(startCol, numOfCols)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Deletes columns.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray(values)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Bulk-sets cell values.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray(formulas)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Bulk-sets formulas.</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Others (30): feature objects — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormatConditions</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAutoFilter</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSort</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComments</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHyperLinks</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getListObjects</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTables</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getShapes</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWindow</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHPageBreaks</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getVPageBreaks</code> / charts — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteChartObject</code> / print setup — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPageSetupObject</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPageSetupObject</code> / selection — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getActiveCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setActiveCell</code> / row, column and cell operations — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCol</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCol</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell</code> / outlining — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setRowOutline</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRowOutline</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColOutline</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearColOutline</code></p>
<h3 id="range-31-methods-main-16" class="text-xl font-bold mt-8 mb-3">Range (31 methods) — main 16</h3>
<p class="mb-4 text-slate-700 leading-7">Corresponds to a cell range. <strong>Reads/writes values and formulas, and is the gateway to formatting objects.</strong></p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(rawValue?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets cell values. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> returns raw values (formula results).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the formula.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(value, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a cell value (general-purpose method).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(value, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a cell value (numeric setter).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(dateTimeObject, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a cell value (date/time setter).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue(str, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a cell value (date/time from a string).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(bool, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a cell value (boolean setter).</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula(formula, isArray?, setAllCell?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets a formula.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFont()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the font object instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBorders()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the borders collection object instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFill()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the fill object instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAlignment()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the cell alignment object instance.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberFormat(format)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Sets the number format.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">merge(mergeEachRow?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Merges the cell range.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearContent()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Clears the cells.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAddress()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets the selected range of the Range.</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Others (15): <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getProtection</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCharacters</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPhonetics</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumberFormat</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearNumberFormat</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearAllFormat</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataTable</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearDataTable</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBuiltinStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalWidth</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalColumnWidth</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowHeight</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">unMerge</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">replaceValue</code></p>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> is "get it and throw it away".</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("A1:A1")</code> returns a new instance on every call, so there is no need to keep holding it in a variable.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> returns a map.</strong> Shaped like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ "A1": "100", "B1": "300" }</code>, and all values are <strong>strings</strong>. Convert with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Number(...)</code> to treat them as numbers.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Two ways to get a Range.</strong><ul class="mt-2 space-y-1 pl-5 list-disc text-slate-700"><li><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("A1:B2")</code> — by <strong>string address</strong>. Even a single cell is passed in range form as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:A1"</code>.</li><li><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getCells(row, col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getCells(startRow, startCol, endRow, endCol)</code> — by <strong>numbers (row, col)</strong>. Rows and columns are <strong>1-based</strong>: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(1,1)</code> = A1, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(3,1,3,3)</code> = A3:C3. Handy when looping over row/column numbers because you don't have to assemble address strings.</li><li>For number ⇔ string conversion use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber(row, col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber2(sr, sc, er, ec)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertToColumnNumber("B")</code>.</li></ul></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Formulas are computed at run time.</strong> Right after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula()</code> you can read the result with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> (no need to wait for <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code>).</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>A workbook is locked internally to guarantee uniqueness, so simultaneous read/write of the same workbook never occurs.</strong></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">close()</code> exists only as a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> method.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> have no close-style methods.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs()</code> only "writes the current content to a different file"; it does not switch the destination of later saves.</strong> This differs from Excel VBA's <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SaveAs</code>.<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-plaintext">wb.save()                  → written to base.xlsx
wb.saveAs("copy.xlsx")     → "the content at this point" is written to copy.xlsx
wb.getBookPath()           → still base.xlsx (not switched)
ws.getRange(...).setValue(...)
wb.save()                  → written to base.xlsx (not copy.xlsx)</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Do not prefix formulas with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code>.</strong> Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>A license banner line is printed to stdout at startup</strong> (once per process). Keep it in mind when writing code that parses standard output.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Next sample: <a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">Building a table with values and formulas</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/nodejs/classes/nodeosbxl.Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/en/python/pyosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/python/pyosbxl.Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/AppWrapper.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/WorkBookWrapper.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/RangeWrapper.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/en/php/classes/phposbxl-WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/php/classes/phposbxl-Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/en/wasm/classes/osbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/en/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/en/wasm/classes/osbxl.Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Overview — Object Model and Basic Flow",
   "description": "Overview — Object Model and Basic Flow. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-01-overview.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-01-overview.php",
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
     "name": "Overview — Object Model and Basic Flow",
     "item": "https://osb-office.com/en/sample/general-01-overview.php"
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
