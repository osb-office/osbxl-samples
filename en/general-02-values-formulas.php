<?php
$lang = 'en';
$page_title = 'Building a Table with Values and Formulas | osbxl Samples';
$no_tagline = true;
$meta_description = 'Building a Table with Values and Formulas. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,value,formula,sum,table,setformulaarray,inputformulaobject,convertfromrowcolnumber';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Building a Table with Values and Formulas</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-02-values-formulas.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Create a new xlsx, write a table (headers, strings, numbers, formulas) and save the file. This applies the basic flow from the <a href="general-01-overview.php" class="text-primary-600 hover:underline">overview</a> to real table building.</p>
<h2 id="resulting-image" class="text-2xl font-bold mt-10 mb-4">Resulting image</h2>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet1</code> of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">sales.xlsx</code> will contain this table:</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap"></th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">A</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">B</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">C</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">D</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>1</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Product</td><td class="border border-surface-200 px-3 py-2 align-top">Unit Price</td><td class="border border-surface-200 px-3 py-2 align-top">Qty</td><td class="border border-surface-200 px-3 py-2 align-top">Amount</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>2</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Apple</td><td class="border border-surface-200 px-3 py-2 align-top">100</td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">300</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>3</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Orange</td><td class="border border-surface-200 px-3 py-2 align-top">50</td><td class="border border-surface-200 px-3 py-2 align-top">6</td><td class="border border-surface-200 px-3 py-2 align-top">300</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>4</strong></td><td class="border border-surface-200 px-3 py-2 align-top">Grape</td><td class="border border-surface-200 px-3 py-2 align-top">150</td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">300</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>5</strong></td><td class="border border-surface-200 px-3 py-2 align-top"></td><td class="border border-surface-200 px-3 py-2 align-top"></td><td class="border border-surface-200 px-3 py-2 align-top">Total</td><td class="border border-surface-200 px-3 py-2 align-top">900</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Column D holds the formula <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">B*C</code>; D5 holds <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUM(D2:D4)</code>.</p>
<h2 id="prerequisites" class="text-2xl font-bold mt-10 mb-4">Prerequisites</h2>
<p class="mb-4 text-slate-700 leading-7">Download the module for your language from the <a href="/pages/download.php" class="text-primary-600 hover:underline">download page</a> and install it (entry points are listed in the <a href="general-01-overview.php#prerequisites" class="text-primary-600 hover:underline">overview Prerequisites</a> table).</p>
<p class="mb-4 text-slate-700 leading-7">For detailed per-OS PHP installation steps (Linux / macOS / Windows, including when Composer is not installed), see the <a href="/en/pages/quickstart.php" class="text-primary-600 hover:underline">Quickstart</a>.</p>
<p class="mb-4 text-slate-700 leading-7">The WebAssembly build uses MEMORY64 (Wasm 3.0), so the runtime must support it (Chrome/Edge 133+, Firefox 134+, Node.js 24+. Safari is not supported yet).</p>
<p class="mb-4 text-slate-700 leading-7">If you omit the directory in the path (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"sales.xlsx"</code>), the file is written to the <strong>process current working directory</strong>.</p>
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
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// values-formulas.js — create a book, write values and formulas, save
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sales.xlsx");   // new workbook
const ws = wb.openWorkSheet("Sheet1");         // open the default sheet

// --- header (row 1) ---
// app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -&gt; "A1"
["Product", "Unit Price", "Qty", "Amount"].forEach((text, i) =&gt; {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- detail rows (2-4) ---
const data = [["Apple", 100, 3], ["Orange", 50, 6], ["Grape", 150, 2]];
data.forEach(([name, price, qty], i) =&gt; {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- amount column (D2:D4) ---
// Each row needs its own relative reference, so use setFormulaArray.
// Formula strings do NOT take a leading "=".
const formulas = data.map((_, i) =&gt; {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(i + 2, 4));  // D2, D3, D4
    f.setFormula(`B${i + 2}*C${i + 2}`);               // "B2*C2" ...
    return f;
});
ws.setFormulaArray(formulas);

// --- total (D5) ---
ws.getRange("C5:C5").setValue("Total");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- read back ---
// getValue(true) returns raw values = formula results
const values = ws.getRange("A1:D5").getValue(true);
console.log("Amount D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("Total:", values["D5"]);

// --- save and close ---
wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># values-formulas.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sales-py.xlsx")   # new workbook
ws = wb.openWorkSheet("Sheet1")            # open the default sheet

# --- header (row 1) ---
# app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -&gt; "A1"
for i, text in enumerate(["Product", "Unit Price", "Qty", "Amount"]):
    cell = app.convertFromRowColNumber(1, i + 1)
    ws.getRange(f"{cell}:{cell}").setValue(text)

# --- detail rows (2-4) ---
data = [("Apple", 100, 3), ("Orange", 50, 6), ("Grape", 150, 2)]
for i, (name, price, qty) in enumerate(data):
    row = i + 2
    ws.getRange(f"A{row}:A{row}").setValue(name)
    ws.getRange(f"B{row}:B{row}").setNumberValue(price)
    ws.getRange(f"C{row}:C{row}").setNumberValue(qty)

# --- amount column (D2:D4) ---
# Each row needs its own relative reference, so use setFormulaArray.
# Formula strings do NOT take a leading "=".
formulas = []
for i in range(len(data)):
    f = pyosbxl.dto.InputFormulaObject()
    f.setA1C1(app.convertFromRowColNumber(i + 2, 4))  # D2, D3, D4
    f.setFormula(f"B{i + 2}*C{i + 2}")                # "B2*C2" ...
    formulas.append(f)
ws.setFormulaArray(formulas)

# --- total (D5) ---
ws.getRange("C5:C5").setValue("Total")
ws.getRange("D5:D5").setFormula("SUM(D2:D4)")

# --- read back ---
# getValue(True) returns raw values = formula results
values = ws.getRange("A1:D5").getValue(True)
print("Amount D2:D4 =", ", ".join([values["D2"], values["D3"], values["D4"]]))
print("Total:", values["D5"])

# --- save and close ---
wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// values-formulas.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("sales-java.xlsx");   // new workbook
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");             // open the default sheet

        // --- header (row 1) ---
        // app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -&gt; "A1"
        String[] headers = { "Product", "Unit Price", "Qty", "Amount" };
        for (int i = 0; i &lt; headers.length; i++) {
            String cell = app.convertFromRowColNumber(1, i + 1);
            ws.getRange(cell + ":" + cell).setValue(headers[i]);
        }

        // --- detail rows (2-4) ---
        String[] names = { "Apple", "Orange", "Grape" };
        double[] prices = { 100, 50, 150 };
        double[] qtys = { 3, 6, 2 };
        for (int i = 0; i &lt; names.length; i++) {
            int row = i + 2;
            ws.getRange("A" + row + ":A" + row).setValue(names[i]);
            ws.getRange("B" + row + ":B" + row).setNumberValue(prices[i]);
            ws.getRange("C" + row + ":C" + row).setNumberValue(qtys[i]);
        }

        // --- amount column (D2:D4) ---
        // Each row needs its own relative reference, so use setFormulaArray.
        // Formula strings do NOT take a leading "=".
        List&lt;InputFormulaObjectWrapper&gt; formulas = new ArrayList&lt;&gt;();
        for (int i = 0; i &lt; names.length; i++) {
            InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
            f.setA1C1(app.convertFromRowColNumber(i + 2, 4));   // D2, D3, D4
            f.setFormula("B" + (i + 2) + "*C" + (i + 2));       // "B2*C2" ...
            formulas.add(f);
        }
        ws.setFormulaArray(formulas);

        // --- total (D5) ---
        ws.getRange("C5:C5").setValue("Total");
        ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

        // --- read back ---
        // getValue(true) returns raw values = formula results
        Map&lt;String, String&gt; values = ws.getRange("A1:D5").getValue(true);
        System.out.println("Amount D2:D4 = "
                + String.join(", ", values.get("D2"), values.get("D3"), values.get("D4")));
        System.out.println("Total: " + values.get("D5"));

        // --- save and close ---
        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// values-formulas.php — same verification as the Node.js sample (phposbxl)
// Body only (&lt;?php and the autoload require are added by the harness)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("sales-php.xlsx");   // new workbook
$ws = $wb-&gt;openWorkSheet("Sheet1");             // open the default sheet

// --- header (row 1) ---
// $app-&gt;convertFromRowColNumber(row, col) gives the A1 address for (1,1) -&gt; "A1"
$headers = ["Product", "Unit Price", "Qty", "Amount"];
foreach ($headers as $i =&gt; $text) {
    $cell = $app-&gt;convertFromRowColNumber(1, $i + 1);
    $ws-&gt;getRange("{$cell}:{$cell}")-&gt;setValue($text);
}

// --- detail rows (2-4) ---
$data = [["Apple", 100, 3], ["Orange", 50, 6], ["Grape", 150, 2]];
foreach ($data as $i =&gt; [$name, $price, $qty]) {
    $row = $i + 2;
    $ws-&gt;getRange("A{$row}:A{$row}")-&gt;setValue($name);
    $ws-&gt;getRange("B{$row}:B{$row}")-&gt;setNumberValue($price);
    $ws-&gt;getRange("C{$row}:C{$row}")-&gt;setNumberValue($qty);
}

// --- amount column (D2:D4) ---
// Each row needs its own relative reference, so use setFormulaArray.
// Formula strings do NOT take a leading "=".
$formulas = [];
for ($i = 0; $i &lt; count($data); $i++) {
    $f = new \phposbxl\dto\InputFormulaObject();
    $row = $i + 2;
    $f-&gt;setA1C1($app-&gt;convertFromRowColNumber($i + 2, 4));   // D2, D3, D4
    $f-&gt;setFormula("B{$row}*C{$row}");                       // "B2*C2" ...
    $formulas[] = $f;
}
$ws-&gt;setFormulaArray($formulas);

// --- total (D5) ---
$ws-&gt;getRange("C5:C5")-&gt;setValue("Total");
$ws-&gt;getRange("D5:D5")-&gt;setFormula("SUM(D2:D4)");

// --- read back ---
// getValue(true) returns raw values = formula results
$values = $ws-&gt;getRange("A1:D5")-&gt;getValue(true);
echo "Amount D2:D4 = " . implode(", ", [$values["D2"], $values["D3"], $values["D4"]]) . "\n";
echo "Total: " . $values["D5"] . "\n";

// --- save and close ---
$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// values-formulas.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("sales-wasm.xlsx");   // new workbook
const ws = wb.openWorkSheet("Sheet1");              // open the default sheet

// --- header (row 1) ---
// app.convertFromRowColNumber(row, col) gives the A1 address for (1,1) -&gt; "A1"
["Product", "Unit Price", "Qty", "Amount"].forEach((text, i) =&gt; {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- detail rows (2-4) ---
const data = [["Apple", 100, 3], ["Orange", 50, 6], ["Grape", 150, 2]];
data.forEach(([name, price, qty], i) =&gt; {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- amount column (D2:D4) ---
// Each row needs its own relative reference, so use setFormulaArray.
// Formula strings do NOT take a leading "=".
const formulas = data.map((_, i) =&gt; {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(i + 2, 4));  // D2, D3, D4
    f.setFormula(`B${i + 2}*C${i + 2}`);               // "B2*C2" ...
    return f;
});
ws.setFormulaArray(formulas);

// --- total (D5) ---
ws.getRange("C5:C5").setValue("Total");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- read back ---
// getValue(true) returns raw values = formula results
const values = ws.getRange("A1:D5").getValue(true);
console.log("Amount D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("Total:", values["D5"]);

// --- save and close ---
wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">Amount D2:D4 = 300, 300, 300
Total: 900</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new nodeosbxl.App()</code> is the single entry point.</strong> From <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> you get a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkBook()</code>, then a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet()</code>, then a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Cell addresses are strings like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:B2"</code>.</strong> Even a single cell uses the range form <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:A1"</code>. To build them from 1-based row/column numbers use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber(row, col)</code>, or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber2(startRow, startCol, endRow, endCol)</code> for ranges.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Each value type has its own setter.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(string)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(number)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(DateTimeObject)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue(date string)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(bool)</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> means "raw value"</strong> (formula result). Without it you get the display value (after number formatting). The return value is a map like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ "A1": "100", ... }</code> and all values are strings; convert with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Number(...)</code> when comparing.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Finish with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.save()</code> then <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.close()</code>.</strong> Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.saveAs(path)</code> to write a second file, but note that <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs()</code> does <strong>not</strong> switch the destination of later saves (see <a href="general-01-overview.php#notes" class="text-primary-600 hover:underline">overview caveats</a>).</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Do not prefix formulas with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code>.</strong> Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code>, not <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("=SUM(A1:A3)")</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue</code> on a multi-cell Range writes the same value to every cell.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("A1:C1").setValue("X")</code> puts <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"X"</code> into A1, B1 and C1.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula()</code> does not auto-fill relative references.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("B1:B3").setFormula("A1*2")</code> puts the formula into <strong>B1 only</strong>; B2 / B3 stay empty.<br>The third argument variant <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("A1*2", false, true)</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAllCell = true</code>) merely copies the same string into every cell — it does <strong>not</strong> become <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A2*2</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A3*2</code> (all of B1:B3 end up with the same value). To shift formulas per row, use <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray()</code> + <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputFormulaObject</code></strong> (A1C1 and formula per cell) as in this sample.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>A license banner line is printed to stdout once per process.</strong> Keep it in mind when parsing standard output.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-01-overview.php" class="text-primary-600 hover:underline">Overview — object model and basic flow</a> / Next: <a href="general-03-sheets.php" class="text-primary-600 hover:underline">Working with books and sheets</a></li>
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
   "headline": "Building a Table with Values and Formulas",
   "description": "Building a Table with Values and Formulas. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-02-values-formulas.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-02-values-formulas.php",
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
     "name": "Building a Table with Values and Formulas",
     "item": "https://osb-office.com/en/sample/general-02-values-formulas.php"
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
