<?php
$lang = 'en';
$page_title = 'Value Types and Number Formats — text, numbers, dates | osbxl Samples';
$no_tagline = true;
$meta_description = 'Value Types and Number Formats — text, numbers, dates. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,setvalue,setnumbervalue,setdatevalue,setdatestringvalue,setbooleanvalue,datetimeobject,numberformat,getvalue,rawvalue,forcestring';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Value Types and Number Formats — text, numbers, dates</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-04-value-types.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Use the correct setter for each type of value you put into a cell (string / number / date / time / boolean), specify number formats, and understand the difference between <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code>.</p>
<h2 id="type-quick-reference" class="text-2xl font-bold mt-10 mb-4">Type quick reference</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Value to store</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Example</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">String</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue("Apple")</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Number</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(num)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(1234.5)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Date / time</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(dateTimeObject)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(d)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Date / time (from a string)</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue("2024-01-15")</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Boolean</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(bool)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(true)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Let the type be detected automatically</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue("2024/1/15")</code> → becomes a date</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Formula</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">All setters share the same second and third arguments <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(forceString?, numberFormat?)</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue</code> included).</p>
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
        <pre><code class="language-javascript">// value-types.js — per-type setters and number formats, raw vs. display values
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("value-types.xlsx");
const ws = wb.openWorkSheet("Sheet1");

const show = (a1) =&gt; {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`  ${a1}: disp=${JSON.stringify(r.getValue()[a1])}` +
                ` raw=${JSON.stringify(r.getValue(true)[a1])}` +
                ` fmt=${JSON.stringify(r.getNumberFormat()[a1])}`);
};

// --- dates &amp; times ---
// Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);

const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);

const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);

console.log("[Dates &amp; times]");
ws.getRange("A1:A1").setDateValue(date);                      // date only
ws.getRange("A2:A2").setDateValue(datetime);                  // date + time
ws.getRange("A3:A3").setDateValue(time);                      // time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15");        // from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");      // Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");  // number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals also work
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);

// --- numbers &amp; strings ---
console.log("[Numbers &amp; strings]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123");      // leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15");  // the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", true); // forceString=true forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0');   // quoted literal ("yen")
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red) → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);

// --- booleans ---
console.log("[Booleans]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE");     // a formula can produce a boolean too
["C1", "C2", "C3"].forEach(show);

// --- date serial values ---
console.log("[Serial values]");
console.log("  serial value of 2024/1/15 =", app.getNumericValue(date));
console.log("  with the 1904 date system =", app.getNumericValue(date, true));
console.log("  wb.isDate1904()           =", wb.isDate1904());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># value-types.py — same verification as the Node.js sample (pyosbxl)
import json
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("value-types-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# Print display value / raw value / format on one line, quoted like JSON.stringify
def show(a1):
    r = ws.getRange(f"{a1}:{a1}")
    disp = json.dumps(r.getValue()[a1], ensure_ascii=False)
    raw = json.dumps(r.getValue(True)[a1], ensure_ascii=False)      # True for the raw value
    fmt = json.dumps(r.getNumberFormat()[a1], ensure_ascii=False)
    print(f"  {a1}: disp={disp} raw={raw} fmt={fmt}")

def num(x):          # render integer-valued floats as ints (to match the JS output)
    return int(x) if isinstance(x, float) and x.is_integer() else x

def boolstr(x):      # render Python's True/False as JS-style true/false
    return str(x).lower()

# --- dates &amp; times ---
# Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 1, 15)

dt = pyosbxl.dto.DateTimeObject()   # variable named dt to avoid the stdlib module name datetime
dt.setYMD(2024, 1, 15)
dt.setHMS(13, 45, 30)

time = pyosbxl.dto.DateTimeObject()
time.setHMS(9, 5, 0)

print("[Dates &amp; times]")
ws.getRange("A1:A1").setDateValue(date)                       # date only
ws.getRange("A2:A2").setDateValue(dt)                         # date + time
ws.getRange("A3:A3").setDateValue(time)                       # time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15")         # from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日")       # Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, False, "yyyy-mm-dd")   # number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, False, "ggge年m月d日") # wareki + unquoted Japanese literals
for a in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"]:
    show(a)

# --- numbers &amp; strings ---
print("[Numbers &amp; strings]")
ws.getRange("B1:B1").setNumberValue(1234567.891)
ws.getRange("B2:B2").setNumberValue(1234567.891, False, "#,##0.00")
ws.getRange("B3:B3").setNumberValue(0.25, False, "0.0%")
ws.getRange("B4:B4").setValue("00123")       # leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15")   # the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", True) # forceString=True forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, False, '"円"#,##0')      # quoted literal
ws.getRange("B8:B8").setNumberValue(-1234.5, False, "#,##0;赤-#,##0") # bare color token (auto-converted to [赤])
for b in ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"]:
    show(b)

# --- booleans ---
print("[Booleans]")
ws.getRange("C1:C1").setBooleanValue(True)
ws.getRange("C2:C2").setBooleanValue(False)
ws.getRange("C3:C3").setFormula("TRUE")      # a formula can produce a boolean too (no leading "=")
for c in ["C1", "C2", "C3"]:
    show(c)

# --- date serial values ---
print("[Serial values]")
print("  serial value of 2024/1/15 =", num(app.getNumericValue(date)))
print("  with the 1904 date system =", num(app.getNumericValue(date, True)))
print("  wb.isDate1904()           =", boolstr(wb.isDate1904()))

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// value-types.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.Map;

public class Main {
    static AppWrapper app;
    static WorkSheetWrapper ws;

    // Quoting equivalent to JS JSON.stringify (assumes no control characters in the sample)
    static String jsonq(String s) {
        return "\"" + s.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
    }

    // Render integer-valued doubles as ints (to match the JS output)
    static String num(double d) {
        if (d == Math.rint(d) &amp;&amp; !Double.isInfinite(d)) return String.valueOf((long) d);
        return String.valueOf(d);
    }

    static void show(String a1) {
        RangeWrapper r = ws.getRange(a1 + ":" + a1);
        Map&lt;String, String&gt; disp = r.getValue();
        Map&lt;String, String&gt; raw = r.getValue(true);          // true for the raw value
        Map&lt;String, String&gt; fmt = r.getNumberFormat();
        System.out.println("  " + a1 + ": disp=" + jsonq(disp.get(a1))
                + " raw=" + jsonq(raw.get(a1))
                + " fmt=" + jsonq(fmt.get(a1)));
    }

    public static void main(String[] args) {
        app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("value-types-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        // --- dates &amp; times ---
        // Create a DateTimeObjectWrapper with no arguments, then set it with setYMD() / setHMS()
        DateTimeObjectWrapper date = new DateTimeObjectWrapper();
        date.setYMD(2024, 1, 15);

        DateTimeObjectWrapper datetime = new DateTimeObjectWrapper();
        datetime.setYMD(2024, 1, 15);
        datetime.setHMS(13, 45, 30);

        DateTimeObjectWrapper time = new DateTimeObjectWrapper();
        time.setHMS(9, 5, 0);

        System.out.println("[Dates &amp; times]");
        ws.getRange("A1:A1").setDateValue(date);                       // date only
        ws.getRange("A2:A2").setDateValue(datetime);                   // date + time
        ws.getRange("A3:A3").setDateValue(time);                       // time only
        ws.getRange("A4:A4").setDateStringValue("2024-01-15");         // from a string
        ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");       // Japanese era (wareki) dates are recognized too
        ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");   // number format can be given at the same time
        ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals
        for (String a : new String[]{"A1", "A2", "A3", "A4", "A5", "A6", "A7"}) show(a);

        // --- numbers &amp; strings ---
        System.out.println("[Numbers &amp; strings]");
        ws.getRange("B1:B1").setNumberValue(1234567.891);
        ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
        ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
        ws.getRange("B4:B4").setValue("00123");       // leading zeros are kept as a string
        ws.getRange("B5:B5").setValue("2024/1/15");   // the generic setter auto-detects dates
        ws.getRange("B6:B6").setValue("12345", true); // forceString=true forces string handling
        ws.getRange("B7:B7").setNumberValue(1234.5, false, "\"円\"#,##0");    // quoted literal ("yen")
        ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red)
        for (String b : new String[]{"B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"}) show(b);

        // --- booleans ---
        System.out.println("[Booleans]");
        ws.getRange("C1:C1").setBooleanValue(true);
        ws.getRange("C2:C2").setBooleanValue(false);
        ws.getRange("C3:C3").setFormula("TRUE");      // a formula can produce a boolean too (no leading "=")
        for (String c : new String[]{"C1", "C2", "C3"}) show(c);

        // --- date serial values ---
        System.out.println("[Serial values]");
        System.out.println("  serial value of 2024/1/15 = " + num(app.getNumericValue(date)));
        System.out.println("  with the 1904 date system = " + num(app.getNumericValue(date, true)));
        System.out.println("  wb.isDate1904()           = " + wb.isDate1904());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// value-types.php — same verification as the Node.js sample (phposbxl)
// Body only (&lt;?php and the autoload require are added by the harness)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("value-types-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// Quoting equivalent to JS JSON.stringify
function jsonq($s) {
    return json_encode($s, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
}

// Render integer-valued floats as ints (to match the JS output)
function num($x) {
    return $x == floor($x) ? (string)(int)$x : (string)$x;
}

// Render PHP booleans as JS-style true/false
function boolstr($x) {
    return $x ? "true" : "false";
}

function show($a1) {
    global $ws;
    $r = $ws-&gt;getRange("{$a1}:{$a1}");
    echo "  {$a1}: disp=" . jsonq($r-&gt;getValue()[$a1])
       . " raw=" . jsonq($r-&gt;getValue(true)[$a1])
       . " fmt=" . jsonq($r-&gt;getNumberFormat()[$a1]) . "\n";
}

// --- dates &amp; times ---
// Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
$date = new \phposbxl\dto\DateTimeObject();
$date-&gt;setYMD(2024, 1, 15);

$datetime = new \phposbxl\dto\DateTimeObject();
$datetime-&gt;setYMD(2024, 1, 15);
$datetime-&gt;setHMS(13, 45, 30);

$time = new \phposbxl\dto\DateTimeObject();
$time-&gt;setHMS(9, 5, 0);

echo "[Dates &amp; times]\n";
$ws-&gt;getRange("A1:A1")-&gt;setDateValue($date);                       // date only
$ws-&gt;getRange("A2:A2")-&gt;setDateValue($datetime);                   // date + time
$ws-&gt;getRange("A3:A3")-&gt;setDateValue($time);                       // time only
$ws-&gt;getRange("A4:A4")-&gt;setDateStringValue("2024-01-15");          // from a string
$ws-&gt;getRange("A5:A5")-&gt;setDateStringValue("令和6年1月15日");        // Japanese era (wareki) dates are recognized too
$ws-&gt;getRange("A6:A6")-&gt;setDateValue($date, false, "yyyy-mm-dd");   // number format can be given at the same time
$ws-&gt;getRange("A7:A7")-&gt;setDateValue($date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals also work
foreach (["A1", "A2", "A3", "A4", "A5", "A6", "A7"] as $a) show($a);

// --- numbers &amp; strings ---
echo "[Numbers &amp; strings]\n";
$ws-&gt;getRange("B1:B1")-&gt;setNumberValue(1234567.891);
$ws-&gt;getRange("B2:B2")-&gt;setNumberValue(1234567.891, false, "#,##0.00");
$ws-&gt;getRange("B3:B3")-&gt;setNumberValue(0.25, false, "0.0%");
$ws-&gt;getRange("B4:B4")-&gt;setValue("00123");       // leading zeros are kept as a string
$ws-&gt;getRange("B5:B5")-&gt;setValue("2024/1/15");   // the generic setter auto-detects dates
$ws-&gt;getRange("B6:B6")-&gt;setValue("12345", true); // forceString=true forces string handling
$ws-&gt;getRange("B7:B7")-&gt;setNumberValue(1234.5, false, '"円"#,##0');      // quoted literal ("yen")
$ws-&gt;getRange("B8:B8")-&gt;setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red) → [Red]
foreach (["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"] as $bx) show($bx);

// --- booleans ---
echo "[Booleans]\n";
$ws-&gt;getRange("C1:C1")-&gt;setBooleanValue(true);
$ws-&gt;getRange("C2:C2")-&gt;setBooleanValue(false);
$ws-&gt;getRange("C3:C3")-&gt;setFormula("TRUE");      // a formula can produce a boolean too
foreach (["C1", "C2", "C3"] as $c) show($c);

// --- date serial values ---
echo "[Serial values]\n";
echo "  serial value of 2024/1/15 = " . num($app-&gt;getNumericValue($date)) . "\n";
echo "  with the 1904 date system = " . num($app-&gt;getNumericValue($date, true)) . "\n";
echo "  wb.isDate1904()           = " . boolstr($wb-&gt;isDate1904()) . "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// value-types.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("value-types-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

const show = (a1) =&gt; {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`  ${a1}: disp=${JSON.stringify(r.getValue()[a1])}` +
                ` raw=${JSON.stringify(r.getValue(true)[a1])}` +
                ` fmt=${JSON.stringify(r.getNumberFormat()[a1])}`);
};

// --- dates &amp; times ---
// Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);

const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);

const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);

console.log("[Dates &amp; times]");
ws.getRange("A1:A1").setDateValue(date);                      // date only
ws.getRange("A2:A2").setDateValue(datetime);                  // date + time
ws.getRange("A3:A3").setDateValue(time);                      // time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15");        // from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");      // Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");  // number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals also work
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);

// --- numbers &amp; strings ---
console.log("[Numbers &amp; strings]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123");      // leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15");  // the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", true); // forceString=true forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0');   // quoted literal ("yen")
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red) → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);

// --- booleans ---
console.log("[Booleans]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE");     // a formula can produce a boolean too
["C1", "C2", "C3"].forEach(show);

// --- date serial values ---
console.log("[Serial values]");
console.log("  serial value of 2024/1/15 =", app.getNumericValue(date));
console.log("  with the 1904 date system =", app.getNumericValue(date, true));
console.log("  wb.isDate1904()           =", wb.isDate1904());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[Dates &amp; times]
  A1: disp="2024/1/15" raw="45306" fmt="yyyy/m/d;@"
  A2: disp="2024/1/15 13:45:30" raw="45306.573263888888" fmt="yyyy/m/d h:mm:ss;@"
  A3: disp="9:05:00" raw="0.37847222222222221" fmt="h:mm:ss;@"
  A4: disp="2024/1/15" raw="45306" fmt="yyyy/m/d;@"
  A5: disp="2024/1/15" raw="45306" fmt="yyyy/m/d;@"
  A6: disp="2024-01-15" raw="45306" fmt="yyyy-mm-dd"
  A7: disp="令和6年1月15日" raw="45306" fmt="ggge年m月d日"
[Numbers &amp; strings]
  B1: disp="1234567.891" raw="1234567.8910000001" fmt="General"
  B2: disp="1,234,567.89" raw="1234567.8910000001" fmt="#,##0.00"
  B3: disp="25.0%" raw="0.25" fmt="0.0%"
  B4: disp="00123" raw="00123" fmt="General"
  B5: disp="2024/01/15" raw="45306" fmt="yyyy/mm/dd"
  B6: disp="12345" raw="12345" fmt="General"
  B7: disp="円1,235" raw="1234.5" fmt="\"円\"#,##0"
  B8: disp="-1,235" raw="-1234.5" fmt="#,##0;赤-#,##0"
[Booleans]
  C1: disp="TRUE" raw="TRUE" fmt="General"
  C2: disp="FALSE" raw="FALSE" fmt="General"
  C3: disp="TRUE" raw="TRUE" fmt="General"
[Serial values]
  serial value of 2024/1/15 = 45306
  with the 1904 date system = 43844
  wb.isDate1904()           = false</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<h3 id="getvalue-vs-getvalue-true" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> vs. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Call</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Returns</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Use it for</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code></td><td class="border border-surface-200 px-3 py-2 align-top"><strong>Display value</strong> (string after applying the number format)</td><td class="border border-surface-200 px-3 py-2 align-top">Report output, on-screen display, CSV export</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><strong>Raw value</strong> (internally stored value; dates are serial values, formulas are results)</td><td class="border border-surface-200 px-3 py-2 align-top">Calculation, comparison, numeric processing</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Both return a map like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ "A1": "...", ... }</code> and <strong>all values are strings</strong>. Convert with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Number(...)</code> when treating them as numbers.</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
  </div>
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">const raw = Number(ws.getRange("A1:A1").getValue(true)["A1"]);   // 45306
const disp = ws.getRange("A1:A1").getValue()["A1"];              // "2024/1/15"</code></pre>
      </div>
    </div>
</div>
<h3 id="dates" class="text-xl font-bold mt-8 mb-3">Dates</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Create a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DateTimeObject</code> with no arguments and set it with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setYMD()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHMS()</code>.</strong> You can also set the fields individually with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setYear()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMonth()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDay()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHour()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMinute()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setSecond()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>If you do not specify <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">numberFormat</code>, a default format is set automatically depending on the contents.</strong><ul class="mt-2 space-y-1 pl-5 list-disc text-slate-700"><li>Date only → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d;@</code></li><li>Time only → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">h:mm:ss;@</code></li><li>Date + time → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d h:mm:ss;@</code></li></ul></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue()</code> recognizes the same date strings Excel does.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"2024-01-15"</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"2024/1/15 13:45:30"</code> and even the Japanese era string <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"令和6年1月15日"</code> go in as dates.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.getNumericValue(dateTimeObject)</code> returns the serial value.</strong> Pass <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> as the second argument <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">is1904</code> to get the serial value in the 1904 date system.<br>Use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.isDate1904()</code> to check which system the book uses.</li>
</ul>
<h3 id="numbers-and-strings" class="text-xl font-bold mt-8 mb-3">Numbers and strings</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue()</code> auto-detects the type.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"123"</code> → number, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"2024/1/15"</code> → date, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"00123"</code> → string (leading zero preserved), <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"12A"</code> → string.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Use the dedicated setter when you want the intended type for sure.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue()</code> for numbers; set <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">forceString</code> to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> to pin a value as a string.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">numberFormat</code> can be given as the setter's third argument at the same time.</strong> The result is the same as calling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberFormat()</code> afterwards.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Set the number format AFTER putting the value in.</strong> If you call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue()</code> after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberFormat()</code>, the date setter <strong>overwrites</strong> the format with the default (measured).<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">// ✗ the format is lost
ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");
ws.getRange("A1:A1").setDateValue(date);      // → format gets overwritten with "yyyy/m/d;@"

// ✓ the format survives
ws.getRange("A1:A1").setDateValue(date);
ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");

// ✓ or give it at the same time via the setter's third argument
ws.getRange("A1:A1").setDateValue(date, false, "yyyy-mm-dd");</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue()</code> throws on strings it cannot recognize.</strong><div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">ws.getRange("A1:A1").setDateStringValue("this is not a date");
// → exception: can't recognize input value as DateTime</code></pre></div>To avoid the exception, pass <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> as the second argument <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">forceString</code>. When recognition fails, the value is then entered as a plain string.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Number formats also support Japanese text, currency symbols and color specs</strong> (verified 2026-09-30). Quoted literals (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"円"#,##0</code>) and unquoted Japanese text (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy年m月d日</code>) display as-is, and the Japanese era calendar (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ggge年m月d日</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">令和6年1月15日</code>) and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">¥#,##0</code> also work.<br>Color specs accept all three spellings <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[Red]</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[赤]</code> / bare <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">赤</code>, and all of them are stored as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[Red]</code>. Examples of verified formats: | Kind | Format | Example display | |---|---|---| | Date | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/mm/dd</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy-mm-dd</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yy/m/d</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">m/d</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">d-mmm-yyyy</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2024/1/15</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2024-01-15</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">15-Jan-2024</code> | | Date (Japanese) | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy年m月d日</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ggge年m月d日</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2024年1月15日</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">令和6年1月15日</code> | | Number | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">General</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0.00</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0.000</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0%</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0.0%</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1,235</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1,234.57</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">25.0%</code> | | Number (Japanese text &amp; currency) | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"円"#,##0</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">¥#,##0</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">円1,235</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">¥1,235</code> | | Number (negative section) | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0;[Red]-#,##0</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0;赤-#,##0</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">-1,235</code> (in red) |</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Raw number values contain double representation error.</strong> The raw value of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1234567.891</code> is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"1234567.8910000001"</code>. For comparisons and display, use <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> (the display value) or round the number first.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code> results have no leading <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code>.</strong> For <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code> returns <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"SUM(A1:A3)"</code>.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-03-sheets.php" class="text-primary-600 hover:underline">Working with books and worksheets</a> / Next: <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">Working with rows, columns and cell ranges</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/en/nodejs/classes/dto.DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/en/python/pyosbxl.dto.DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/RangeWrapper.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/DateTimeObjectWrapper.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/en/wasm/classes/dto.DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Value Types and Number Formats — text, numbers, dates",
   "description": "Value Types and Number Formats — text, numbers, dates. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-04-value-types.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-04-value-types.php",
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
     "name": "Value Types and Number Formats — text, numbers, dates",
     "item": "https://osb-office.com/en/sample/general-04-value-types.php"
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
