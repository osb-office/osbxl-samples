<?php
$lang = 'en';
$page_title = 'Conditional Formatting — rules, scales, data bars | osbxl Samples';
$no_tagline = true;
$meta_description = 'Conditional Formatting — rules, scales, data bars. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,formatconditions,cellvalue,colorscale,databar,top10,aboveaverage,iconset,uniquevalues,apply,removecondition,colorobject,patternfill';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Conditional Formatting — rules, scales, data bars</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-09-conditional-format.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7">Set <strong>conditional formatting rules</strong> programmatically, such as "make only cells greater than 60 red" or "show the magnitude of values with a gradient or a bar". Also <strong>read back the count, type, thresholds and applied ranges of the registered rules</strong> and <strong>remove</strong> them.</p>
<p class="mb-4 text-slate-700 leading-7">Conditional formatting is handled by the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditions</code> class obtained via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getFormatConditions()</code>. The flow is always these 3 steps:</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li>Create a condition object from <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> (e.g. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CellValueObject</code>) and <strong>set its apply range and format</strong></li>
  <li>Register it with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">fc.set…Condition(priority, obj)</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">priority</code> is the evaluation order,</li>
</ol>
<p class="mb-4 text-slate-700 leading-7">1-based and unique per sheet)</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>Call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">fc.apply()</code> after all conditions are registered</strong> (required)</li>
</ol>
<h2 id="classes-used" class="text-2xl font-bold mt-10 mb-4">Classes used</h2>
<h3 id="formatconditions-main-methods" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditions</code> (main methods)</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Condition type</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Set</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Read back</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Condition object</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Cell value (greater, less, ...)</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCellValueCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCellValueCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CellValueObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Cell value range (between)</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCellValueRangeCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCellValueRangeCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CellValueRangeObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Color scale</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColorScaleFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColorScaleFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.ColorScaleObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Data bar</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataBarFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDataBarFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.DataBarObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Top / bottom N</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTop10FormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTop10FormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.Top10Object</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Above/below average</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAboveAverageCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAboveAverageCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.AboveAverageObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Icon set</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setIconsetFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getIconsetFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.IconsetConditionObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Unique values</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getUniqueValuesFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.UniqueValuesObject</code></td></tr>
  </tbody></table></div>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Other methods</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Evaluates and reflects the conditions. <strong>Always call after setting all conditions.</strong></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code></td><td class="border border-surface-200 px-3 py-2 align-top">List of all rules; a nested map <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ priority: { cellRange: conditionType } }</code>.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeCondition(p)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeAllCondition()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Remove one rule / all rules.</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAppliedFontObject(range)</code> etc.</td><td class="border border-surface-200 px-3 py-2 align-top">Get the cells that met a condition and actually received formatting (use after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code>).</td></tr>
  </tbody></table></div>
<h3 id="methods-common-to-all-condition-objects" class="text-xl font-bold mt-8 mb-3">Methods common to all condition objects</h3>
<p class="mb-4 text-slate-700 leading-7">They all have <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1(applyRange)</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFont(dto.FontObject)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPatternFill(dto.PatternFillObject)</code>. Colors are set on a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.ColorObject</code> via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHexColor()</code> with a <strong>6-digit hex string without <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#</code></strong> (e.g. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"9C0006"</code>).</p>
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
        <pre><code class="language-javascript">// conditional-format.js — adding, reading back and removing conditional format rules
const { nodeosbxl, enums, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("condfmt.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// color helper: ColorObject takes a 6-digit hex without "#"
const color = (hex) =&gt; {
    const c = new dto.ColorObject();
    c.setHexColor(hex);
    return c;
};

// ---------- data table ----------
// col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"].forEach((h, i) =&gt; {
    const col = "ABCDE"[i];
    ws.getRange(`${col}1:${col}1`).setValue(h);
});
const sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
sales.forEach((v, i) =&gt; {
    ws.getRange(`A${i + 2}:A${i + 2}`).setNumberValue(i + 1);
    for (const col of ["B", "C", "D", "E"]) {
        ws.getRange(`${col}${i + 2}:${col}${i + 2}`).setNumberValue(v);
    }
});

const fc = ws.getFormatConditions();

// ---------- (1) cell value rule: greater than 60 -&gt; dark red bold + light red fill ----------
const cellVal = new dto.CellValueObject();
cellVal.setValueCondition(60, enums.XlFormatConditionOperator.FormatConditionGreater);
cellVal.setA1C1("B2:B11");
const redFont = new dto.FontObject();
redFont.setBold(true);
redFont.setColorObject(color("9C0006"));
cellVal.setFont(redFont);
const pinkFill = new dto.PatternFillObject();
pinkFill.setPattern(enums.XlPatternType.PatternSolid);
pinkFill.setCellColor(color("FFC7CE"));
cellVal.setPatternFill(pinkFill);
fc.setCellValueCondition(1, cellVal);

// ---------- (2) color scale: 3-color gradient on column C ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);

// ---------- (3) data bar: column D ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);

// ---------- (4) top 3 + above average: column E ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true);                       // true = top (false = bottom)
top3.setA1C1("E2:E11");
const blueFont = new dto.FontObject();
blueFont.setBold(true);
blueFont.setColorObject(color("0070C0"));
top3.setFont(blueFont);
fc.setTop10FormatCondition(4, top3);

const aboveAvg = new dto.AboveAverageObject();
aboveAvg.setAboveBelow(enums.XlAboveBelow.XlAboveAverage);
aboveAvg.setA1C1("E2:E11");
const greenFont = new dto.FontObject();
greenFont.setBold(true);
greenFont.setColorObject(color("006100"));
aboveAvg.setFont(greenFont);
fc.setAboveAverageCondition(5, aboveAvg);

fc.apply();                                  // apply() is required after setting all conditions
console.log("[1] Registered 5 rules and called apply()");

// ---------- read back the rule list ----------
const typeName = Object.fromEntries(
    Object.entries(enums.XlFormatConditionType).map(([name, val]) =&gt; [val, name]));
console.log(`[2] rule count = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
    for (const [range, type] of Object.entries(ranges)) {
        console.log(`    priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
    }
}

// ---------- read back each rule's details ----------
console.log("[3] Reading back rule details");
const g1 = fc.getCellValueCondition(1);
console.log(`    cell value : ${g1.getA1C1()} value=${g1.getConditionValue()}` +
    ` operator=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
    ` fontColor=${g1.getFont().getColorObject().getHexColor()} bold=${g1.getFont().isBold()}` +
    ` fill=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(`    color scale: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
    ` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, threshold=${g2.getMediumValue()})` +
    ` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(`    data bar: color=${g3.getDataBarColor().getHexColor()}` +
    ` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(`    top N: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(`    above average: aboveBelow=${g5.getAboveBelow()}` +
    ` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);

// ---------- cells that met the conditions and got formatted (available after apply()) ----------
const byRow = (a, b) =&gt; Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] Cells that received formatting");
console.log(`    col B (&gt;60)     : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(`    col E (top/avg) : ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(`    col D data bars : ${Object.values(appliedBar).filter(Boolean).length} cells`);

// ---------- removal ----------
fc.removeCondition(4);                       // remove only the top-3 rule
console.log(`[5] rule count after removeCondition(4) = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition();                     // remove all rules
console.log(`[6] rule count after removeAllCondition() = ${Object.keys(fc.getConditionTypes()).length}`);

// re-register the cell value rule and save (so the effect is visible when the file is opened)
fc.setCellValueCondition(1, cellVal);
fc.apply();

wb.save();
wb.close();
console.log("[done] saved condfmt.xlsx (cell-value rule re-registered)");</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># conditional-format.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl

e = pyosbxl.enums


def color(hex):  # ColorObject takes a 6-digit hex without "#"
    c = pyosbxl.dto.ColorObject()
    c.setHexColor(hex)
    return c


def num(v):  # print numbers like JS does (60.0 -&gt; "60")
    return str(int(v)) if float(v) == int(v) else str(v)


def bol(v):  # print booleans like JS does (True -&gt; "true")
    return str(bool(v)).lower()


app = pyosbxl.App()
wb = app.createWorkBook("condfmt-py.xlsx")  # output text says condfmt.xlsx to match the Node.js sample
ws = wb.openWorkSheet("Sheet1")

# ---------- data table ----------
# col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
for i, h in enumerate(["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"]):
    col = "ABCDE"[i]
    ws.getRange(f"{col}1:{col}1").setValue(h)
sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49]
for i, v in enumerate(sales):
    ws.getRange(f"A{i + 2}:A{i + 2}").setNumberValue(i + 1)
    for col in ["B", "C", "D", "E"]:
        ws.getRange(f"{col}{i + 2}:{col}{i + 2}").setNumberValue(v)

fc = ws.getFormatConditions()

# ---------- (1) cell value rule: greater than 60 -&gt; dark red bold + light red fill ----------
cellVal = pyosbxl.dto.CellValueObject()
cellVal.setValueCondition(60, e.XlFormatConditionOperator.FormatConditionGreater)
cellVal.setA1C1("B2:B11")
redFont = pyosbxl.dto.FontObject()
redFont.setBold(True)
redFont.setColorObject(color("9C0006"))
cellVal.setFont(redFont)
pinkFill = pyosbxl.dto.PatternFillObject()
pinkFill.setPattern(e.XlPatternType.PatternSolid)
pinkFill.setCellColor(color("FFC7CE"))
cellVal.setPatternFill(pinkFill)
fc.setCellValueCondition(1, cellVal)

# ---------- (2) color scale: 3-color gradient on column C ----------
colorScale = pyosbxl.dto.ColorScaleObject()
colorScale.setA1C1("C2:C11")
colorScale.setMinimumCondition(color("F8696B"), e.XlConditionValueType.ConditionValueTypeMinimum)
colorScale.setMediumCondition(color("FFEB84"), e.XlConditionValueType.ConditionValueTypePercentile, 50)
colorScale.setMaximumCondition(color("63BE7B"), e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setColorScaleFormatCondition(2, colorScale)

# ---------- (3) data bar: column D ----------
dataBar = pyosbxl.dto.DataBarObject()
dataBar.setA1C1("D2:D11")
dataBar.setDataBarColor(color("638EC6"))
dataBar.setMinimumCondition(e.XlConditionValueType.ConditionValueTypeMinimum)
dataBar.setMaximumCondition(e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setDataBarFormatCondition(3, dataBar)

# ---------- (4) top 3 + above average: column E ----------
top3 = pyosbxl.dto.Top10Object()
top3.setRank(3, True)                        # True = top (False = bottom)
top3.setA1C1("E2:E11")
blueFont = pyosbxl.dto.FontObject()
blueFont.setBold(True)
blueFont.setColorObject(color("0070C0"))
top3.setFont(blueFont)
fc.setTop10FormatCondition(4, top3)

aboveAvg = pyosbxl.dto.AboveAverageObject()
aboveAvg.setAboveBelow(e.XlAboveBelow.XlAboveAverage)
aboveAvg.setA1C1("E2:E11")
greenFont = pyosbxl.dto.FontObject()
greenFont.setBold(True)
greenFont.setColorObject(color("006100"))
aboveAvg.setFont(greenFont)
fc.setAboveAverageCondition(5, aboveAvg)

fc.apply()                                   # apply() is required after setting all conditions
print("[1] Registered 5 rules and called apply()")

# ---------- read back the rule list ----------
types = fc.getConditionTypes()
print(f"[2] rule count = {len(types)}")
for prio in sorted(types):
    for rng, t in types[prio].items():
        print(f"    priority {prio}: {rng} → {t.name} (type={int(t)})")

# ---------- read back each rule's details ----------
print("[3] Reading back rule details")
g1 = fc.getCellValueCondition(1)
print(f"    cell value : {g1.getA1C1()} value={num(g1.getConditionValue())}"
      f" operator={int(g1.getConditionOperetor())}"
      f" (Greater={int(e.XlFormatConditionOperator.FormatConditionGreater)})"
      f" fontColor={g1.getFont().getColorObject().getHexColor()} bold={bol(g1.getFont().isBold())}"
      f" fill={g1.getPatternFill().getCellColor().getHexColor()}")
g2 = fc.getColorScaleFormatCondition(2)
print(f"    color scale: min={g2.getMinimumColor().getHexColor()} (type={int(g2.getMinimumValueType())})"
      f" mid={g2.getMediumColor().getHexColor()} (type={int(g2.getMediumValueType())}, threshold={num(g2.getMediumValue())})"
      f" max={g2.getMaximumColor().getHexColor()} (type={int(g2.getMaximumValueType())})")
g3 = fc.getDataBarFormatCondition(3)
print(f"    data bar: color={g3.getDataBarColor().getHexColor()}"
      f" minType={int(g3.getMinimumCondition())} maxType={int(g3.getMaximumCondition())}")
g4 = fc.getTop10FormatCondition(4)
print(f"    top N: rank={g4.getRank()} isTop={bol(g4.getTopBottom())}")
g5 = fc.getAboveAverageCondition(5)
print(f"    above average: aboveBelow={int(g5.getAboveBelow())}"
      f" (XlAboveAverage={int(e.XlAboveBelow.XlAboveAverage)})")

# ---------- cells that met the conditions and got formatted (available after apply()) ----------
byRow = lambda a: int(a[1:])
print("[4] Cells that received formatting")
appliedB = sorted(fc.getAppliedFontObject("B2:B11").keys(), key=byRow)
print(f"    col B (&gt;60)     : {' '.join(appliedB)}")
appliedE = sorted(fc.getAppliedFontObject("E2:E11").keys(), key=byRow)
print(f"    col E (top/avg) : {' '.join(appliedE)}")
appliedBar = fc.getAppliedDataBar("D2:D11")
print(f"    col D data bars : {sum(1 for v in appliedBar.values() if v)} cells")

# ---------- removal ----------
fc.removeCondition(4)                        # remove only the top-3 rule
print(f"[5] rule count after removeCondition(4) = {len(fc.getConditionTypes())}")
fc.removeAllCondition()                      # remove all rules
print(f"[6] rule count after removeAllCondition() = {len(fc.getConditionTypes())}")

# re-register the cell value rule and save (so the effect is visible when the file is opened)
fc.setCellValueCondition(1, cellVal)
fc.apply()

wb.save()
wb.close()
print("[done] saved condfmt.xlsx (cell-value rule re-registered)")</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// conditional-format.java — checks how much of the Node.js sample javaosbxl supports
//
// [API gap] In javaosbxl 1.5.0, the condition-object wrappers (CellValueObjectWrapper etc.)
// do not expose the base-class methods setA1C1() / setFont() / setPatternFill().
// With no way to set the apply range (A1C1), setCellValueCondition() throws
// "apply range(A1C1) must be set", so conditional format rules cannot be registered.
// This block demonstrates that registering the same rule as the Node.js version throws,
// while apply() / getConditionTypes() / removeCondition() / removeAllCondition() can all
// be called without exceptions, making the gap explicit.
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("condfmt-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- data table (same as the Node.js version) ----------
        String[] headers = {"Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"};
        for (int i = 0; i &lt; headers.length; i++) {
            String col = "ABCDE".substring(i, i + 1);
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        int[] sales = {45, 78, 23, 91, 56, 34, 87, 12, 65, 49};
        for (int i = 0; i &lt; sales.length; i++) {
            ws.getRange("A" + (i + 2) + ":A" + (i + 2)).setNumberValue(i + 1);
            for (String col : new String[]{"B", "C", "D", "E"}) {
                ws.getRange(col + (i + 2) + ":" + col + (i + 2)).setNumberValue(sales[i]);
            }
        }

        FormatConditionsWrapper fc = ws.getFormatConditions();

        // ---------- (1) cell value rule: greater than 60 ----------
        // The Node.js version calls cellVal.setA1C1("B2:B11") / setFont(...) / setPatternFill(...),
        // but javaosbxl's CellValueObjectWrapper does not have these methods (API gap)
        CellValueObjectWrapper cellVal = new CellValueObjectWrapper();
        cellVal.setValueCondition(60, XlFormatConditionOperator.FormatConditionGreater);
        try {
            fc.setCellValueCondition(1, cellVal);
            System.out.println("[1] Registered the cell-value rule");
        } catch (RuntimeException e) {
            // registration always throws because A1C1 cannot be set (known API gap in javaosbxl 1.5.0)
            System.out.println("[1] Registering the cell-value rule threw: " + e.getMessage());
        }
        fc.apply();  // apply() can be called without exceptions even with 0 rules

        // ---------- read back the rule list (nothing was registered, so 0) ----------
        Map&lt;Integer, Map&lt;String, XlFormatConditionType&gt;&gt; types = fc.getConditionTypes();
        System.out.println("[2] rule count = " + types.size());

        // ---------- removal methods also work without exceptions at 0 rules ----------
        fc.removeCondition(4);
        System.out.println("[5] rule count after removeCondition(4) = " + fc.getConditionTypes().size());
        fc.removeAllCondition();
        System.out.println("[6] rule count after removeAllCondition() = " + fc.getConditionTypes().size());

        wb.save();
        wb.close();
        System.out.println("[done] saved condfmt-java.xlsx (0 rules because javaosbxl cannot register conditional format rules)");
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// conditional-format.php — same verification as the Node.js sample (phposbxl). Body only (&lt;?php / autoload provided by the harness)

// Output helpers: match the Node.js stdout representation (bool -&gt; "true"/"false")
$bol = fn($v) =&gt; $v ? "true" : "false";
// Color helper: ColorObject takes a 6-digit hex without "#"
$color = function (string $hex): \phposbxl\dto\ColorObject {
    $c = new \phposbxl\dto\ColorObject();
    $c-&gt;setHexColor($hex);
    return $c;
};

$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("condfmt-php.xlsx"); // output text says condfmt.xlsx to match the Node.js sample
$ws = $wb-&gt;openWorkSheet("Sheet1");

// ---------- data table ----------
// col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
foreach (["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"] as $i =&gt; $h) {
    $col = "ABCDE"[$i];
    $ws-&gt;getRange("{$col}1:{$col}1")-&gt;setValue($h);
}
$sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
foreach ($sales as $i =&gt; $v) {
    $r = $i + 2;
    $ws-&gt;getRange("A{$r}:A{$r}")-&gt;setNumberValue($i + 1);
    foreach (["B", "C", "D", "E"] as $col) {
        $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;setNumberValue($v);
    }
}

$fc = $ws-&gt;getFormatConditions();

// ---------- (1) cell value rule: greater than 60 -&gt; dark red bold + light red fill ----------
$cellVal = new \phposbxl\dto\CellValueObject();
$cellVal-&gt;setValueCondition(60, \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater);
$cellVal-&gt;setA1C1("B2:B11");
$redFont = new \phposbxl\dto\FontObject();
$redFont-&gt;setBold(true);
$redFont-&gt;setColorObject($color("9C0006"));
$cellVal-&gt;setFont($redFont);
$pinkFill = new \phposbxl\dto\PatternFillObject();
$pinkFill-&gt;setPattern(\phposbxl\enums\XlPatternType::PatternSolid);
$pinkFill-&gt;setCellColor($color("FFC7CE"));
$cellVal-&gt;setPatternFill($pinkFill);
$fc-&gt;setCellValueCondition(1, $cellVal);

// ---------- (2) color scale: 3-color gradient on column C ----------
$colorScale = new \phposbxl\dto\ColorScaleObject();
$colorScale-&gt;setA1C1("C2:C11");
$colorScale-&gt;setMinimumCondition($color("F8696B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$colorScale-&gt;setMediumCondition($color("FFEB84"), \phposbxl\enums\XlConditionValueType::ConditionValueTypePercentile, 50);
$colorScale-&gt;setMaximumCondition($color("63BE7B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc-&gt;setColorScaleFormatCondition(2, $colorScale);

// ---------- (3) data bar: column D ----------
$dataBar = new \phposbxl\dto\DataBarObject();
$dataBar-&gt;setA1C1("D2:D11");
$dataBar-&gt;setDataBarColor($color("638EC6"));
$dataBar-&gt;setMinimumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$dataBar-&gt;setMaximumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc-&gt;setDataBarFormatCondition(3, $dataBar);

// ---------- (4) top 3 + above average: column E ----------
$top3 = new \phposbxl\dto\Top10Object();
$top3-&gt;setRank(3, true);                      // true = top (false = bottom)
$top3-&gt;setA1C1("E2:E11");
$blueFont = new \phposbxl\dto\FontObject();
$blueFont-&gt;setBold(true);
$blueFont-&gt;setColorObject($color("0070C0"));
$top3-&gt;setFont($blueFont);
$fc-&gt;setTop10FormatCondition(4, $top3);

$aboveAvg = new \phposbxl\dto\AboveAverageObject();
$aboveAvg-&gt;setAboveBelow(\phposbxl\enums\XlAboveBelow::XlAboveAverage);
$aboveAvg-&gt;setA1C1("E2:E11");
$greenFont = new \phposbxl\dto\FontObject();
$greenFont-&gt;setBold(true);
$greenFont-&gt;setColorObject($color("006100"));
$aboveAvg-&gt;setFont($greenFont);
$fc-&gt;setAboveAverageCondition(5, $aboveAvg);

$fc-&gt;apply();                                 // apply() is required after setting all conditions
echo "[1] Registered 5 rules and called apply()\n";

// ---------- read back the rule list ----------
$types = $fc-&gt;getConditionTypes();            // array&lt;int, array&lt;string, int&gt;&gt;
ksort($types);                                // iterate in ascending priority order, like Object.entries in the Node.js version
echo "[2] rule count = " . count($types) . "\n";
foreach ($types as $prio =&gt; $ranges) {
    foreach ($ranges as $range =&gt; $type) {
        $name = \phposbxl\enums\XlFormatConditionType::from($type)-&gt;name;
        echo "    priority {$prio}: {$range} → {$name} (type={$type})\n";
    }
}

// ---------- read back each rule's details ----------
echo "[3] Reading back rule details\n";
$g1 = $fc-&gt;getCellValueCondition(1);
$greater = \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater-&gt;value;
echo "    cell value : {$g1-&gt;getA1C1()} value={$g1-&gt;getConditionValue()}"
    . " operator={$g1-&gt;getConditionOperetor()} (Greater={$greater})"
    . " fontColor={$g1-&gt;getFont()-&gt;getColorObject()-&gt;getHexColor()} bold=" . $bol($g1-&gt;getFont()-&gt;isBold())
    . " fill={$g1-&gt;getPatternFill()-&gt;getCellColor()-&gt;getHexColor()}\n";
$g2 = $fc-&gt;getColorScaleFormatCondition(2);
echo "    color scale: min={$g2-&gt;getMinimumColor()-&gt;getHexColor()} (type={$g2-&gt;getMinimumValueType()})"
    . " mid={$g2-&gt;getMediumColor()-&gt;getHexColor()} (type={$g2-&gt;getMediumValueType()}, threshold={$g2-&gt;getMediumValue()})"
    . " max={$g2-&gt;getMaximumColor()-&gt;getHexColor()} (type={$g2-&gt;getMaximumValueType()})\n";
$g3 = $fc-&gt;getDataBarFormatCondition(3);
echo "    data bar: color={$g3-&gt;getDataBarColor()-&gt;getHexColor()}"
    . " minType={$g3-&gt;getMinimumCondition()} maxType={$g3-&gt;getMaximumCondition()}\n";
$g4 = $fc-&gt;getTop10FormatCondition(4);
echo "    top N: rank={$g4-&gt;getRank()} isTop=" . $bol($g4-&gt;getTopBottom()) . "\n";
$g5 = $fc-&gt;getAboveAverageCondition(5);
$xlAbove = \phposbxl\enums\XlAboveBelow::XlAboveAverage-&gt;value;
echo "    above average: aboveBelow={$g5-&gt;getAboveBelow()} (XlAboveAverage={$xlAbove})\n";

// ---------- cells that met the conditions and got formatted (available after apply()) ----------
$byRow = fn($a, $b) =&gt; (int)substr($a, 1) - (int)substr($b, 1);
echo "[4] Cells that received formatting\n";
$appliedB = array_keys($fc-&gt;getAppliedFontObject("B2:B11"));
usort($appliedB, $byRow);
echo "    col B (&gt;60)     : " . implode(" ", $appliedB) . "\n";
$appliedE = array_keys($fc-&gt;getAppliedFontObject("E2:E11"));
usort($appliedE, $byRow);
echo "    col E (top/avg) : " . implode(" ", $appliedE) . "\n";
$appliedBar = $fc-&gt;getAppliedDataBar("D2:D11");
echo "    col D data bars : " . count(array_filter($appliedBar)) . " cells\n";

// ---------- removal ----------
$fc-&gt;removeCondition(4);                      // remove only the top-3 rule
echo "[5] rule count after removeCondition(4) = " . count($fc-&gt;getConditionTypes()) . "\n";
$fc-&gt;removeAllCondition();                    // remove all rules
echo "[6] rule count after removeAllCondition() = " . count($fc-&gt;getConditionTypes()) . "\n";

// re-register the cell value rule and save (so the effect is visible when the file is opened)
$fc-&gt;setCellValueCondition(1, $cellVal);
$fc-&gt;apply();

$wb-&gt;save();
$wb-&gt;close();
echo "[done] saved condfmt.xlsx (cell-value rule re-registered)\n";</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// conditional-format.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("condfmt-wasm.xlsx"); // output text says condfmt.xlsx, same as the Node.js version
const ws = wb.openWorkSheet("Sheet1");

// color helper: ColorObject takes a 6-digit hex without "#"
const color = (hex) =&gt; {
    const c = new dto.ColorObject();
    c.setHexColor(hex);
    return c;
};

// ---------- data table ----------
// col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"].forEach((h, i) =&gt; {
    const col = "ABCDE"[i];
    ws.getRange(`${col}1:${col}1`).setValue(h);
});
const sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
sales.forEach((v, i) =&gt; {
    ws.getRange(`A${i + 2}:A${i + 2}`).setNumberValue(i + 1);
    for (const col of ["B", "C", "D", "E"]) {
        ws.getRange(`${col}${i + 2}:${col}${i + 2}`).setNumberValue(v);
    }
});

const fc = ws.getFormatConditions();

// ---------- (1) cell value rule: greater than 60 -&gt; dark red bold + light red fill ----------
const cellVal = new dto.CellValueObject();
cellVal.setValueCondition(60, enums.XlFormatConditionOperator.FormatConditionGreater);
cellVal.setA1C1("B2:B11");
const redFont = new dto.FontObject();
redFont.setBold(true);
redFont.setColorObject(color("9C0006"));
cellVal.setFont(redFont);
const pinkFill = new dto.PatternFillObject();
pinkFill.setPattern(enums.XlPatternType.PatternSolid);
pinkFill.setCellColor(color("FFC7CE"));
cellVal.setPatternFill(pinkFill);
fc.setCellValueCondition(1, cellVal);

// ---------- (2) color scale: 3-color gradient on column C ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);

// ---------- (3) data bar: column D ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);

// ---------- (4) top 3 + above average: column E ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true);                       // true = top (false = bottom)
top3.setA1C1("E2:E11");
const blueFont = new dto.FontObject();
blueFont.setBold(true);
blueFont.setColorObject(color("0070C0"));
top3.setFont(blueFont);
fc.setTop10FormatCondition(4, top3);

const aboveAvg = new dto.AboveAverageObject();
aboveAvg.setAboveBelow(enums.XlAboveBelow.XlAboveAverage);
aboveAvg.setA1C1("E2:E11");
const greenFont = new dto.FontObject();
greenFont.setBold(true);
greenFont.setColorObject(color("006100"));
aboveAvg.setFont(greenFont);
fc.setAboveAverageCondition(5, aboveAvg);

fc.apply();                                  // apply() is required after setting all conditions
console.log("[1] Registered 5 rules and called apply()");

// ---------- read back the rule list ----------
const typeName = Object.fromEntries(
    Object.entries(enums.XlFormatConditionType).map(([name, val]) =&gt; [val, name]));
console.log(`[2] rule count = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
    for (const [range, type] of Object.entries(ranges)) {
        console.log(`    priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
    }
}

// ---------- read back each rule's details ----------
console.log("[3] Reading back rule details");
const g1 = fc.getCellValueCondition(1);
console.log(`    cell value : ${g1.getA1C1()} value=${g1.getConditionValue()}` +
    ` operator=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
    ` fontColor=${g1.getFont().getColorObject().getHexColor()} bold=${g1.getFont().isBold()}` +
    ` fill=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(`    color scale: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
    ` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, threshold=${g2.getMediumValue()})` +
    ` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(`    data bar: color=${g3.getDataBarColor().getHexColor()}` +
    ` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(`    top N: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(`    above average: aboveBelow=${g5.getAboveBelow()}` +
    ` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);

// ---------- cells that met the conditions and got formatted (available after apply()) ----------
const byRow = (a, b) =&gt; Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] Cells that received formatting");
console.log(`    col B (&gt;60)     : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(`    col E (top/avg) : ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(`    col D data bars : ${Object.values(appliedBar).filter(Boolean).length} cells`);

// ---------- removal ----------
fc.removeCondition(4);                       // remove only the top-3 rule
console.log(`[5] rule count after removeCondition(4) = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition();                     // remove all rules
console.log(`[6] rule count after removeAllCondition() = ${Object.keys(fc.getConditionTypes()).length}`);

// re-register the cell value rule and save (so the effect is visible when the file is opened)
fc.setCellValueCondition(1, cellVal);
fc.apply();

wb.save();
wb.close();
console.log("[done] saved condfmt.xlsx (cell-value rule re-registered)");</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] Registered 5 rules and called apply()
[2] rule count = 5
    priority 1: B2:B11 → CellValue (type=1)
    priority 2: C2:C11 → ColorScale (type=3)
    priority 3: D2:D11 → Databar (type=4)
    priority 4: E2:E11 → Top10 (type=5)
    priority 5: E2:E11 → AboveAverageCondition (type=12)
[3] Reading back rule details
    cell value : B2:B11 value=60 operator=5 (Greater=5) fontColor=9C0006 bold=true fill=FFC7CE
    color scale: min=F8696B (type=5) mid=FFEB84 (type=4, threshold=50) max=63BE7B (type=6)
    data bar: color=638EC6 minType=5 maxType=6
    top N: rank=3 isTop=true
    above average: aboveBelow=0 (XlAboveAverage=0)
[4] Cells that received formatting
    col B (&gt;60)     : B3 B5 B8 B10
    col E (top/avg) : E3 E5 E6 E8 E10
    col D data bars : 10 cells
[5] rule count after removeCondition(4) = 4
[6] rule count after removeAllCondition() = 0
[done] saved condfmt.xlsx (cell-value rule re-registered)</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">B3 B5 B8 B10</code> in column B are the cells whose values are 78, 91, 87 and 65 (= greater than 60); <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">E3 E5 E6 E8 E10</code> in column E are the union of "top 3" and "above the average (54.2)". Visual formatting cannot be verified programmatically, so the sample confirms the rules were set as intended by <strong>reading back the rule count, types, thresholds and the cells that got formatted</strong> and printing them.</p>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<h3 id="priority-evaluation-order" class="text-xl font-bold mt-8 mb-3">priority (evaluation order)</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Rules are managed by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">priority</code> (1-based evaluation order) and are <strong>unique per sheet</strong>. The keys of the map returned by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code> are these priorities.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Registering at an existing priority shifts the existing rule back</strong> (measured). For example, registering twice at priority 1 moves the first rule to priority 2.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Removing with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeCondition(p)</code> <strong>bumps the later priorities up</strong>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">The read-back methods (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCellValueCondition(p)</code> etc.) <strong>throw when a different type of rule sits at that priority</strong>. Check the types with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code> first.</li>
</ul>
<h3 id="what-apply-does" class="text-xl font-bold mt-8 mb-3">What apply() does</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Conditions are not evaluated until you call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code>.</strong> Measured: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAppliedFontObject()</code> returns an empty <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{}</code> before <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code>, and after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> it returned the cells meeting the condition (B3 B5 B8 B10 in this sample).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">Call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> every time you add or remove rules.</li>
</ul>
<h3 id="main-enums" class="text-xl font-bold mt-8 mb-3">Main enums</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">enum</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Common values</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlFormatConditionOperator</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionGreater</code>(5) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionGreaterEqual</code>(7) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionLess</code>(6) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionLessEqual</code>(8) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionEqual</code>(3) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionBetween</code>(1)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlFormatConditionType</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CellValue</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ColorScale</code>(3) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Databar</code>(4) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Top10</code>(5) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlIconSet</code>(6) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">UniqueValues</code>(8) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">AboveAverageCondition</code>(12)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlConditionValueType</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypeMinimum</code>(5) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypeMaximum</code>(6) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypePercentile</code>(4) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypeNumeric</code>(1)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlAboveBelow</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlAboveAverage</code>(0) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlBelowAverage</code>(2) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlAboveStdDev</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlBelowStdDev</code>(3)</td></tr>
  </tbody></table></div>
<h3 id="persistence" class="text-xl font-bold mt-8 mb-3">Persistence</h3>
<p class="mb-4 text-slate-700 leading-7">Conditional formatting is saved into the .xlsx by <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.save()</code>, and reopening the file with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.openWorkBook()</code> restores <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code> and the individual read-back methods (verified).</p>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Do not forget <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code>.</strong> Forgetting it does not throw — the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getApplied*</code> methods simply return empty results, which makes the mistake easy to miss (see <a href="#what-apply-does" class="text-primary-600 hover:underline">Notes</a>).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The operator getter is spelled <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionOperetor()</code>.</strong> It is <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Operetor</code></strong>, not <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Operator</code> (the library's own spelling), so be careful.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Data bar read-back method names differ from color scale.</strong> On <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DataBarObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumCondition()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumCondition()</code> return the <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlConditionValueType</code> (kind of threshold)</strong> and the values come from <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumValue()</code>. On <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ColorScaleObject</code>, the kind is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumValueType()</code> and the color is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumColor()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Colors are passed to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHexColor()</code> as a 6-digit hex without <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#</code></strong> (e.g. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"9C0006"</code>).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The second argument of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Top10Object.setRank(rank, isTop)</code> is a bool for top vs bottom</strong> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> = top N, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code> = bottom N). Despite the name "Top10", N can be any number.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Every condition object needs its apply range set with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1()</code>.</strong> The condition object owns the range; the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">set…Condition()</code> calls take no range argument.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**The <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getApplied*</code> methods return only "the cells that met the condition and actually got formatted"** — not the whole apply range of the rule.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">Tables and Comments</a> / Next: <a href="general-10-charts.php" class="text-primary-600 hover:underline">Charts Basics</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/nodeosbxl.FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/en/nodejs/classes/dto.CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/en/nodejs/classes/dto.ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/en/nodejs/classes/dto.DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/en/nodejs/classes/dto.Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/en/nodejs/enums/enums.XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/en/nodejs/enums/enums.XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/en/nodejs/enums/enums.XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/en/python/pyosbxl.dto.CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/en/python/pyosbxl.dto.ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/en/python/pyosbxl.dto.DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/en/python/pyosbxl.dto.Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/FormatConditionsWrapper.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/CellValueObjectWrapper.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/ColorScaleObjectWrapper.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/DataBarObjectWrapper.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/dto/Top10ObjectWrapper.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/en/php/classes/phposbxl-dto-Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/osbxl.FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/en/wasm/classes/dto.CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/en/wasm/classes/dto.ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/en/wasm/classes/dto.DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/en/wasm/classes/dto.Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/en/wasm/enums/enums.XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/en/wasm/enums/enums.XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/en/wasm/enums/enums.XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Conditional Formatting — rules, scales, data bars",
   "description": "Conditional Formatting — rules, scales, data bars. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-09-conditional-format.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-09-conditional-format.php",
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
     "name": "Conditional Formatting — rules, scales, data bars",
     "item": "https://osb-office.com/en/sample/general-09-conditional-format.php"
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
