<?php
$lang = 'en';
$page_title = 'Charts Basics — add, configure, read back, delete | osbxl Samples';
$no_tagline = true;
$meta_description = 'Charts Basics — add, configure, read back, delete. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.';
$meta_keywords = 'osbxl,sample,Excel manipulation library,Excel generation library,addchartobject,chartwizard,getchartobject,deletechartobject,getseriescollection,getaxes,getlegend,charttitle,charttype';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">Charts Basics — add, configure, read back, delete</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ <a href="/sample/general-10-charts.php" class="text-primary-600 hover:underline">日本語</a></p><h2 id="what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">What you want to do</h2>
<p class="mb-4 text-slate-700 leading-7"><strong>Create charts (graphs)</strong> from a data table on a worksheet. Build two kinds — a clustered column chart and a line chart — set their titles, legend and axes, and <strong>read the results back as properties to verify them</strong>. Finally, delete a chart.</p>
<p class="mb-4 text-slate-700 leading-7">A chart's appearance cannot be checked from the console, so this sample prints the <strong>type, title string, series values, axis settings and chart count</strong> with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">console.log</code>, letting you programmatically confirm that everything was built as intended.</p>
<h2 id="classes-involved" class="text-2xl font-bold mt-10 mb-4">Classes involved</h2>
<p class="mb-4 text-slate-700 leading-7">Charts are split into two layers: the "frame" (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartObject</code>) and the "content" (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Chart</code>).</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Class</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">How to obtain</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Role</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.ChartObject</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.addChartObject(name, left, top, w, h)</code> / `ws.getChartObject(name\</td><td class="border border-surface-200 px-3 py-2 align-top">index)`</td><td class="border border-surface-200 px-3 py-2 align-top">Placement frame on the sheet. Position, size, name, visible/hidden</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Chart</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartObject.getChart()</code></td><td class="border border-surface-200 px-3 py-2 align-top">The chart itself. Built with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code>; handles type, data source, title, axes and series</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.ChartTitle</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getChartTitle()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Chart title. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setText()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getText()</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Axis</code></td><td class="border border-surface-200 px-3 py-2 align-top">`chart.getAxes(enums.XlAxisType.Category\</td><td class="border border-surface-200 px-3 py-2 align-top">Value)`</td><td class="border border-surface-200 px-3 py-2 align-top">Axis. Tick marks, max/min, axis title</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Series</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getSeriesCollection(index)</code></td><td class="border border-surface-200 px-3 py-2 align-top">One series. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValues()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getXValues()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPointSize()</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Legend</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getLegend()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Legend. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPosition()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPosition()</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">Chart-related methods on the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> side:</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Method</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Description</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject(name, left, top, width, height)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Adds a chart frame. <strong>Position and size are in points (pt)</strong></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject(name)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject(index)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Gets by name or index (<strong>1-based</strong>). Out of range throws</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteChartObject(name)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteChartObject(index)</code></td><td class="border border-surface-200 px-3 py-2 align-top">Deletes by name or index</td></tr>
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
        <pre><code class="language-javascript">// charts.js — charts (graphs) basics: add, title, type, reading back series/axes, delete
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("charts.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// Helper: chart type numeric code -&gt; enum name
const chartTypeName = (code) =&gt;
    Object.keys(enums.XlChartType).find((k) =&gt; enums.XlChartType[k] === code) || String(code);
// Helper: number of charts on the sheet (getChartObject throws out of range -&gt; count until it does)
const countCharts = () =&gt; {
    let n = 0;
    for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
    return n;
};
// Helper: number of series in a chart (getSeriesCollection throws out of range)
const countSeries = (chart) =&gt; {
    let n = 0;
    for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
    return n;
};

// ---------- 1. Build the data table (monthly sales of 2 products) ----------
["Month", "Product A", "Product B"].forEach((h, c) =&gt; ws.getCells(1, c + 1).setValue(h));
const sales = [
    ["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
    ["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
];
sales.forEach((row, i) =&gt; {
    ws.getCells(i + 2, 1).setValue(row[0]);
    ws.getCells(i + 2, 2).setNumberValue(row[1]);
    ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] Created data table A1:C7 / Product A total =",
    sales.reduce((s, r) =&gt; s + r[1], 0));

// ---------- 2. Add a clustered column chart (chartWizard) ----------
// addChartObject(name, left, top, width, height) — position and size are in points (pt)
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
//             styleNo, colorNo, hasLegend, title)
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   // take series by column (Product A / Product B, one series each)
    1,                        // seriesLabelLines: first 1 row is the series-name header
    1,                        // categoryLabelLines: first 1 column is the category-name header
    1,                        // chartStyleNo (1-based)
    enums.XlChartColorPalette.XlPaletteColorful1,
    true,                     // hasLegend: show the legend
    "Monthly Sales"           // title: setting this displays the title
);
console.log("[2] type =", chartTypeName(chart1.getChartType()),
    "/ hasTitle =", chart1.hasTitle(),
    "/ title =", JSON.stringify(chart1.getChartTitle().getText()));
console.log("    data source =", chart1.getDataSource(),
    "/ plotBy =", chart1.getPlotBy(),
    "/ hasLegend =", chart1.hasLegend());

// ---------- 3. Read back series and axes ----------
// There is no API that returns the series name directly. Series names come from the
// chartWizard header row, so count them with getSeriesCollection and match each one
// to its header cell.
const n1 = countSeries(chart1);
console.log("[3] series count =", n1);
for (let i = 1; i &lt;= n1; i++) {
    const s = chart1.getSeriesCollection(i);
    const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // series i -&gt; header of column i+1
    console.log(`    series ${i} "${name}" values =`, JSON.stringify(s.getValues()),
        "/ point count =", s.getPointSize());
}
console.log("    category names =",
    JSON.stringify(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()));

// ---------- 4. Second chart (line with markers) + axis/legend settings ----------
const co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
const chart2 = co2.getChart();
chart2.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeLineMarkers,
    enums.XlRowCol.Columns, 1, 1,
    1, enums.XlChartColorPalette.XlPaletteColorful1, true); // without hasLegend=true, getLegend() throws
// The title can also be set via setTitle(true) -&gt; getChartTitle().setText()
chart2.setTitle(true);
chart2.getChartTitle().setText("Monthly Sales Trend");
// Value axis: fix max/min/major unit (read back with getMajorUnit and the ...IsAuto family)
const valAxis = chart2.getAxes(enums.XlAxisType.Value);
valAxis.setMaximumScale(200);
valAxis.setMinimumScale(0);
valAxis.setMajorUnit(50);
valAxis.setTitle(true);
valAxis.getTitle().setText("Sales (10K JPY)");
// Move the legend to the right
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight);

console.log("[4] type =", chartTypeName(chart2.getChartType()),
    "/ title =", JSON.stringify(chart2.getChartTitle().getText()));
console.log("    value axis major unit =", valAxis.getMajorUnit(),
    "/ max auto =", valAxis.getMaximumScaleIsAuto(),
    "/ min auto =", valAxis.getMinimumScaleIsAuto());
console.log("    value axis title =", JSON.stringify(valAxis.getTitle().getText()),
    "/ legend position =", chart2.getLegend().getPosition(),
    `(Right=${enums.XlLegendPosition.LegendPositionRight})`);

// ---------- 5. Read back chart count, position and size ----------
console.log("[5] chart count =", countCharts(),
    "/ SalesColumn(left,top,width,height) =",
    [co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));

// ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine");   // by name; deleteChartObject(2) by index (1-based) also works
console.log("[Done] Chart count after delete =", countCharts(),
    "/ remaining =", ws.getChartObject(1).getName());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># charts.py — charts (graphs) basics: same verification as the Node.js sample (pyosbxl)
import json
import pyosbxl

enums = pyosbxl.enums

app = pyosbxl.App()
wb = app.createWorkBook("charts-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# Helper: JSON-equivalent string (no separator spaces, non-ASCII kept as-is)
def jdump(v):
    return json.dumps(v, ensure_ascii=False, separators=(",", ":"))

# Helper: enum member name (pyosbxl enums expose their name via __name__)
def enum_name(v):
    return v.__name__

# Helper: number of charts on the sheet (getChartObject throws out of range -&gt; count until it does)
def count_charts():
    n = 0
    while True:
        try:
            ws.getChartObject(n + 1)
            n += 1
        except Exception:
            return n

# Helper: number of series in a chart (getSeriesCollection throws out of range)
def count_series(chart):
    n = 0
    while True:
        try:
            chart.getSeriesCollection(n + 1)
            n += 1
        except Exception:
            return n

# ---------- 1. Build the data table (monthly sales of 2 products) ----------
for c, h in enumerate(["Month", "Product A", "Product B"]):
    ws.getCells(1, c + 1).setValue(h)
sales = [
    ["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
    ["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
]
for i, row in enumerate(sales):
    ws.getCells(i + 2, 1).setValue(row[0])
    ws.getCells(i + 2, 2).setNumberValue(row[1])
    ws.getCells(i + 2, 3).setNumberValue(row[2])
print("[1] Created data table A1:C7 / Product A total =", sum(r[1] for r in sales))

# ---------- 2. Add a clustered column chart (chartWizard) ----------
co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220)
chart1 = co1.getChart()
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   # take series by column (Product A / Product B, one series each)
    1,                        # seriesLabelLines: first 1 row is the series-name header
    1,                        # categoryLabelLines: first 1 column is the category-name header
    1,                        # chartStyleNo (1-based)
    enums.XlChartColorPalette.XlPaletteColorful1,
    True,                     # hasLegend: show the legend
    "Monthly Sales",          # title: setting this displays the title
)
print("[2] type =", enum_name(chart1.getChartType()),
      "/ hasTitle =", str(chart1.hasTitle()).lower(),
      "/ title =", jdump(chart1.getChartTitle().getText()))
print("    data source =", chart1.getDataSource(),
      "/ plotBy =", int(chart1.getPlotBy()),
      "/ hasLegend =", str(chart1.hasLegend()).lower())

# ---------- 3. Read back series and axes ----------
n1 = count_series(chart1)
print("[3] series count =", n1)
for i in range(1, n1 + 1):
    s = chart1.getSeriesCollection(i)
    name = list(ws.getCells(1, i + 1).getValue().values())[0]  # series i -&gt; header of column i+1
    print(f'    series {i} "{name}" values =', jdump(s.getValues()),
          "/ point count =", s.getPointSize())
print("    category names =",
      jdump(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()))

# ---------- 4. Second chart (line with markers) + axis/legend settings ----------
co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220)
chart2 = co2.getChart()
chart2.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeLineMarkers,
    enums.XlRowCol.Columns, 1, 1,
    1, enums.XlChartColorPalette.XlPaletteColorful1, True)  # without hasLegend=True, getLegend() throws
chart2.setTitle(True)
chart2.getChartTitle().setText("Monthly Sales Trend")
val_axis = chart2.getAxes(enums.XlAxisType.Value)
val_axis.setMaximumScale(200)
val_axis.setMinimumScale(0)
val_axis.setMajorUnit(50)
val_axis.setTitle(True)
val_axis.getTitle().setText("Sales (10K JPY)")
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight)

print("[4] type =", enum_name(chart2.getChartType()),
      "/ title =", jdump(chart2.getChartTitle().getText()))
print("    value axis major unit =", int(val_axis.getMajorUnit()),
      "/ max auto =", str(val_axis.getMaximumScaleIsAuto()).lower(),
      "/ min auto =", str(val_axis.getMinimumScaleIsAuto()).lower())
print("    value axis title =", jdump(val_axis.getTitle().getText()),
      "/ legend position =", int(chart2.getLegend().getPosition()),
      f"(Right={int(enums.XlLegendPosition.LegendPositionRight)})")

# ---------- 5. Read back chart count, position and size ----------
print("[5] chart count =", count_charts(),
      "/ SalesColumn(left,top,width,height) =",
      f"{int(co1.getLeft())},{int(co1.getTop())},{int(co1.getWidth())},{int(co1.getHeight())}")

# ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine")   # by name; deleteChartObject(2) by index (1-based) also works
print("[Done] Chart count after delete =", count_charts(),
      "/ remaining =", ws.getChartObject(1).getName())

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// charts.java — charts (graphs) basics: same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.chart.*;
import com.osboffice.osbxl.enums.*;
import java.util.List;
import java.util.Map;

public class Main {
    // Helper: string list -&gt; JSON-equivalent (["a","b"] form, same look as node's JSON.stringify)
    private static String jsonArr(List&lt;String&gt; list) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i &lt; list.size(); i++) {
            if (i &gt; 0) sb.append(",");
            sb.append("\"").append(list.get(i)).append("\"");
        }
        return sb.append("]").toString();
    }

    // Helper: number of charts on the sheet (getChartObject throws out of range -&gt; count until it does)
    private static int countCharts(WorkSheetWrapper ws) {
        int n = 0;
        while (true) {
            try { ws.getChartObject(n + 1); n++; } catch (Throwable t) { break; }
        }
        return n;
    }

    // Helper: number of series in a chart (getSeriesCollection throws out of range)
    private static int countSeries(ChartWrapper chart) {
        int n = 0;
        while (true) {
            try { chart.getSeriesCollection(n + 1); n++; } catch (Throwable t) { break; }
        }
        return n;
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("charts-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. Build the data table (monthly sales of 2 products) ----------
        String[] headers = {"Month", "Product A", "Product B"};
        for (int c = 0; c &lt; headers.length; c++) ws.getCells(1, c + 1).setValue(headers[c]);
        String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
        int[] salesA = {120, 135, 150, 142, 160, 175};
        int[] salesB = {90, 95, 110, 125, 118, 130};
        int totalA = 0;
        for (int i = 0; i &lt; months.length; i++) {
            ws.getCells(i + 2, 1).setValue(months[i]);
            ws.getCells(i + 2, 2).setNumberValue(salesA[i]);
            ws.getCells(i + 2, 3).setNumberValue(salesB[i]);
            totalA += salesA[i];
        }
        System.out.println("[1] Created data table A1:C7 / Product A total = " + totalA);

        // ---------- 2. Add a clustered column chart (chartWizard) ----------
        ChartObjectWrapper co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
        ChartWrapper chart1 = co1.getChart();
        chart1.chartWizard(
            "Sheet1!A1:C7",
            XlChartType.ChartTypeColumnClustered,
            XlRowCol.Columns,   // take series by column (Product A / Product B, one series each)
            1,                  // seriesLabelLines: first 1 row is the series-name header
            1,                  // categoryLabelLines: first 1 column is the category-name header
            1,                  // chartStyleNo (1-based)
            XlChartColorPalette.XlPaletteColorful1,
            true,               // hasLegend: show the legend
            "Monthly Sales");   // title: setting this displays the title
        System.out.println("[2] type = " + chart1.getChartType().name()
            + " / hasTitle = " + chart1.hasTitle()
            + " / title = \"" + chart1.getChartTitle().getText() + "\"");
        System.out.println("    data source = " + chart1.getDataSource()
            + " / plotBy = " + chart1.getPlotBy().getCode()   // enums are numeric via getCode()
            + " / hasLegend = " + chart1.hasLegend());

        // ---------- 3. Read back series and axes ----------
        int n1 = countSeries(chart1);
        System.out.println("[3] series count = " + n1);
        for (int i = 1; i &lt;= n1; i++) {
            SeriesWrapper s = chart1.getSeriesCollection(i);
            Map&lt;String, String&gt; h = ws.getCells(1, i + 1).getValue();
            String name = h.values().iterator().next();       // series i -&gt; header of column i+1
            System.out.println("    series " + i + " \"" + name + "\" values = " + jsonArr(s.getValues())
                + " / point count = " + s.getPointSize());
        }
        System.out.println("    category names = "
            + jsonArr(chart1.getAxes(XlAxisType.Category).getCategoryNames()));

        // ---------- 4. Second chart (line with markers) + axis/legend settings ----------
        ChartObjectWrapper co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
        ChartWrapper chart2 = co2.getChart();
        chart2.chartWizard("Sheet1!A1:C7",
            XlChartType.ChartTypeLineMarkers,
            XlRowCol.Columns, 1, 1,
            1, XlChartColorPalette.XlPaletteColorful1, true); // without hasLegend=true, getLegend() throws
        chart2.setTitle(true);
        chart2.getChartTitle().setText("Monthly Sales Trend");
        AxisWrapper valAxis = chart2.getAxes(XlAxisType.Value);
        valAxis.setMaximumScale(200);
        valAxis.setMinimumScale(0);
        valAxis.setMajorUnit(50);
        valAxis.setTitle(true);
        valAxis.getTitle().setText("Sales (10K JPY)");
        chart2.getLegend().setPosition(XlLegendPosition.LegendPositionRight);

        System.out.println("[4] type = " + chart2.getChartType().name()
            + " / title = \"" + chart2.getChartTitle().getText() + "\"");
        System.out.println("    value axis major unit = " + (int) valAxis.getMajorUnit()   // double -&gt; int for display
            + " / max auto = " + valAxis.getMaximumScaleIsAuto()
            + " / min auto = " + valAxis.getMinimumScaleIsAuto());
        System.out.println("    value axis title = \"" + valAxis.getTitle().getText() + "\""
            + " / legend position = " + chart2.getLegend().getPosition().getCode()
            + " (Right=" + XlLegendPosition.LegendPositionRight.getCode() + ")");

        // ---------- 5. Read back chart count, position and size ----------
        System.out.println("[5] chart count = " + countCharts(ws)
            + " / SalesColumn(left,top,width,height) = "
            + (int) co1.getLeft() + "," + (int) co1.getTop() + ","
            + (int) co1.getWidth() + "," + (int) co1.getHeight());

        // ---------- 6. Delete a chart ----------
        ws.deleteChartObject("SalesLine");   // by name; deleteChartObject(2) by index (1-based) also works
        System.out.println("[Done] Chart count after delete = " + countCharts(ws)
            + " / remaining = " + ws.getChartObject(1).getName());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// charts.php — charts (graphs) basics: same verification as the Node.js sample (phposbxl). Body only (&lt;?php / autoload provided by the harness)

// Output helpers: match the Node.js stdout representation
$bol = fn($v) =&gt; $v ? "true" : "false";                                   // bool -&gt; "true"/"false"
$jstr = fn($s) =&gt; json_encode($s, JSON_UNESCAPED_UNICODE);                 // JSON.stringify equivalent (string)
$jarr = fn($a) =&gt; json_encode(array_values($a), JSON_UNESCAPED_UNICODE);   // JSON.stringify equivalent (array)
// Helper: chart type numeric code -&gt; enum name
$chartTypeName = fn($code) =&gt; \phposbxl\enums\XlChartType::tryFrom($code)?-&gt;name ?? (string)$code;
// Helper: number of charts on the sheet (getChartObject throws out of range -&gt; count until it does)
$countCharts = function () use (&amp;$ws) {
    $n = 0;
    for (;;) {
        try { $ws-&gt;getChartObject($n + 1); $n++; } catch (\Throwable $e) { break; }
    }
    return $n;
};
// Helper: number of series in a chart (getSeriesCollection throws out of range)
$countSeries = function ($chart) {
    $n = 0;
    for (;;) {
        try { $chart-&gt;getSeriesCollection($n + 1); $n++; } catch (\Throwable $e) { break; }
    }
    return $n;
};

$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("charts-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// ---------- 1. Build the data table (monthly sales of 2 products) ----------
foreach (["Month", "Product A", "Product B"] as $c =&gt; $h) {
    $ws-&gt;getCells(1, $c + 1)-&gt;setValue($h);
}
$sales = [
    ["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
    ["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
];
$totalA = 0;
foreach ($sales as $i =&gt; $row) {
    $ws-&gt;getCells($i + 2, 1)-&gt;setValue($row[0]);
    $ws-&gt;getCells($i + 2, 2)-&gt;setNumberValue($row[1]);
    $ws-&gt;getCells($i + 2, 3)-&gt;setNumberValue($row[2]);
    $totalA += $row[1];
}
echo "[1] Created data table A1:C7 / Product A total = {$totalA}\n";

// ---------- 2. Add a clustered column chart (chartWizard) ----------
// addChartObject(name, left, top, width, height) — position and size are in points (pt)
$co1 = $ws-&gt;addChartObject("SalesColumn", 320, 10, 360, 220);
$chart1 = $co1-&gt;getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
//             styleNo, colorNo, hasLegend, title)
$chart1-&gt;chartWizard(
    "Sheet1!A1:C7",
    \phposbxl\enums\XlChartType::ChartTypeColumnClustered,
    \phposbxl\enums\XlRowCol::Columns,   // take series by column (Product A / Product B, one series each)
    1,                                   // seriesLabelLines: first 1 row is the series-name header
    1,                                   // categoryLabelLines: first 1 column is the category-name header
    1,                                   // chartStyleNo (1-based)
    \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1,
    true,                                // hasLegend: show the legend
    "Monthly Sales"                      // title: setting this displays the title
);
echo "[2] type = " . $chartTypeName($chart1-&gt;getChartType())
    . " / hasTitle = " . $bol($chart1-&gt;hasTitle())
    . " / title = " . $jstr($chart1-&gt;getChartTitle()-&gt;getText()) . "\n";
echo "    data source = {$chart1-&gt;getDataSource()}"
    . " / plotBy = {$chart1-&gt;getPlotBy()}"
    . " / hasLegend = " . $bol($chart1-&gt;hasLegend()) . "\n";

// ---------- 3. Read back series and axes ----------
// There is no API that returns the series name directly. Series names come from the
// chartWizard header row, so count them with getSeriesCollection and match each one
// to its header cell.
$n1 = $countSeries($chart1);
echo "[3] series count = {$n1}\n";
for ($i = 1; $i &lt;= $n1; $i++) {
    $s = $chart1-&gt;getSeriesCollection($i);
    $hv = $ws-&gt;getCells(1, $i + 1)-&gt;getValue();   // series i -&gt; header of column i+1
    $name = array_values($hv)[0];
    echo "    series {$i} \"{$name}\" values = " . $jarr($s-&gt;getValues())
        . " / point count = {$s-&gt;getPointSize()}\n";
}
$catAxis = $chart1-&gt;getAxes(\phposbxl\enums\XlAxisType::Category);
echo "    category names = " . $jarr($catAxis-&gt;getCategoryNames()) . "\n";

// ---------- 4. Second chart (line with markers) + axis/legend settings ----------
$co2 = $ws-&gt;addChartObject("SalesLine", 320, 250, 360, 220);
$chart2 = $co2-&gt;getChart();
$chart2-&gt;chartWizard(
    "Sheet1!A1:C7",
    \phposbxl\enums\XlChartType::ChartTypeLineMarkers,
    \phposbxl\enums\XlRowCol::Columns, 1, 1,
    1, \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1, true); // without hasLegend=true, getLegend() throws
// The title can also be set via setTitle(true) -&gt; getChartTitle()-&gt;setText()
$chart2-&gt;setTitle(true);
$chart2-&gt;getChartTitle()-&gt;setText("Monthly Sales Trend");
// Value axis: fix max/min/major unit (read back with getMajorUnit and the ...IsAuto family)
$valAxis = $chart2-&gt;getAxes(\phposbxl\enums\XlAxisType::Value);
$valAxis-&gt;setMaximumScale(200);
$valAxis-&gt;setMinimumScale(0);
$valAxis-&gt;setMajorUnit(50);
$valAxis-&gt;setTitle(true);
$valAxis-&gt;getTitle()-&gt;setText("Sales (10K JPY)");
// Move the legend to the right
$chart2-&gt;getLegend()-&gt;setPosition(\phposbxl\enums\XlLegendPosition::LegendPositionRight);

$right = \phposbxl\enums\XlLegendPosition::LegendPositionRight-&gt;value;
echo "[4] type = " . $chartTypeName($chart2-&gt;getChartType())
    . " / title = " . $jstr($chart2-&gt;getChartTitle()-&gt;getText()) . "\n";
echo "    value axis major unit = {$valAxis-&gt;getMajorUnit()}"
    . " / max auto = " . $bol($valAxis-&gt;getMaximumScaleIsAuto())
    . " / min auto = " . $bol($valAxis-&gt;getMinimumScaleIsAuto()) . "\n";
echo "    value axis title = " . $jstr($valAxis-&gt;getTitle()-&gt;getText())
    . " / legend position = {$chart2-&gt;getLegend()-&gt;getPosition()} (Right={$right})\n";

// ---------- 5. Read back chart count, position and size ----------
echo "[5] chart count = " . $countCharts()
    . " / SalesColumn(left,top,width,height) = "
    . implode(",", [$co1-&gt;getLeft(), $co1-&gt;getTop(), $co1-&gt;getWidth(), $co1-&gt;getHeight()]) . "\n";

// ---------- 6. Delete a chart ----------
$ws-&gt;deleteChartObject("SalesLine");   // by name; deleteChartObject(2) by index (1-based) also works
echo "[Done] Chart count after delete = " . $countCharts()
    . " / remaining = {$ws-&gt;getChartObject(1)-&gt;getName()}\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// charts.wasm.js — charts (graphs) basics: same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("charts-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// Helper: chart type numeric code -&gt; enum name
const chartTypeName = (code) =&gt;
    Object.keys(enums.XlChartType).find((k) =&gt; enums.XlChartType[k] === code) || String(code);
// Helper: number of charts on the sheet (getChartObject throws out of range -&gt; count until it does)
const countCharts = () =&gt; {
    let n = 0;
    for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
    return n;
};
// Helper: number of series in a chart (getSeriesCollection throws out of range)
const countSeries = (chart) =&gt; {
    let n = 0;
    for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
    return n;
};

// ---------- 1. Build the data table (monthly sales of 2 products) ----------
["Month", "Product A", "Product B"].forEach((h, c) =&gt; ws.getCells(1, c + 1).setValue(h));
const sales = [
    ["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
    ["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
];
sales.forEach((row, i) =&gt; {
    ws.getCells(i + 2, 1).setValue(row[0]);
    ws.getCells(i + 2, 2).setNumberValue(row[1]);
    ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] Created data table A1:C7 / Product A total =",
    sales.reduce((s, r) =&gt; s + r[1], 0));

// ---------- 2. Add a clustered column chart (chartWizard) ----------
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   // take series by column (Product A / Product B, one series each)
    1,                        // seriesLabelLines: first 1 row is the series-name header
    1,                        // categoryLabelLines: first 1 column is the category-name header
    1,                        // chartStyleNo (1-based)
    enums.XlChartColorPalette.XlPaletteColorful1,
    true,                     // hasLegend: show the legend
    "Monthly Sales"           // title: setting this displays the title
);
console.log("[2] type =", chartTypeName(chart1.getChartType()),
    "/ hasTitle =", chart1.hasTitle(),
    "/ title =", JSON.stringify(chart1.getChartTitle().getText()));
console.log("    data source =", chart1.getDataSource(),
    "/ plotBy =", chart1.getPlotBy(),
    "/ hasLegend =", chart1.hasLegend());

// ---------- 3. Read back series and axes ----------
const n1 = countSeries(chart1);
console.log("[3] series count =", n1);
for (let i = 1; i &lt;= n1; i++) {
    const s = chart1.getSeriesCollection(i);
    const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // series i -&gt; header of column i+1
    console.log(`    series ${i} "${name}" values =`, JSON.stringify(s.getValues()),
        "/ point count =", s.getPointSize());
}
console.log("    category names =",
    JSON.stringify(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()));

// ---------- 4. Second chart (line with markers) + axis/legend settings ----------
const co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
const chart2 = co2.getChart();
chart2.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeLineMarkers,
    enums.XlRowCol.Columns, 1, 1,
    1, enums.XlChartColorPalette.XlPaletteColorful1, true); // without hasLegend=true, getLegend() throws
chart2.setTitle(true);
chart2.getChartTitle().setText("Monthly Sales Trend");
const valAxis = chart2.getAxes(enums.XlAxisType.Value);
valAxis.setMaximumScale(200);
valAxis.setMinimumScale(0);
valAxis.setMajorUnit(50);
valAxis.setTitle(true);
valAxis.getTitle().setText("Sales (10K JPY)");
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight);

console.log("[4] type =", chartTypeName(chart2.getChartType()),
    "/ title =", JSON.stringify(chart2.getChartTitle().getText()));
console.log("    value axis major unit =", valAxis.getMajorUnit(),
    "/ max auto =", valAxis.getMaximumScaleIsAuto(),
    "/ min auto =", valAxis.getMinimumScaleIsAuto());
console.log("    value axis title =", JSON.stringify(valAxis.getTitle().getText()),
    "/ legend position =", chart2.getLegend().getPosition(),
    `(Right=${enums.XlLegendPosition.LegendPositionRight})`);

// ---------- 5. Read back chart count, position and size ----------
console.log("[5] chart count =", countCharts(),
    "/ SalesColumn(left,top,width,height) =",
    [co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));

// ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine");   // by name; deleteChartObject(2) by index (1-based) also works
console.log("[Done] Chart count after delete =", countCharts(),
    "/ remaining =", ws.getChartObject(1).getName());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] Created data table A1:C7 / Product A total = 882
[2] type = ChartTypeColumnClustered / hasTitle = true / title = "Monthly Sales"
    data source = Sheet1!$A$1:$C$7 / plotBy = 2 / hasLegend = true
[3] series count = 2
    series 1 "Product A" values = ["120","135","150","142","160","175"] / point count = 6
    series 2 "Product B" values = ["90","95","110","125","118","130"] / point count = 6
    category names = ["Jan","Feb","Mar","Apr","May","Jun"]
[4] type = ChartTypeLineMarkers / title = "Monthly Sales Trend"
    value axis major unit = 50 / max auto = false / min auto = false
    value axis title = "Sales (10K JPY)" / legend position = 4 (Right=4)
[5] chart count = 2 / SalesColumn(left,top,width,height) = 320,10,360,220
[Done] Chart count after delete = 1 / remaining = SalesColumn</code></pre>
</div>
<h2 id="notes" class="text-2xl font-bold mt-10 mb-4">Notes</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Creation is a two-step process: build the frame with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject()</code>, then the content with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChart().chartWizard()</code>.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject(name, left, top, width, height)</code> only returns the placement frame (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartObject</code>); at that point the chart is <strong>not built yet</strong>. Only after calling <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> on the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Chart</code> obtained via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartObject.getChart()</code> are the type and data set.<br>Position and size are in <strong>points (pt)</strong>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The argument order of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> differs from Excel macros — be careful.</strong> The signature is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard(source, chartType, plotBy, seriesLabelLines, categoryLabelLines, styleNo?, colorNo?, hasLegend?, title?, ...)</code>. <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">source</code> is an A1C1-style string like <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"Sheet1!A1:C7"</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">seriesLabelLines=1</code> means "the first row holds series names" and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">categoryLabelLines=1</code> means "the first column holds category names".</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">plotBy</code> is "which direction series are taken from".</strong> In a table like the one above — months in rows, products in columns (B/C) — passing <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlRowCol.Columns</code>(=2) gives <strong>one series per column</strong> (Product A, Product B); the sample confirms 2 series.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>The title can be set in two ways.</strong> Either pass a string as the <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">title</code> argument of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> (which makes <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">hasTitle()</code> return <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code>), or call <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.setTitle(true)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getChartTitle().setText("...")</code>. Both can be read back with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartTitle().getText()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Axes are obtained with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAxes(enums.XlAxisType.Category|Value)</code>.</strong> Set the value axis max/min with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMaximumScale()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMinimumScale()</code> and the tick interval with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMajorUnit()</code>. The axis title is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">axis.setTitle(true)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">axis.getTitle().setText()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Series values can be read back with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSeriesCollection(i).getValues()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getXValues()</code> (both string arrays).</strong> There is no API that returns the series count itself, so the sample counts until <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSeriesCollection()</code> throws.</li>
</ul>
<h2 id="caveats" class="text-2xl font-bold mt-10 mb-4">Caveats</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>A chart's appearance (colors, the actual legend rendering, line shapes, etc.) cannot be verified from the console.</strong> This sample confirms correctness by <strong>reading properties back</strong> — type, title string, series values, axis settings, count. <strong>Always open <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">charts.xlsx</code> in Excel to visually check the final appearance.</strong></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setChartType()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataSource()</code> on an unbuilt chart throws.</strong> Trying to set the type or data source on a chart right after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject()</code> (before <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code>) raises a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart is not built</code> exception (it does not segfault). Always build the chart with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> first.<br>Invalid chart type values (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">-2</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0</code>, out of range) are also rejected with a <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartType is invalid</code> exception.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getLegend()</code> only works on charts whose legend is enabled.</strong> On a chart without a legend it throws <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">legend not found</code>. There are two ways to enable it: pass <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard(…, hasLegend=true, …)</code> at build time, or call <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.setLegend(true)</code></strong> afterwards (verified: after <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setLegend(true)</code> both <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getLegend()</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPosition()</code> work; <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setLegend(false)</code> makes it throw again).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>There are no getters for the value axis max/min.</strong> There is no <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumScale()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumScale()</code> counterpart to <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMaximumScale()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMinimumScale()</code>; you can only check "is it automatic" via <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumScaleIsAuto()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumScaleIsAuto()</code> (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">boolean</code>), which return <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code> after manual settings.<br>The major unit can be read back with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMajorUnit()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>There is no API that returns series names (legend entry names).</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Series</code> has no <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> equivalent and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">LegendEntry</code> has no text getter, so the sample <strong>reads the series names from the source table's header cells (B1/C1) and matches them up</strong>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValues()</code> returns numeric cells as numeric strings and string cells as empty strings <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">""</code> (verified).</strong> Chart series values are expected to be numbers, so cells containing only strings come back as <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">""</code>. Feed series data in as numbers with <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue()</code>.</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>There is no API that returns the chart count, and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getShapes().getCount()</code> cannot be used (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">typeof</code> is <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code>).</strong> You must count <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject(1), getChartObject(2), ...</code> until they throw (the sample's <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">countCharts()</code>).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Optional arguments of <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> cannot be skipped in the middle.</strong> If you want to pass <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">title</code> or <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">hasLegend</code>, you must also give values for the preceding <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">styleNo</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colorNo</code> (passing <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">null</code> in a middle position throws).</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Some chart types are unsupported.</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlChartType.ChartTypeRegionMap</code> raises an <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">unsupported</code> exception in <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code>.</li>
</ul>
<h2 id="see-also" class="text-2xl font-bold mt-10 mb-4">See also</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">Index</a></li>
  <li>Previous: <a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">Conditional formatting</a> / Next: <a href="general-11-pivottables.php" class="text-primary-600 hover:underline">Pivot tables</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/en/nodejs/classes/chart.ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/en/nodejs/classes/chart.Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/en/nodejs/classes/chart.ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/en/nodejs/classes/chart.Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/en/nodejs/classes/chart.Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/en/nodejs/classes/chart.Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/en/nodejs/enums/enums.XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/en/nodejs/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/en/nodejs/enums/enums.XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/en/nodejs/enums/enums.XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/en/python/pyosbxl.chart.ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/en/python/pyosbxl.chart.Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/en/python/pyosbxl.chart.ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/en/python/pyosbxl.chart.Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/en/python/pyosbxl.chart.Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/en/python/pyosbxl.chart.Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/en/python/pyosbxl.enums.XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/en/java/com/osboffice/osbxl/chart/ChartObjectWrapper.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/chart/ChartWrapper.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/chart/ChartTitleWrapper.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/chart/AxisWrapper.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/chart/SeriesWrapper.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/chart/LegendWrapper.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/en/java/com/osboffice/osbxl/enums/XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/en/php/classes/phposbxl-chart-ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/en/php/classes/phposbxl-chart-Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/en/php/classes/phposbxl-chart-ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/en/php/classes/phposbxl-chart-Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/en/php/classes/phposbxl-chart-Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/en/php/classes/phposbxl-chart-Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/en/php/classes/phposbxl-enums-XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/en/wasm/classes/chart.ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/en/wasm/classes/chart.Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/en/wasm/classes/chart.ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/en/wasm/classes/chart.Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/en/wasm/classes/chart.Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/en/wasm/classes/chart.Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/en/wasm/enums/enums.XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/en/wasm/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/en/wasm/enums/enums.XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/en/wasm/enums/enums.XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "Charts Basics — add, configure, read back, delete",
   "description": "Charts Basics — add, configure, read back, delete. Execution-verified osbxl sample (verified 2026-10-04). Code tabs: Node.js/Python/Java/PHP/WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/en/sample/general-10-charts.php",
   "mainEntityOfPage": "https://osb-office.com/en/sample/general-10-charts.php",
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
     "name": "Charts Basics — add, configure, read back, delete",
     "item": "https://osb-office.com/en/sample/general-10-charts.php"
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
