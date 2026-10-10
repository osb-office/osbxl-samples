# Charts Basics — add, configure, read back, delete

level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](/ja/general-10-charts.md)

## What you want to do

**Create charts (graphs)** from a data table on a worksheet. Build two kinds — a clustered column chart and a line chart — set their titles, legend and axes, and **read the results back as properties to verify them**. Finally, delete a chart.

A chart's appearance cannot be checked from the console, so this sample prints the **type, title string, series values, axis settings and chart count** with `console.log`, letting you programmatically confirm that everything was built as intended.

## Classes involved

Charts are split into two layers: the "frame" (`ChartObject`) and the "content" (`Chart`).

| Class | How to obtain | Role |
| --- | --- | --- |
| `chart.ChartObject` | `ws.addChartObject(name, left, top, w, h)` / `ws.getChartObject(name\ | index)` | Placement frame on the sheet. Position, size, name, visible/hidden |
| `chart.Chart` | `chartObject.getChart()` | The chart itself. Built with `chartWizard()`; handles type, data source, title, axes and series |
| `chart.ChartTitle` | `chart.getChartTitle()` | Chart title. `setText()` / `getText()` |
| `chart.Axis` | `chart.getAxes(enums.XlAxisType.Category\ | Value)` | Axis. Tick marks, max/min, axis title |
| `chart.Series` | `chart.getSeriesCollection(index)` | One series. `getValues()` / `getXValues()` / `getPointSize()` |
| `chart.Legend` | `chart.getLegend()` | Legend. `setPosition()` / `getPosition()` |

Chart-related methods on the `WorkSheet` side:

| Method | Description |
| --- | --- |
| `addChartObject(name, left, top, width, height)` | Adds a chart frame. **Position and size are in points (pt)** |
| `getChartObject(name)` / `getChartObject(index)` | Gets by name or index (**1-based**). Out of range throws |
| `deleteChartObject(name)` / `deleteChartObject(index)` | Deletes by name or index |

## Code

Output:

#### [Node.js]

```javascript
// charts.js — charts (graphs) basics: add, title, type, reading back series/axes, delete
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("charts.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// Helper: chart type numeric code -> enum name
const chartTypeName = (code) =>
Object.keys(enums.XlChartType).find((k) => enums.XlChartType[k] === code) || String(code);
// Helper: number of charts on the sheet (getChartObject throws out of range -> count until it does)
const countCharts = () => {
let n = 0;
for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
return n;
};
// Helper: number of series in a chart (getSeriesCollection throws out of range)
const countSeries = (chart) => {
let n = 0;
for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
return n;
};
// ---------- 1. Build the data table (monthly sales of 2 products) ----------
["Month", "Product A", "Product B"].forEach((h, c) => ws.getCells(1, c + 1).setValue(h));
const sales = [
["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
];
sales.forEach((row, i) => {
ws.getCells(i + 2, 1).setValue(row[0]);
ws.getCells(i + 2, 2).setNumberValue(row[1]);
ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] Created data table A1:C7 / Product A total =",
sales.reduce((s, r) => s + r[1], 0));
// ---------- 2. Add a clustered column chart (chartWizard) ----------
// addChartObject(name, left, top, width, height) — position and size are in points (pt)
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
// styleNo, colorNo, hasLegend, title)
chart1.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeColumnClustered,
enums.XlRowCol.Columns, // take series by column (Product A / Product B, one series each)
1, // seriesLabelLines: first 1 row is the series-name header
1, // categoryLabelLines: first 1 column is the category-name header
1, // chartStyleNo (1-based)
enums.XlChartColorPalette.XlPaletteColorful1,
true, // hasLegend: show the legend
"Monthly Sales" // title: setting this displays the title
);
console.log("[2] type =", chartTypeName(chart1.getChartType()),
"/ hasTitle =", chart1.hasTitle(),
"/ title =", JSON.stringify(chart1.getChartTitle().getText()));
console.log(" data source =", chart1.getDataSource(),
"/ plotBy =", chart1.getPlotBy(),
"/ hasLegend =", chart1.hasLegend());
// ---------- 3. Read back series and axes ----------
// There is no API that returns the series name directly. Series names come from the
// chartWizard header row, so count them with getSeriesCollection and match each one
// to its header cell.
const n1 = countSeries(chart1);
console.log("[3] series count =", n1);
for (let i = 1; i <= n1; i++) {
const s = chart1.getSeriesCollection(i);
const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // series i -> header of column i+1
console.log(` series ${i} "${name}" values =`, JSON.stringify(s.getValues()),
"/ point count =", s.getPointSize());
}
console.log(" category names =",
JSON.stringify(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()));
// ---------- 4. Second chart (line with markers) + axis/legend settings ----------
const co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
const chart2 = co2.getChart();
chart2.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeLineMarkers,
enums.XlRowCol.Columns, 1, 1,
1, enums.XlChartColorPalette.XlPaletteColorful1, true); // without hasLegend=true, getLegend() throws
// The title can also be set via setTitle(true) -> getChartTitle().setText()
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
console.log(" value axis major unit =", valAxis.getMajorUnit(),
"/ max auto =", valAxis.getMaximumScaleIsAuto(),
"/ min auto =", valAxis.getMinimumScaleIsAuto());
console.log(" value axis title =", JSON.stringify(valAxis.getTitle().getText()),
"/ legend position =", chart2.getLegend().getPosition(),
`(Right=${enums.XlLegendPosition.LegendPositionRight})`);
// ---------- 5. Read back chart count, position and size ----------
console.log("[5] chart count =", countCharts(),
"/ SalesColumn(left,top,width,height) =",
[co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));
// ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine"); // by name; deleteChartObject(2) by index (1-based) also works
console.log("[Done] Chart count after delete =", countCharts(),
"/ remaining =", ws.getChartObject(1).getName());
wb.save();
wb.close();
```

#### [Python]

```python
# charts.py — charts (graphs) basics: same verification as the Node.js sample (pyosbxl)
import json
import pyosbxl
enums = pyosbxl.enums
app = pyosbxl.App()
wb = app.createWorkBook("charts-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# Helper: JSON-equivalent string (no separator spaces, non-ASCII kept as-is)
def jdump(v):
return json.dumps(v, ensure\_ascii=False, separators=(",", ":"))
# Helper: enum member name (pyosbxl enums expose their name via \_\_name\_\_)
def enum\_name(v):
return v.\_\_name\_\_
# Helper: number of charts on the sheet (getChartObject throws out of range -> count until it does)
def count\_charts():
n = 0
while True:
try:
ws.getChartObject(n + 1)
n += 1
except Exception:
return n
# Helper: number of series in a chart (getSeriesCollection throws out of range)
def count\_series(chart):
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
enums.XlRowCol.Columns, # take series by column (Product A / Product B, one series each)
1, # seriesLabelLines: first 1 row is the series-name header
1, # categoryLabelLines: first 1 column is the category-name header
1, # chartStyleNo (1-based)
enums.XlChartColorPalette.XlPaletteColorful1,
True, # hasLegend: show the legend
"Monthly Sales", # title: setting this displays the title
)
print("[2] type =", enum\_name(chart1.getChartType()),
"/ hasTitle =", str(chart1.hasTitle()).lower(),
"/ title =", jdump(chart1.getChartTitle().getText()))
print(" data source =", chart1.getDataSource(),
"/ plotBy =", int(chart1.getPlotBy()),
"/ hasLegend =", str(chart1.hasLegend()).lower())
# ---------- 3. Read back series and axes ----------
n1 = count\_series(chart1)
print("[3] series count =", n1)
for i in range(1, n1 + 1):
s = chart1.getSeriesCollection(i)
name = list(ws.getCells(1, i + 1).getValue().values())[0] # series i -> header of column i+1
print(f' series {i} "{name}" values =', jdump(s.getValues()),
"/ point count =", s.getPointSize())
print(" category names =",
jdump(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()))
# ---------- 4. Second chart (line with markers) + axis/legend settings ----------
co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220)
chart2 = co2.getChart()
chart2.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeLineMarkers,
enums.XlRowCol.Columns, 1, 1,
1, enums.XlChartColorPalette.XlPaletteColorful1, True) # without hasLegend=True, getLegend() throws
chart2.setTitle(True)
chart2.getChartTitle().setText("Monthly Sales Trend")
val\_axis = chart2.getAxes(enums.XlAxisType.Value)
val\_axis.setMaximumScale(200)
val\_axis.setMinimumScale(0)
val\_axis.setMajorUnit(50)
val\_axis.setTitle(True)
val\_axis.getTitle().setText("Sales (10K JPY)")
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight)
print("[4] type =", enum\_name(chart2.getChartType()),
"/ title =", jdump(chart2.getChartTitle().getText()))
print(" value axis major unit =", int(val\_axis.getMajorUnit()),
"/ max auto =", str(val\_axis.getMaximumScaleIsAuto()).lower(),
"/ min auto =", str(val\_axis.getMinimumScaleIsAuto()).lower())
print(" value axis title =", jdump(val\_axis.getTitle().getText()),
"/ legend position =", int(chart2.getLegend().getPosition()),
f"(Right={int(enums.XlLegendPosition.LegendPositionRight)})")
# ---------- 5. Read back chart count, position and size ----------
print("[5] chart count =", count\_charts(),
"/ SalesColumn(left,top,width,height) =",
f"{int(co1.getLeft())},{int(co1.getTop())},{int(co1.getWidth())},{int(co1.getHeight())}")
# ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine") # by name; deleteChartObject(2) by index (1-based) also works
print("[Done] Chart count after delete =", count\_charts(),
"/ remaining =", ws.getChartObject(1).getName())
wb.save()
wb.close()
```

#### [Java]

```java
// charts.java — charts (graphs) basics: same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.chart.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.List;
import java.util.Map;
public class Main {
// Helper: string list -> JSON-equivalent (["a","b"] form, same look as node's JSON.stringify)
private static String jsonArr(List<String> list) {
StringBuilder sb = new StringBuilder("[");
for (int i = 0; i < list.size(); i++) {
if (i > 0) sb.append(",");
sb.append("\"").append(list.get(i)).append("\"");
}
return sb.append("]").toString();
}
// Helper: number of charts on the sheet (getChartObject throws out of range -> count until it does)
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
for (int c = 0; c < headers.length; c++) ws.getCells(1, c + 1).setValue(headers[c]);
String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
int[] salesA = {120, 135, 150, 142, 160, 175};
int[] salesB = {90, 95, 110, 125, 118, 130};
int totalA = 0;
for (int i = 0; i < months.length; i++) {
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
XlRowCol.Columns, // take series by column (Product A / Product B, one series each)
1, // seriesLabelLines: first 1 row is the series-name header
1, // categoryLabelLines: first 1 column is the category-name header
1, // chartStyleNo (1-based)
XlChartColorPalette.XlPaletteColorful1,
true, // hasLegend: show the legend
"Monthly Sales"); // title: setting this displays the title
System.out.println("[2] type = " + chart1.getChartType().name()
+ " / hasTitle = " + chart1.hasTitle()
+ " / title = \"" + chart1.getChartTitle().getText() + "\"");
System.out.println(" data source = " + chart1.getDataSource()
+ " / plotBy = " + chart1.getPlotBy().getCode() // enums are numeric via getCode()
+ " / hasLegend = " + chart1.hasLegend());
// ---------- 3. Read back series and axes ----------
int n1 = countSeries(chart1);
System.out.println("[3] series count = " + n1);
for (int i = 1; i <= n1; i++) {
SeriesWrapper s = chart1.getSeriesCollection(i);
Map<String, String> h = ws.getCells(1, i + 1).getValue();
String name = h.values().iterator().next(); // series i -> header of column i+1
System.out.println(" series " + i + " \"" + name + "\" values = " + jsonArr(s.getValues())
+ " / point count = " + s.getPointSize());
}
System.out.println(" category names = "
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
System.out.println(" value axis major unit = " + (int) valAxis.getMajorUnit() // double -> int for display
+ " / max auto = " + valAxis.getMaximumScaleIsAuto()
+ " / min auto = " + valAxis.getMinimumScaleIsAuto());
System.out.println(" value axis title = \"" + valAxis.getTitle().getText() + "\""
+ " / legend position = " + chart2.getLegend().getPosition().getCode()
+ " (Right=" + XlLegendPosition.LegendPositionRight.getCode() + ")");
// ---------- 5. Read back chart count, position and size ----------
System.out.println("[5] chart count = " + countCharts(ws)
+ " / SalesColumn(left,top,width,height) = "
+ (int) co1.getLeft() + "," + (int) co1.getTop() + ","
+ (int) co1.getWidth() + "," + (int) co1.getHeight());
// ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine"); // by name; deleteChartObject(2) by index (1-based) also works
System.out.println("[Done] Chart count after delete = " + countCharts(ws)
+ " / remaining = " + ws.getChartObject(1).getName());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// charts.md — charts (graphs) basics: same verification as the Node.js sample (phposbxl). Body only (<?php / autoload provided by the harness)
// Output helpers: match the Node.js stdout representation
$bol = fn($v) => $v ? "true" : "false"; // bool -> "true"/"false"
$jstr = fn($s) => json\_encode($s, JSON\_UNESCAPED\_UNICODE); // JSON.stringify equivalent (string)
$jarr = fn($a) => json\_encode(array\_values($a), JSON\_UNESCAPED\_UNICODE); // JSON.stringify equivalent (array)
// Helper: chart type numeric code -> enum name
$chartTypeName = fn($code) => \phposbxl\enums\XlChartType::tryFrom($code)?->name ?? (string)$code;
// Helper: number of charts on the sheet (getChartObject throws out of range -> count until it does)
$countCharts = function () use (&$ws) {
$n = 0;
for (;;) {
try { $ws->getChartObject($n + 1); $n++; } catch (\Throwable $e) { break; }
}
return $n;
};
// Helper: number of series in a chart (getSeriesCollection throws out of range)
$countSeries = function ($chart) {
$n = 0;
for (;;) {
try { $chart->getSeriesCollection($n + 1); $n++; } catch (\Throwable $e) { break; }
}
return $n;
};
$app = new \phposbxl\App();
$wb = $app->createWorkBook("charts-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// ---------- 1. Build the data table (monthly sales of 2 products) ----------
foreach (["Month", "Product A", "Product B"] as $c => $h) {
$ws->getCells(1, $c + 1)->setValue($h);
}
$sales = [
["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
];
$totalA = 0;
foreach ($sales as $i => $row) {
$ws->getCells($i + 2, 1)->setValue($row[0]);
$ws->getCells($i + 2, 2)->setNumberValue($row[1]);
$ws->getCells($i + 2, 3)->setNumberValue($row[2]);
$totalA += $row[1];
}
echo "[1] Created data table A1:C7 / Product A total = {$totalA}\n";
// ---------- 2. Add a clustered column chart (chartWizard) ----------
// addChartObject(name, left, top, width, height) — position and size are in points (pt)
$co1 = $ws->addChartObject("SalesColumn", 320, 10, 360, 220);
$chart1 = $co1->getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
// styleNo, colorNo, hasLegend, title)
$chart1->chartWizard(
"Sheet1!A1:C7",
\phposbxl\enums\XlChartType::ChartTypeColumnClustered,
\phposbxl\enums\XlRowCol::Columns, // take series by column (Product A / Product B, one series each)
1, // seriesLabelLines: first 1 row is the series-name header
1, // categoryLabelLines: first 1 column is the category-name header
1, // chartStyleNo (1-based)
\phposbxl\enums\XlChartColorPalette::XlPaletteColorful1,
true, // hasLegend: show the legend
"Monthly Sales" // title: setting this displays the title
);
echo "[2] type = " . $chartTypeName($chart1->getChartType())
. " / hasTitle = " . $bol($chart1->hasTitle())
. " / title = " . $jstr($chart1->getChartTitle()->getText()) . "\n";
echo " data source = {$chart1->getDataSource()}"
. " / plotBy = {$chart1->getPlotBy()}"
. " / hasLegend = " . $bol($chart1->hasLegend()) . "\n";
// ---------- 3. Read back series and axes ----------
// There is no API that returns the series name directly. Series names come from the
// chartWizard header row, so count them with getSeriesCollection and match each one
// to its header cell.
$n1 = $countSeries($chart1);
echo "[3] series count = {$n1}\n";
for ($i = 1; $i <= $n1; $i++) {
$s = $chart1->getSeriesCollection($i);
$hv = $ws->getCells(1, $i + 1)->getValue(); // series i -> header of column i+1
$name = array\_values($hv)[0];
echo " series {$i} \"{$name}\" values = " . $jarr($s->getValues())
. " / point count = {$s->getPointSize()}\n";
}
$catAxis = $chart1->getAxes(\phposbxl\enums\XlAxisType::Category);
echo " category names = " . $jarr($catAxis->getCategoryNames()) . "\n";
// ---------- 4. Second chart (line with markers) + axis/legend settings ----------
$co2 = $ws->addChartObject("SalesLine", 320, 250, 360, 220);
$chart2 = $co2->getChart();
$chart2->chartWizard(
"Sheet1!A1:C7",
\phposbxl\enums\XlChartType::ChartTypeLineMarkers,
\phposbxl\enums\XlRowCol::Columns, 1, 1,
1, \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1, true); // without hasLegend=true, getLegend() throws
// The title can also be set via setTitle(true) -> getChartTitle()->setText()
$chart2->setTitle(true);
$chart2->getChartTitle()->setText("Monthly Sales Trend");
// Value axis: fix max/min/major unit (read back with getMajorUnit and the ...IsAuto family)
$valAxis = $chart2->getAxes(\phposbxl\enums\XlAxisType::Value);
$valAxis->setMaximumScale(200);
$valAxis->setMinimumScale(0);
$valAxis->setMajorUnit(50);
$valAxis->setTitle(true);
$valAxis->getTitle()->setText("Sales (10K JPY)");
// Move the legend to the right
$chart2->getLegend()->setPosition(\phposbxl\enums\XlLegendPosition::LegendPositionRight);
$right = \phposbxl\enums\XlLegendPosition::LegendPositionRight->value;
echo "[4] type = " . $chartTypeName($chart2->getChartType())
. " / title = " . $jstr($chart2->getChartTitle()->getText()) . "\n";
echo " value axis major unit = {$valAxis->getMajorUnit()}"
. " / max auto = " . $bol($valAxis->getMaximumScaleIsAuto())
. " / min auto = " . $bol($valAxis->getMinimumScaleIsAuto()) . "\n";
echo " value axis title = " . $jstr($valAxis->getTitle()->getText())
. " / legend position = {$chart2->getLegend()->getPosition()} (Right={$right})\n";
// ---------- 5. Read back chart count, position and size ----------
echo "[5] chart count = " . $countCharts()
. " / SalesColumn(left,top,width,height) = "
. implode(",", [$co1->getLeft(), $co1->getTop(), $co1->getWidth(), $co1->getHeight()]) . "\n";
// ---------- 6. Delete a chart ----------
$ws->deleteChartObject("SalesLine"); // by name; deleteChartObject(2) by index (1-based) also works
echo "[Done] Chart count after delete = " . $countCharts()
. " / remaining = {$ws->getChartObject(1)->getName()}\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// charts.wasm.js — charts (graphs) basics: same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("charts-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// Helper: chart type numeric code -> enum name
const chartTypeName = (code) =>
Object.keys(enums.XlChartType).find((k) => enums.XlChartType[k] === code) || String(code);
// Helper: number of charts on the sheet (getChartObject throws out of range -> count until it does)
const countCharts = () => {
let n = 0;
for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
return n;
};
// Helper: number of series in a chart (getSeriesCollection throws out of range)
const countSeries = (chart) => {
let n = 0;
for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
return n;
};
// ---------- 1. Build the data table (monthly sales of 2 products) ----------
["Month", "Product A", "Product B"].forEach((h, c) => ws.getCells(1, c + 1).setValue(h));
const sales = [
["Jan", 120, 90], ["Feb", 135, 95], ["Mar", 150, 110],
["Apr", 142, 125], ["May", 160, 118], ["Jun", 175, 130],
];
sales.forEach((row, i) => {
ws.getCells(i + 2, 1).setValue(row[0]);
ws.getCells(i + 2, 2).setNumberValue(row[1]);
ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] Created data table A1:C7 / Product A total =",
sales.reduce((s, r) => s + r[1], 0));
// ---------- 2. Add a clustered column chart (chartWizard) ----------
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
chart1.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeColumnClustered,
enums.XlRowCol.Columns, // take series by column (Product A / Product B, one series each)
1, // seriesLabelLines: first 1 row is the series-name header
1, // categoryLabelLines: first 1 column is the category-name header
1, // chartStyleNo (1-based)
enums.XlChartColorPalette.XlPaletteColorful1,
true, // hasLegend: show the legend
"Monthly Sales" // title: setting this displays the title
);
console.log("[2] type =", chartTypeName(chart1.getChartType()),
"/ hasTitle =", chart1.hasTitle(),
"/ title =", JSON.stringify(chart1.getChartTitle().getText()));
console.log(" data source =", chart1.getDataSource(),
"/ plotBy =", chart1.getPlotBy(),
"/ hasLegend =", chart1.hasLegend());
// ---------- 3. Read back series and axes ----------
const n1 = countSeries(chart1);
console.log("[3] series count =", n1);
for (let i = 1; i <= n1; i++) {
const s = chart1.getSeriesCollection(i);
const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // series i -> header of column i+1
console.log(` series ${i} "${name}" values =`, JSON.stringify(s.getValues()),
"/ point count =", s.getPointSize());
}
console.log(" category names =",
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
console.log(" value axis major unit =", valAxis.getMajorUnit(),
"/ max auto =", valAxis.getMaximumScaleIsAuto(),
"/ min auto =", valAxis.getMinimumScaleIsAuto());
console.log(" value axis title =", JSON.stringify(valAxis.getTitle().getText()),
"/ legend position =", chart2.getLegend().getPosition(),
`(Right=${enums.XlLegendPosition.LegendPositionRight})`);
// ---------- 5. Read back chart count, position and size ----------
console.log("[5] chart count =", countCharts(),
"/ SalesColumn(left,top,width,height) =",
[co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));
// ---------- 6. Delete a chart ----------
ws.deleteChartObject("SalesLine"); // by name; deleteChartObject(2) by index (1-based) also works
console.log("[Done] Chart count after delete =", countCharts(),
"/ remaining =", ws.getChartObject(1).getName());
wb.save();
wb.close();
```

```plaintext
[1] Created data table A1:C7 / Product A total = 882
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
[Done] Chart count after delete = 1 / remaining = SalesColumn
```

## Notes

* **Creation is a two-step process: build the frame with `addChartObject()`, then the content with `getChart().chartWizard()`.** `addChartObject(name, left, top, width, height)` only returns the placement frame (`ChartObject`); at that point the chart is **not built yet**. Only after calling `chartWizard()` on the `Chart` obtained via `chartObject.getChart()` are the type and data set.  
  Position and size are in **points (pt)**.
* **The argument order of `chartWizard()` differs from Excel macros — be careful.** The signature is `chartWizard(source, chartType, plotBy, seriesLabelLines, categoryLabelLines, styleNo?, colorNo?, hasLegend?, title?, ...)`. `source` is an A1C1-style string like `"Sheet1!A1:C7"`, `seriesLabelLines=1` means "the first row holds series names" and `categoryLabelLines=1` means "the first column holds category names".
* **`plotBy` is "which direction series are taken from".** In a table like the one above — months in rows, products in columns (B/C) — passing `enums.XlRowCol.Columns`(=2) gives **one series per column** (Product A, Product B); the sample confirms 2 series.
* **The title can be set in two ways.** Either pass a string as the `title` argument of `chartWizard()` (which makes `hasTitle()` return `true`), or call `chart.setTitle(true)` → `chart.getChartTitle().setText("...")`. Both can be read back with `getChartTitle().getText()`.
* **Axes are obtained with `getAxes(enums.XlAxisType.Category|Value)`.** Set the value axis max/min with `setMaximumScale()` / `setMinimumScale()` and the tick interval with `setMajorUnit()`. The axis title is `axis.setTitle(true)` → `axis.getTitle().setText()`.
* **Series values can be read back with `getSeriesCollection(i).getValues()` / `getXValues()` (both string arrays).** There is no API that returns the series count itself, so the sample counts until `getSeriesCollection()` throws.

## Caveats

* **A chart's appearance (colors, the actual legend rendering, line shapes, etc.) cannot be verified from the console.** This sample confirms correctness by **reading properties back** — type, title string, series values, axis settings, count. **Always open `charts.xlsx` in Excel to visually check the final appearance.**
* **`setChartType()` / `setDataSource()` on an unbuilt chart throws.** Trying to set the type or data source on a chart right after `addChartObject()` (before `chartWizard()`) raises a `chart is not built` exception (it does not segfault). Always build the chart with `chartWizard()` first.  
  Invalid chart type values (`-2`, `0`, out of range) are also rejected with a `chartType is invalid` exception.
* **`getLegend()` only works on charts whose legend is enabled.** On a chart without a legend it throws `legend not found`. There are two ways to enable it: pass `chartWizard(…, hasLegend=true, …)` at build time, or call **`chart.setLegend(true)`** afterwards (verified: after `setLegend(true)` both `getLegend()` and `setPosition()` work; `setLegend(false)` makes it throw again).
* **There are no getters for the value axis max/min.** There is no `getMaximumScale()` / `getMinimumScale()` counterpart to `setMaximumScale()` / `setMinimumScale()`; you can only check "is it automatic" via `getMaximumScaleIsAuto()` / `getMinimumScaleIsAuto()` (`boolean`), which return `false` after manual settings.  
  The major unit can be read back with `getMajorUnit()`.
* **There is no API that returns series names (legend entry names).** `Series` has no `getName()` equivalent and `LegendEntry` has no text getter, so the sample **reads the series names from the source table's header cells (B1/C1) and matches them up**.
* **`getValues()` returns numeric cells as numeric strings and string cells as empty strings `""` (verified).** Chart series values are expected to be numbers, so cells containing only strings come back as `""`. Feed series data in as numbers with `setNumberValue()`.
* **There is no API that returns the chart count, and `ws.getShapes().getCount()` cannot be used (`typeof` is `undefined`).** You must count `getChartObject(1), getChartObject(2), ...` until they throw (the sample's `countCharts()`).
* **Optional arguments of `chartWizard()` cannot be skipped in the middle.** If you want to pass `title` or `hasLegend`, you must also give values for the preceding `styleNo` / `colorNo` (passing `undefined` / `null` in a middle position throws).
* **Some chart types are unsupported.** `enums.XlChartType.ChartTypeRegionMap` raises an `unsupported` exception in `chartWizard()`.

## See also

* [Index](index.md)
* Previous: [Conditional formatting](general-09-conditional-format.md) / Next: [Pivot tables](general-11-pivottables.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
