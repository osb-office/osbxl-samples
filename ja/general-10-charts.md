# チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除

level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-10-charts.md)

## やりたいこと

ワークシートのデータ表から**チャート（グラフ）を作成**する。 縦棒グラフと折れ線グラフの 2 種類を作り、タイトル・凡例・軸を設定し、 **作成した内容をプロパティとして読み戻して検証**する。最後にチャートを削除する。

チャートは見た目をコンソールから確認できないため、このサンプルでは **型・タイトル文字列・系列の値・軸の設定値・チャート個数**を `console.log` で出力し、 意図どおりに作れたかをプログラムで確かめられるようにしている。

## 対象クラス

チャートは「枠（`ChartObject`）」と「中身（`Chart`）」の 2 層に分かれる。

| クラス | 取得方法 | 役割 |
| --- | --- | --- |
| `chart.ChartObject` | `ws.addChartObject(name, left, top, w, h)` / `ws.getChartObject(name\ | index)` | シート上の配置枠。位置・サイズ・名前・表示/非表示 |
| `chart.Chart` | `chartObject.getChart()` | チャート本体。`chartWizard()` で構築し、型・データ源・タイトル・軸・系列を扱う |
| `chart.ChartTitle` | `chart.getChartTitle()` | チャート标题。`setText()` / `getText()` |
| `chart.Axis` | `chart.getAxes(enums.XlAxisType.Category\ | Value)` | 軸。目盛り・最大/最小・軸タイトル |
| `chart.Series` | `chart.getSeriesCollection(index)` | 1 系列。`getValues()` / `getXValues()` / `getPointSize()` |
| `chart.Legend` | `chart.getLegend()` | 凡例。`setPosition()` / `getPosition()` |

`WorkSheet` 側のチャート操作メソッド:

| メソッド | 説明 |
| --- | --- |
| `addChartObject(name, left, top, width, height)` | チャート枠を追加。**位置・サイズはポイント(pt)単位** |
| `getChartObject(name)` / `getChartObject(index)` | 名前 or 索引（**1 始まり**）で取得。範囲外は例外 |
| `deleteChartObject(name)` / `deleteChartObject(index)` | 名前 or 索引で削除 |

## コード

実行結果:

#### [Node.js]

```javascript
// charts.js — チャート（グラフ）の基本: 追加・タイトル・型・系列/軸の読み戻し・削除
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("charts.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// 補助: チャートタイプ数値 → enum 名
const chartTypeName = (code) =>
Object.keys(enums.XlChartType).find((k) => enums.XlChartType[k] === code) || String(code);
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
const countCharts = () => {
let n = 0;
for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
return n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
const countSeries = (chart) => {
let n = 0;
for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
return n;
};
// ---------- 1. データ表（2商品の月次売上）を作る ----------
["月", "商品A", "商品B"].forEach((h, c) => ws.getCells(1, c + 1).setValue(h));
const sales = [
["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
sales.forEach((row, i) => {
ws.getCells(i + 2, 1).setValue(row[0]);
ws.getCells(i + 2, 2).setNumberValue(row[1]);
ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] データ表 A1:C7 を作成 / 商品A合計 =",
sales.reduce((s, r) => s + r[1], 0));
// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
// addChartObject(名前, left, top, width, height) — 位置・サイズはポイント(pt)単位
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
// styleNo, colorNo, hasLegend, title)
chart1.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeColumnClustered,
enums.XlRowCol.Columns, // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
1, // seriesLabelLines: 先頭 1 行が系列名ヘッダ
1, // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
1, // chartStyleNo（1 始まり）
enums.XlChartColorPalette.XlPaletteColorful1,
true, // hasLegend: 凡例を表示
"月次売上" // title: 設定するとタイトルが表示される
);
console.log("[2] 型 =", chartTypeName(chart1.getChartType()),
"/ hasTitle =", chart1.hasTitle(),
"/ タイトル =", JSON.stringify(chart1.getChartTitle().getText()));
console.log(" データ源 =", chart1.getDataSource(),
"/ plotBy =", chart1.getPlotBy(),
"/ hasLegend =", chart1.hasLegend());
// ---------- 3. 系列・軸の読み戻し ----------
// 系列名を直接返す API は無い。系列名は chartWizard のヘッダ行なので、
// 個数は getSeriesCollection で数え、名前はヘッダセルから対応付けて確認する。
const n1 = countSeries(chart1);
console.log("[3] 系列数 =", n1);
for (let i = 1; i <= n1; i++) {
const s = chart1.getSeriesCollection(i);
const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // 系列 i → 列 i+1 のヘッダ
console.log(` 系列${i}「${name}」値 =`, JSON.stringify(s.getValues()),
"/ ポイント数 =", s.getPointSize());
}
console.log(" カテゴリ名 =",
JSON.stringify(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()));
// ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
const co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
const chart2 = co2.getChart();
chart2.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeLineMarkers,
enums.XlRowCol.Columns, 1, 1,
1, enums.XlChartColorPalette.XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
// タイトルは setTitle(true) → getChartTitle().setText() でも設定できる
chart2.setTitle(true);
chart2.getChartTitle().setText("月次売上推移");
// 値軸: 最大/最小/主単位を固定（読み戻しは getMajorUnit と ...IsAuto 系）
const valAxis = chart2.getAxes(enums.XlAxisType.Value);
valAxis.setMaximumScale(200);
valAxis.setMinimumScale(0);
valAxis.setMajorUnit(50);
valAxis.setTitle(true);
valAxis.getTitle().setText("売上(万円)");
// 凡例を右へ
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight);
console.log("[4] 型 =", chartTypeName(chart2.getChartType()),
"/ タイトル =", JSON.stringify(chart2.getChartTitle().getText()));
console.log(" 値軸 主単位 =", valAxis.getMajorUnit(),
"/ 最大自動 =", valAxis.getMaximumScaleIsAuto(),
"/ 最小自動 =", valAxis.getMinimumScaleIsAuto());
console.log(" 値軸タイトル =", JSON.stringify(valAxis.getTitle().getText()),
"/ 凡例位置 =", chart2.getLegend().getPosition(),
`(Right=${enums.XlLegendPosition.LegendPositionRight})`);
// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
console.log("[5] チャート個数 =", countCharts(),
"/ SalesColumn(left,top,width,height) =",
[co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));
// ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine"); // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
console.log("[完了] 削除後のチャート個数 =", countCharts(),
"/ 残 =", ws.getChartObject(1).getName());
wb.save();
wb.close();
```

#### [Python]

```python
# charts.py — チャート（グラフ）の基本: node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl
enums = pyosbxl.enums
app = pyosbxl.App()
wb = app.createWorkBook("charts-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# 補助: JSON 相当の文字列（区切りスペース無し・非 ASCII はそのまま）
def jdump(v):
return json.dumps(v, ensure\_ascii=False, separators=(",", ":"))
# 補助: enum のメンバー名（pyosbxl の enum は \_\_name\_\_ で名前が取れる）
def enum\_name(v):
return v.\_\_name\_\_
# 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
def count\_charts():
n = 0
while True:
try:
ws.getChartObject(n + 1)
n += 1
except Exception:
return n
# 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
def count\_series(chart):
n = 0
while True:
try:
chart.getSeriesCollection(n + 1)
n += 1
except Exception:
return n
# ---------- 1. データ表（2商品の月次売上）を作る ----------
for c, h in enumerate(["月", "商品A", "商品B"]):
ws.getCells(1, c + 1).setValue(h)
sales = [
["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
]
for i, row in enumerate(sales):
ws.getCells(i + 2, 1).setValue(row[0])
ws.getCells(i + 2, 2).setNumberValue(row[1])
ws.getCells(i + 2, 3).setNumberValue(row[2])
print("[1] データ表 A1:C7 を作成 / 商品A合計 =", sum(r[1] for r in sales))
# ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220)
chart1 = co1.getChart()
chart1.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeColumnClustered,
enums.XlRowCol.Columns, # 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
1, # seriesLabelLines: 先頭 1 行が系列名ヘッダ
1, # categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
1, # chartStyleNo（1 始まり）
enums.XlChartColorPalette.XlPaletteColorful1,
True, # hasLegend: 凡例を表示
"月次売上", # title: 設定するとタイトルが表示される
)
print("[2] 型 =", enum\_name(chart1.getChartType()),
"/ hasTitle =", str(chart1.hasTitle()).lower(),
"/ タイトル =", jdump(chart1.getChartTitle().getText()))
print(" データ源 =", chart1.getDataSource(),
"/ plotBy =", int(chart1.getPlotBy()),
"/ hasLegend =", str(chart1.hasLegend()).lower())
# ---------- 3. 系列・軸の読み戻し ----------
n1 = count\_series(chart1)
print("[3] 系列数 =", n1)
for i in range(1, n1 + 1):
s = chart1.getSeriesCollection(i)
name = list(ws.getCells(1, i + 1).getValue().values())[0] # 系列 i → 列 i+1 のヘッダ
print(f" 系列{i}「{name}」値 =", jdump(s.getValues()),
"/ ポイント数 =", s.getPointSize())
print(" カテゴリ名 =",
jdump(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()))
# ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220)
chart2 = co2.getChart()
chart2.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeLineMarkers,
enums.XlRowCol.Columns, 1, 1,
1, enums.XlChartColorPalette.XlPaletteColorful1, True) # hasLegend=True にしないと getLegend() が例外
chart2.setTitle(True)
chart2.getChartTitle().setText("月次売上推移")
val\_axis = chart2.getAxes(enums.XlAxisType.Value)
val\_axis.setMaximumScale(200)
val\_axis.setMinimumScale(0)
val\_axis.setMajorUnit(50)
val\_axis.setTitle(True)
val\_axis.getTitle().setText("売上(万円)")
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight)
print("[4] 型 =", enum\_name(chart2.getChartType()),
"/ タイトル =", jdump(chart2.getChartTitle().getText()))
print(" 値軸 主単位 =", int(val\_axis.getMajorUnit()),
"/ 最大自動 =", str(val\_axis.getMaximumScaleIsAuto()).lower(),
"/ 最小自動 =", str(val\_axis.getMinimumScaleIsAuto()).lower())
print(" 値軸タイトル =", jdump(val\_axis.getTitle().getText()),
"/ 凡例位置 =", int(chart2.getLegend().getPosition()),
f"(Right={int(enums.XlLegendPosition.LegendPositionRight)})")
# ---------- 5. チャート個数・位置・サイズの読み戻し ----------
print("[5] チャート個数 =", count\_charts(),
"/ SalesColumn(left,top,width,height) =",
f"{int(co1.getLeft())},{int(co1.getTop())},{int(co1.getWidth())},{int(co1.getHeight())}")
# ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine") # 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
print("[完了] 削除後のチャート個数 =", count\_charts(),
"/ 残 =", ws.getChartObject(1).getName())
wb.save()
wb.close()
```

#### [Java]

```java
// charts.java — チャート（グラフ）の基本: node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.chart.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.List;
import java.util.Map;
public class Main {
// 補助: 文字列リスト → JSON 相当（["a","b"] 形式。node の JSON.stringify と同じ見た目にする）
private static String jsonArr(List<String> list) {
StringBuilder sb = new StringBuilder("[");
for (int i = 0; i < list.size(); i++) {
if (i > 0) sb.append(",");
sb.append("\"").append(list.get(i)).append("\"");
}
return sb.append("]").toString();
}
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
private static int countCharts(WorkSheetWrapper ws) {
int n = 0;
while (true) {
try { ws.getChartObject(n + 1); n++; } catch (Throwable t) { break; }
}
return n;
}
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
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
// ---------- 1. データ表（2商品の月次売上）を作る ----------
String[] headers = {"月", "商品A", "商品B"};
for (int c = 0; c < headers.length; c++) ws.getCells(1, c + 1).setValue(headers[c]);
String[] months = {"1月", "2月", "3月", "4月", "5月", "6月"};
int[] salesA = {120, 135, 150, 142, 160, 175};
int[] salesB = {90, 95, 110, 125, 118, 130};
int totalA = 0;
for (int i = 0; i < months.length; i++) {
ws.getCells(i + 2, 1).setValue(months[i]);
ws.getCells(i + 2, 2).setNumberValue(salesA[i]);
ws.getCells(i + 2, 3).setNumberValue(salesB[i]);
totalA += salesA[i];
}
System.out.println("[1] データ表 A1:C7 を作成 / 商品A合計 = " + totalA);
// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
ChartObjectWrapper co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
ChartWrapper chart1 = co1.getChart();
chart1.chartWizard(
"Sheet1!A1:C7",
XlChartType.ChartTypeColumnClustered,
XlRowCol.Columns, // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
1, // seriesLabelLines: 先頭 1 行が系列名ヘッダ
1, // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
1, // chartStyleNo（1 始まり）
XlChartColorPalette.XlPaletteColorful1,
true, // hasLegend: 凡例を表示
"月次売上"); // title: 設定するとタイトルが表示される
System.out.println("[2] 型 = " + chart1.getChartType().name()
+ " / hasTitle = " + chart1.hasTitle()
+ " / タイトル = \"" + chart1.getChartTitle().getText() + "\"");
System.out.println(" データ源 = " + chart1.getDataSource()
+ " / plotBy = " + chart1.getPlotBy().getCode() // enum は getCode() で数値にする
+ " / hasLegend = " + chart1.hasLegend());
// ---------- 3. 系列・軸の読み戻し ----------
int n1 = countSeries(chart1);
System.out.println("[3] 系列数 = " + n1);
for (int i = 1; i <= n1; i++) {
SeriesWrapper s = chart1.getSeriesCollection(i);
Map<String, String> h = ws.getCells(1, i + 1).getValue();
String name = h.values().iterator().next(); // 系列 i → 列 i+1 のヘッダ
System.out.println(" 系列" + i + "「" + name + "」値 = " + jsonArr(s.getValues())
+ " / ポイント数 = " + s.getPointSize());
}
System.out.println(" カテゴリ名 = "
+ jsonArr(chart1.getAxes(XlAxisType.Category).getCategoryNames()));
// ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
ChartObjectWrapper co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
ChartWrapper chart2 = co2.getChart();
chart2.chartWizard("Sheet1!A1:C7",
XlChartType.ChartTypeLineMarkers,
XlRowCol.Columns, 1, 1,
1, XlChartColorPalette.XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
chart2.setTitle(true);
chart2.getChartTitle().setText("月次売上推移");
AxisWrapper valAxis = chart2.getAxes(XlAxisType.Value);
valAxis.setMaximumScale(200);
valAxis.setMinimumScale(0);
valAxis.setMajorUnit(50);
valAxis.setTitle(true);
valAxis.getTitle().setText("売上(万円)");
chart2.getLegend().setPosition(XlLegendPosition.LegendPositionRight);
System.out.println("[4] 型 = " + chart2.getChartType().name()
+ " / タイトル = \"" + chart2.getChartTitle().getText() + "\"");
System.out.println(" 値軸 主単位 = " + (int) valAxis.getMajorUnit() // double → int で表示
+ " / 最大自動 = " + valAxis.getMaximumScaleIsAuto()
+ " / 最小自動 = " + valAxis.getMinimumScaleIsAuto());
System.out.println(" 値軸タイトル = \"" + valAxis.getTitle().getText() + "\""
+ " / 凡例位置 = " + chart2.getLegend().getPosition().getCode()
+ " (Right=" + XlLegendPosition.LegendPositionRight.getCode() + ")");
// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
System.out.println("[5] チャート個数 = " + countCharts(ws)
+ " / SalesColumn(left,top,width,height) = "
+ (int) co1.getLeft() + "," + (int) co1.getTop() + ","
+ (int) co1.getWidth() + "," + (int) co1.getHeight());
// ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine"); // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
System.out.println("[完了] 削除後のチャート個数 = " + countCharts(ws)
+ " / 残 = " + ws.getChartObject(1).getName());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// charts.md — チャート（グラフ）の基本: node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
// 出力ヘルパー: node 版 stdout と同じ表記にする
$bol = fn($v) => $v ? "true" : "false"; // bool → "true"/"false"
$jstr = fn($s) => json\_encode($s, JSON\_UNESCAPED\_UNICODE); // JSON.stringify 相当（文字列）
$jarr = fn($a) => json\_encode(array\_values($a), JSON\_UNESCAPED\_UNICODE); // JSON.stringify 相当（配列）
// 補助: チャートタイプ数値 → enum 名
$chartTypeName = fn($code) => \phposbxl\enums\XlChartType::tryFrom($code)?->name ?? (string)$code;
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
$countCharts = function () use (&$ws) {
$n = 0;
for (;;) {
try { $ws->getChartObject($n + 1); $n++; } catch (\Throwable $e) { break; }
}
return $n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
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
// ---------- 1. データ表（2商品の月次売上）を作る ----------
foreach (["月", "商品A", "商品B"] as $c => $h) {
$ws->getCells(1, $c + 1)->setValue($h);
}
$sales = [
["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
$totalA = 0;
foreach ($sales as $i => $row) {
$ws->getCells($i + 2, 1)->setValue($row[0]);
$ws->getCells($i + 2, 2)->setNumberValue($row[1]);
$ws->getCells($i + 2, 3)->setNumberValue($row[2]);
$totalA += $row[1];
}
echo "[1] データ表 A1:C7 を作成 / 商品A合計 = {$totalA}\n";
// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
// addChartObject(名前, left, top, width, height) — 位置・サイズはポイント(pt)単位
$co1 = $ws->addChartObject("SalesColumn", 320, 10, 360, 220);
$chart1 = $co1->getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
// styleNo, colorNo, hasLegend, title)
$chart1->chartWizard(
"Sheet1!A1:C7",
\phposbxl\enums\XlChartType::ChartTypeColumnClustered,
\phposbxl\enums\XlRowCol::Columns, // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
1, // seriesLabelLines: 先頭 1 行が系列名ヘッダ
1, // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
1, // chartStyleNo（1 始まり）
\phposbxl\enums\XlChartColorPalette::XlPaletteColorful1,
true, // hasLegend: 凡例を表示
"月次売上" // title: 設定するとタイトルが表示される
);
echo "[2] 型 = " . $chartTypeName($chart1->getChartType())
. " / hasTitle = " . $bol($chart1->hasTitle())
. " / タイトル = " . $jstr($chart1->getChartTitle()->getText()) . "\n";
echo " データ源 = {$chart1->getDataSource()}"
. " / plotBy = {$chart1->getPlotBy()}"
. " / hasLegend = " . $bol($chart1->hasLegend()) . "\n";
// ---------- 3. 系列・軸の読み戻し ----------
// 系列名を直接返す API は無い。系列名は chartWizard のヘッダ行なので、
// 個数は getSeriesCollection で数え、名前はヘッダセルから対応付けて確認する。
$n1 = $countSeries($chart1);
echo "[3] 系列数 = {$n1}\n";
for ($i = 1; $i <= $n1; $i++) {
$s = $chart1->getSeriesCollection($i);
$hv = $ws->getCells(1, $i + 1)->getValue(); // 系列 i → 列 i+1 のヘッダ
$name = array\_values($hv)[0];
echo " 系列{$i}「{$name}」値 = " . $jarr($s->getValues())
. " / ポイント数 = {$s->getPointSize()}\n";
}
$catAxis = $chart1->getAxes(\phposbxl\enums\XlAxisType::Category);
echo " カテゴリ名 = " . $jarr($catAxis->getCategoryNames()) . "\n";
// ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
$co2 = $ws->addChartObject("SalesLine", 320, 250, 360, 220);
$chart2 = $co2->getChart();
$chart2->chartWizard(
"Sheet1!A1:C7",
\phposbxl\enums\XlChartType::ChartTypeLineMarkers,
\phposbxl\enums\XlRowCol::Columns, 1, 1,
1, \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
// タイトルは setTitle(true) → getChartTitle()->setText() でも設定できる
$chart2->setTitle(true);
$chart2->getChartTitle()->setText("月次売上推移");
// 値軸: 最大/最小/主単位を固定（読み戻しは getMajorUnit と ...IsAuto 系）
$valAxis = $chart2->getAxes(\phposbxl\enums\XlAxisType::Value);
$valAxis->setMaximumScale(200);
$valAxis->setMinimumScale(0);
$valAxis->setMajorUnit(50);
$valAxis->setTitle(true);
$valAxis->getTitle()->setText("売上(万円)");
// 凡例を右へ
$chart2->getLegend()->setPosition(\phposbxl\enums\XlLegendPosition::LegendPositionRight);
$right = \phposbxl\enums\XlLegendPosition::LegendPositionRight->value;
echo "[4] 型 = " . $chartTypeName($chart2->getChartType())
. " / タイトル = " . $jstr($chart2->getChartTitle()->getText()) . "\n";
echo " 値軸 主単位 = {$valAxis->getMajorUnit()}"
. " / 最大自動 = " . $bol($valAxis->getMaximumScaleIsAuto())
. " / 最小自動 = " . $bol($valAxis->getMinimumScaleIsAuto()) . "\n";
echo " 値軸タイトル = " . $jstr($valAxis->getTitle()->getText())
. " / 凡例位置 = {$chart2->getLegend()->getPosition()} (Right={$right})\n";
// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
echo "[5] チャート個数 = " . $countCharts()
. " / SalesColumn(left,top,width,height) = "
. implode(",", [$co1->getLeft(), $co1->getTop(), $co1->getWidth(), $co1->getHeight()]) . "\n";
// ---------- 6. チャートの削除 ----------
$ws->deleteChartObject("SalesLine"); // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
echo "[完了] 削除後のチャート個数 = " . $countCharts()
. " / 残 = {$ws->getChartObject(1)->getName()}\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// charts.wasm.js — チャート（グラフ）の基本: node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("charts-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// 補助: チャートタイプ数値 → enum 名
const chartTypeName = (code) =>
Object.keys(enums.XlChartType).find((k) => enums.XlChartType[k] === code) || String(code);
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
const countCharts = () => {
let n = 0;
for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
return n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
const countSeries = (chart) => {
let n = 0;
for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
return n;
};
// ---------- 1. データ表（2商品の月次売上）を作る ----------
["月", "商品A", "商品B"].forEach((h, c) => ws.getCells(1, c + 1).setValue(h));
const sales = [
["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
sales.forEach((row, i) => {
ws.getCells(i + 2, 1).setValue(row[0]);
ws.getCells(i + 2, 2).setNumberValue(row[1]);
ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] データ表 A1:C7 を作成 / 商品A合計 =",
sales.reduce((s, r) => s + r[1], 0));
// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
chart1.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeColumnClustered,
enums.XlRowCol.Columns, // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
1, // seriesLabelLines: 先頭 1 行が系列名ヘッダ
1, // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
1, // chartStyleNo（1 始まり）
enums.XlChartColorPalette.XlPaletteColorful1,
true, // hasLegend: 凡例を表示
"月次売上" // title: 設定するとタイトルが表示される
);
console.log("[2] 型 =", chartTypeName(chart1.getChartType()),
"/ hasTitle =", chart1.hasTitle(),
"/ タイトル =", JSON.stringify(chart1.getChartTitle().getText()));
console.log(" データ源 =", chart1.getDataSource(),
"/ plotBy =", chart1.getPlotBy(),
"/ hasLegend =", chart1.hasLegend());
// ---------- 3. 系列・軸の読み戻し ----------
const n1 = countSeries(chart1);
console.log("[3] 系列数 =", n1);
for (let i = 1; i <= n1; i++) {
const s = chart1.getSeriesCollection(i);
const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // 系列 i → 列 i+1 のヘッダ
console.log(` 系列${i}「${name}」値 =`, JSON.stringify(s.getValues()),
"/ ポイント数 =", s.getPointSize());
}
console.log(" カテゴリ名 =",
JSON.stringify(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()));
// ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
const co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
const chart2 = co2.getChart();
chart2.chartWizard(
"Sheet1!A1:C7",
enums.XlChartType.ChartTypeLineMarkers,
enums.XlRowCol.Columns, 1, 1,
1, enums.XlChartColorPalette.XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
chart2.setTitle(true);
chart2.getChartTitle().setText("月次売上推移");
const valAxis = chart2.getAxes(enums.XlAxisType.Value);
valAxis.setMaximumScale(200);
valAxis.setMinimumScale(0);
valAxis.setMajorUnit(50);
valAxis.setTitle(true);
valAxis.getTitle().setText("売上(万円)");
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight);
console.log("[4] 型 =", chartTypeName(chart2.getChartType()),
"/ タイトル =", JSON.stringify(chart2.getChartTitle().getText()));
console.log(" 値軸 主単位 =", valAxis.getMajorUnit(),
"/ 最大自動 =", valAxis.getMaximumScaleIsAuto(),
"/ 最小自動 =", valAxis.getMinimumScaleIsAuto());
console.log(" 値軸タイトル =", JSON.stringify(valAxis.getTitle().getText()),
"/ 凡例位置 =", chart2.getLegend().getPosition(),
`(Right=${enums.XlLegendPosition.LegendPositionRight})`);
// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
console.log("[5] チャート個数 =", countCharts(),
"/ SalesColumn(left,top,width,height) =",
[co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));
// ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine"); // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
console.log("[完了] 削除後のチャート個数 =", countCharts(),
"/ 残 =", ws.getChartObject(1).getName());
wb.save();
wb.close();
```

```plaintext
[1] データ表 A1:C7 を作成 / 商品A合計 = 882
[2] 型 = ChartTypeColumnClustered / hasTitle = true / タイトル = "月次売上"
データ源 = Sheet1!$A$1:$C$7 / plotBy = 2 / hasLegend = true
[3] 系列数 = 2
系列1「商品A」値 = ["120","135","150","142","160","175"] / ポイント数 = 6
系列2「商品B」値 = ["90","95","110","125","118","130"] / ポイント数 = 6
カテゴリ名 = ["1月","2月","3月","4月","5月","6月"]
[4] 型 = ChartTypeLineMarkers / タイトル = "月次売上推移"
値軸 主単位 = 50 / 最大自動 = false / 最小自動 = false
値軸タイトル = "売上(万円)" / 凡例位置 = 4 (Right=4)
[5] チャート個数 = 2 / SalesColumn(left,top,width,height) = 320,10,360,220
[完了] 削除後のチャート個数 = 1 / 残 = SalesColumn
```

## 解説

* **作成は 2 段階: `addChartObject()` で枠を作り、`getChart().chartWizard()` で中身を構築する。** `addChartObject(name, left, top, width, height)` が返すのは配置枠（`ChartObject`）だけで、 この時点のチャートは**未構築**です。  
  `chartObject.getChart()` で得た `Chart` に対して `chartWizard()` を呼んで初めて型・データが設定されます。  
  位置・サイズの単位は**ポイント(pt)**です。
* **`chartWizard()` の引数順は Excel マクロと違うので注意。** シグネチャは `chartWizard(source, chartType, plotBy, seriesLabelLines, categoryLabelLines, styleNo?, colorNo?, hasLegend?, title?, ...)`。  
   `source` は `"Sheet1!A1:C7"` のような A1C1 形式、`seriesLabelLines=1` は「先頭 1 行が系列名」、 `categoryLabelLines=1` は「先頭 1 列がカテゴリ名」を意味します。
* **`plotBy` は「系列をどの方向に取るか」。** 上のように月を行・商品を列（B/C 列）に並べた表で `enums.XlRowCol.Columns`(=2) を渡すと、**列ごとに 1 系列**（商品A・商品B）になります（実測で 2 系列を確認）。
* **タイトルは 2 通りで設定できる。** `chartWizard()` の `title` 引数に文字列を渡す（→ `hasTitle()` が `true` になる）か、 `chart.setTitle(true)` → `chart.getChartTitle().setText("...")` を呼ぶ方法です。  
  どちらも `getChartTitle().getText()` で読み戻せます。
* **軸は `getAxes(enums.XlAxisType.Category|Value)` で取る。** 値軸の最大/最小は `setMaximumScale()` / `setMinimumScale()`、 目盛りは `setMajorUnit()` で設定します。  
  軸タイトルは `axis.setTitle(true)` → `axis.getTitle().setText()` です。
* **系列の値は `getSeriesCollection(i).getValues()` / `getXValues()`（どちらも文字列配列）で読み戻せる。** 系列数そのものを返す API は無いため、上のサンプルは `getSeriesCollection()` が例外を投げるまで数えています。

## 注意点

* **チャートの見た目（色・凡例の実表示・折れ線の形など）はコンソールから検証できません。** このサンプルは型・タイトル文字列・系列の値・軸の設定値・個数といった**プロパティの読み戻し**で 正しく作れたかを確認しています。  
  **最終的な見栄えは必ず Excel で `charts.xlsx` を開いて目視確認してください。**
* **未構築チャートへの `setChartType()` / `setDataSource()` は例外になります。** `addChartObject()` 直後（`chartWizard()` 前）のチャートに型やデータ源を設定しようとすると `chart is not built` 例外が送出されます（セグメントフォルトはしない）。  
  必ず `chartWizard()` で構築してください。  
   無効なチャートタイプ値（`-2` や `0`、範囲外）も `chartType is invalid` 例外で拒否されます。
* **`getLegend()` は凡例が有効なチャートでのみ取得できます。** 凡例が無効のチャートでは `legend not found` 例外になります。  
  有効化する方法は 2 つ: 構築時に `chartWizard(…, hasLegend=true, …)` を渡すか、後から **`chart.setLegend(true)`** を呼ぶ （実測: `setLegend(true)` 後は `getLegend()` / `setPosition()` とも正常。  
  `setLegend(false)` で再び例外）。
* **値軸の最大/最小を読み戻す getter はありません。** `setMaximumScale()` / `setMinimumScale()` に対応する `getMaximumScale()` / `getMinimumScale()` は存在せず、`getMaximumScaleIsAuto()` / `getMinimumScaleIsAuto()`（`boolean`）で 「自動かどうか」しか確認できません。  
  手動設定後は `false` になります。  
  主単位は `getMajorUnit()` で読み戻せます。
* **系列名（凡例の項目名）を返す API はありません。** `Series` に `getName()` 相当が無く、`LegendEntry` にもテキスト取得が無いため、 サンプルでは系列名を**元表のヘッダセル（B1/C1）から読み出して対応付け**ています。
* **`getValues()` は数値セルを数値文字列で返し、文字列セルは空文字列 `""` を返します（実測）。** グラフの系列値は数値が前提のため、文字列しか入っていないセルは `""` になります。  
  系列データは `setNumberValue()` で数値として入れてください。
* **チャート個数を返す API は無く、`ws.getShapes().getCount()` も使えません（`typeof` が `undefined`）。** 個数は `getChartObject(1), getChartObject(2), ...` が例外を投げるまで数える必要があります（サンプルの `countCharts()`）。
* **`chartWizard()` の省略可能引数は途中を飛ばせません。** `title` や `hasLegend` を渡したいなら、 その手前の `styleNo` / `colorNo` にも値を与える必要があります（中間位置に `undefined` / `null` を渡すと例外）。
* **一部のチャート型は未対応。** `enums.XlChartType.ChartTypeRegionMap` は `chartWizard()` で `unsupported` 例外になります。

## 関連

* [目次](index.md)
* 前: [条件付き書式](general-09-conditional-format.md) / 次: [ピボットテーブル](general-11-pivottables.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
