<?php
$page_title = 'チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除 | osbxl サンプル';
$no_tagline = true;
$meta_description = 'チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,addchartobject,chartwizard,getchartobject,deletechartobject,getseriescollection,getaxes,getlegend,charttitle,charttype';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-10-charts.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">ワークシートのデータ表から<strong>チャート（グラフ）を作成</strong>する。 縦棒グラフと折れ線グラフの 2 種類を作り、タイトル・凡例・軸を設定し、 <strong>作成した内容をプロパティとして読み戻して検証</strong>する。最後にチャートを削除する。</p>
<p class="mb-4 text-slate-700 leading-7">チャートは見た目をコンソールから確認できないため、このサンプルでは <strong>型・タイトル文字列・系列の値・軸の設定値・チャート個数</strong>を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">console.log</code> で出力し、 意図どおりに作れたかをプログラムで確かめられるようにしている。</p>
<h2 id="対象クラス" class="text-2xl font-bold mt-10 mb-4">対象クラス</h2>
<p class="mb-4 text-slate-700 leading-7">チャートは「枠（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartObject</code>）」と「中身（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Chart</code>）」の 2 層に分かれる。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">クラス</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">取得方法</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">役割</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.ChartObject</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.addChartObject(name, left, top, w, h)</code> / `ws.getChartObject(name\</td><td class="border border-surface-200 px-3 py-2 align-top">index)`</td><td class="border border-surface-200 px-3 py-2 align-top">シート上の配置枠。位置・サイズ・名前・表示/非表示</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Chart</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartObject.getChart()</code></td><td class="border border-surface-200 px-3 py-2 align-top">チャート本体。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> で構築し、型・データ源・タイトル・軸・系列を扱う</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.ChartTitle</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getChartTitle()</code></td><td class="border border-surface-200 px-3 py-2 align-top">チャート标题。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setText()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getText()</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Axis</code></td><td class="border border-surface-200 px-3 py-2 align-top">`chart.getAxes(enums.XlAxisType.Category\</td><td class="border border-surface-200 px-3 py-2 align-top">Value)`</td><td class="border border-surface-200 px-3 py-2 align-top">軸。目盛り・最大/最小・軸タイトル</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Series</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getSeriesCollection(index)</code></td><td class="border border-surface-200 px-3 py-2 align-top">1 系列。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValues()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getXValues()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPointSize()</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.Legend</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getLegend()</code></td><td class="border border-surface-200 px-3 py-2 align-top">凡例。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPosition()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPosition()</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> 側のチャート操作メソッド:</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject(name, left, top, width, height)</code></td><td class="border border-surface-200 px-3 py-2 align-top">チャート枠を追加。<strong>位置・サイズはポイント(pt)単位</strong></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject(name)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject(index)</code></td><td class="border border-surface-200 px-3 py-2 align-top">名前 or 索引（<strong>1 始まり</strong>）で取得。範囲外は例外</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteChartObject(name)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteChartObject(index)</code></td><td class="border border-surface-200 px-3 py-2 align-top">名前 or 索引で削除</td></tr>
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
        <pre><code class="language-javascript">// charts.js — チャート（グラフ）の基本: 追加・タイトル・型・系列/軸の読み戻し・削除
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("charts.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// 補助: チャートタイプ数値 → enum 名
const chartTypeName = (code) =&gt;
    Object.keys(enums.XlChartType).find((k) =&gt; enums.XlChartType[k] === code) || String(code);
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
const countCharts = () =&gt; {
    let n = 0;
    for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
    return n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
const countSeries = (chart) =&gt; {
    let n = 0;
    for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
    return n;
};

// ---------- 1. データ表（2商品の月次売上）を作る ----------
["月", "商品A", "商品B"].forEach((h, c) =&gt; ws.getCells(1, c + 1).setValue(h));
const sales = [
    ["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
    ["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
sales.forEach((row, i) =&gt; {
    ws.getCells(i + 2, 1).setValue(row[0]);
    ws.getCells(i + 2, 2).setNumberValue(row[1]);
    ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] データ表 A1:C7 を作成 / 商品A合計 =",
    sales.reduce((s, r) =&gt; s + r[1], 0));

// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
// addChartObject(名前, left, top, width, height) — 位置・サイズはポイント(pt)単位
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
//             styleNo, colorNo, hasLegend, title)
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                        // seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                        // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                        // chartStyleNo（1 始まり）
    enums.XlChartColorPalette.XlPaletteColorful1,
    true,                     // hasLegend: 凡例を表示
    "月次売上"                // title: 設定するとタイトルが表示される
);
console.log("[2] 型 =", chartTypeName(chart1.getChartType()),
    "/ hasTitle =", chart1.hasTitle(),
    "/ タイトル =", JSON.stringify(chart1.getChartTitle().getText()));
console.log("    データ源 =", chart1.getDataSource(),
    "/ plotBy =", chart1.getPlotBy(),
    "/ hasLegend =", chart1.hasLegend());

// ---------- 3. 系列・軸の読み戻し ----------
// 系列名を直接返す API は無い。系列名は chartWizard のヘッダ行なので、
// 個数は getSeriesCollection で数え、名前はヘッダセルから対応付けて確認する。
const n1 = countSeries(chart1);
console.log("[3] 系列数 =", n1);
for (let i = 1; i &lt;= n1; i++) {
    const s = chart1.getSeriesCollection(i);
    const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // 系列 i → 列 i+1 のヘッダ
    console.log(`    系列${i}「${name}」値 =`, JSON.stringify(s.getValues()),
        "/ ポイント数 =", s.getPointSize());
}
console.log("    カテゴリ名 =",
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
console.log("    値軸 主単位 =", valAxis.getMajorUnit(),
    "/ 最大自動 =", valAxis.getMaximumScaleIsAuto(),
    "/ 最小自動 =", valAxis.getMinimumScaleIsAuto());
console.log("    値軸タイトル =", JSON.stringify(valAxis.getTitle().getText()),
    "/ 凡例位置 =", chart2.getLegend().getPosition(),
    `(Right=${enums.XlLegendPosition.LegendPositionRight})`);

// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
console.log("[5] チャート個数 =", countCharts(),
    "/ SalesColumn(left,top,width,height) =",
    [co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));

// ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine");   // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
console.log("[完了] 削除後のチャート個数 =", countCharts(),
    "/ 残 =", ws.getChartObject(1).getName());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># charts.py — チャート（グラフ）の基本: node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl

enums = pyosbxl.enums

app = pyosbxl.App()
wb = app.createWorkBook("charts-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# 補助: JSON 相当の文字列（区切りスペース無し・非 ASCII はそのまま）
def jdump(v):
    return json.dumps(v, ensure_ascii=False, separators=(",", ":"))

# 補助: enum のメンバー名（pyosbxl の enum は __name__ で名前が取れる）
def enum_name(v):
    return v.__name__

# 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
def count_charts():
    n = 0
    while True:
        try:
            ws.getChartObject(n + 1)
            n += 1
        except Exception:
            return n

# 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
def count_series(chart):
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
    enums.XlRowCol.Columns,   # 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                        # seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                        # categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                        # chartStyleNo（1 始まり）
    enums.XlChartColorPalette.XlPaletteColorful1,
    True,                     # hasLegend: 凡例を表示
    "月次売上",               # title: 設定するとタイトルが表示される
)
print("[2] 型 =", enum_name(chart1.getChartType()),
      "/ hasTitle =", str(chart1.hasTitle()).lower(),
      "/ タイトル =", jdump(chart1.getChartTitle().getText()))
print("    データ源 =", chart1.getDataSource(),
      "/ plotBy =", int(chart1.getPlotBy()),
      "/ hasLegend =", str(chart1.hasLegend()).lower())

# ---------- 3. 系列・軸の読み戻し ----------
n1 = count_series(chart1)
print("[3] 系列数 =", n1)
for i in range(1, n1 + 1):
    s = chart1.getSeriesCollection(i)
    name = list(ws.getCells(1, i + 1).getValue().values())[0]  # 系列 i → 列 i+1 のヘッダ
    print(f"    系列{i}「{name}」値 =", jdump(s.getValues()),
          "/ ポイント数 =", s.getPointSize())
print("    カテゴリ名 =",
      jdump(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()))

# ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220)
chart2 = co2.getChart()
chart2.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeLineMarkers,
    enums.XlRowCol.Columns, 1, 1,
    1, enums.XlChartColorPalette.XlPaletteColorful1, True)  # hasLegend=True にしないと getLegend() が例外
chart2.setTitle(True)
chart2.getChartTitle().setText("月次売上推移")
val_axis = chart2.getAxes(enums.XlAxisType.Value)
val_axis.setMaximumScale(200)
val_axis.setMinimumScale(0)
val_axis.setMajorUnit(50)
val_axis.setTitle(True)
val_axis.getTitle().setText("売上(万円)")
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight)

print("[4] 型 =", enum_name(chart2.getChartType()),
      "/ タイトル =", jdump(chart2.getChartTitle().getText()))
print("    値軸 主単位 =", int(val_axis.getMajorUnit()),
      "/ 最大自動 =", str(val_axis.getMaximumScaleIsAuto()).lower(),
      "/ 最小自動 =", str(val_axis.getMinimumScaleIsAuto()).lower())
print("    値軸タイトル =", jdump(val_axis.getTitle().getText()),
      "/ 凡例位置 =", int(chart2.getLegend().getPosition()),
      f"(Right={int(enums.XlLegendPosition.LegendPositionRight)})")

# ---------- 5. チャート個数・位置・サイズの読み戻し ----------
print("[5] チャート個数 =", count_charts(),
      "/ SalesColumn(left,top,width,height) =",
      f"{int(co1.getLeft())},{int(co1.getTop())},{int(co1.getWidth())},{int(co1.getHeight())}")

# ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine")   # 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
print("[完了] 削除後のチャート個数 =", count_charts(),
      "/ 残 =", ws.getChartObject(1).getName())

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// charts.java — チャート（グラフ）の基本: node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.chart.*;
import com.osboffice.osbxl.enums.*;
import java.util.List;
import java.util.Map;

public class Main {
    // 補助: 文字列リスト → JSON 相当（["a","b"] 形式。node の JSON.stringify と同じ見た目にする）
    private static String jsonArr(List&lt;String&gt; list) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i &lt; list.size(); i++) {
            if (i &gt; 0) sb.append(",");
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
        for (int c = 0; c &lt; headers.length; c++) ws.getCells(1, c + 1).setValue(headers[c]);
        String[] months = {"1月", "2月", "3月", "4月", "5月", "6月"};
        int[] salesA = {120, 135, 150, 142, 160, 175};
        int[] salesB = {90, 95, 110, 125, 118, 130};
        int totalA = 0;
        for (int i = 0; i &lt; months.length; i++) {
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
            XlRowCol.Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
            1,                  // seriesLabelLines: 先頭 1 行が系列名ヘッダ
            1,                  // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
            1,                  // chartStyleNo（1 始まり）
            XlChartColorPalette.XlPaletteColorful1,
            true,               // hasLegend: 凡例を表示
            "月次売上");         // title: 設定するとタイトルが表示される
        System.out.println("[2] 型 = " + chart1.getChartType().name()
            + " / hasTitle = " + chart1.hasTitle()
            + " / タイトル = \"" + chart1.getChartTitle().getText() + "\"");
        System.out.println("    データ源 = " + chart1.getDataSource()
            + " / plotBy = " + chart1.getPlotBy().getCode()   // enum は getCode() で数値にする
            + " / hasLegend = " + chart1.hasLegend());

        // ---------- 3. 系列・軸の読み戻し ----------
        int n1 = countSeries(chart1);
        System.out.println("[3] 系列数 = " + n1);
        for (int i = 1; i &lt;= n1; i++) {
            SeriesWrapper s = chart1.getSeriesCollection(i);
            Map&lt;String, String&gt; h = ws.getCells(1, i + 1).getValue();
            String name = h.values().iterator().next();       // 系列 i → 列 i+1 のヘッダ
            System.out.println("    系列" + i + "「" + name + "」値 = " + jsonArr(s.getValues())
                + " / ポイント数 = " + s.getPointSize());
        }
        System.out.println("    カテゴリ名 = "
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
        System.out.println("    値軸 主単位 = " + (int) valAxis.getMajorUnit()   // double → int で表示
            + " / 最大自動 = " + valAxis.getMaximumScaleIsAuto()
            + " / 最小自動 = " + valAxis.getMinimumScaleIsAuto());
        System.out.println("    値軸タイトル = \"" + valAxis.getTitle().getText() + "\""
            + " / 凡例位置 = " + chart2.getLegend().getPosition().getCode()
            + " (Right=" + XlLegendPosition.LegendPositionRight.getCode() + ")");

        // ---------- 5. チャート個数・位置・サイズの読み戻し ----------
        System.out.println("[5] チャート個数 = " + countCharts(ws)
            + " / SalesColumn(left,top,width,height) = "
            + (int) co1.getLeft() + "," + (int) co1.getTop() + ","
            + (int) co1.getWidth() + "," + (int) co1.getHeight());

        // ---------- 6. チャートの削除 ----------
        ws.deleteChartObject("SalesLine");   // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
        System.out.println("[完了] 削除後のチャート個数 = " + countCharts(ws)
            + " / 残 = " + ws.getChartObject(1).getName());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// charts.php — チャート（グラフ）の基本: node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)

// 出力ヘルパー: node 版 stdout と同じ表記にする
$bol = fn($v) =&gt; $v ? "true" : "false";                                   // bool → "true"/"false"
$jstr = fn($s) =&gt; json_encode($s, JSON_UNESCAPED_UNICODE);                 // JSON.stringify 相当（文字列）
$jarr = fn($a) =&gt; json_encode(array_values($a), JSON_UNESCAPED_UNICODE);   // JSON.stringify 相当（配列）
// 補助: チャートタイプ数値 → enum 名
$chartTypeName = fn($code) =&gt; \phposbxl\enums\XlChartType::tryFrom($code)?-&gt;name ?? (string)$code;
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
$countCharts = function () use (&amp;$ws) {
    $n = 0;
    for (;;) {
        try { $ws-&gt;getChartObject($n + 1); $n++; } catch (\Throwable $e) { break; }
    }
    return $n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
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

// ---------- 1. データ表（2商品の月次売上）を作る ----------
foreach (["月", "商品A", "商品B"] as $c =&gt; $h) {
    $ws-&gt;getCells(1, $c + 1)-&gt;setValue($h);
}
$sales = [
    ["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
    ["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
$totalA = 0;
foreach ($sales as $i =&gt; $row) {
    $ws-&gt;getCells($i + 2, 1)-&gt;setValue($row[0]);
    $ws-&gt;getCells($i + 2, 2)-&gt;setNumberValue($row[1]);
    $ws-&gt;getCells($i + 2, 3)-&gt;setNumberValue($row[2]);
    $totalA += $row[1];
}
echo "[1] データ表 A1:C7 を作成 / 商品A合計 = {$totalA}\n";

// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
// addChartObject(名前, left, top, width, height) — 位置・サイズはポイント(pt)単位
$co1 = $ws-&gt;addChartObject("SalesColumn", 320, 10, 360, 220);
$chart1 = $co1-&gt;getChart();
// chartWizard(source, type, plotBy, seriesLabelLines, categoryLabelLines,
//             styleNo, colorNo, hasLegend, title)
$chart1-&gt;chartWizard(
    "Sheet1!A1:C7",
    \phposbxl\enums\XlChartType::ChartTypeColumnClustered,
    \phposbxl\enums\XlRowCol::Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                                   // seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                                   // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                                   // chartStyleNo（1 始まり）
    \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1,
    true,                                // hasLegend: 凡例を表示
    "月次売上"                           // title: 設定するとタイトルが表示される
);
echo "[2] 型 = " . $chartTypeName($chart1-&gt;getChartType())
    . " / hasTitle = " . $bol($chart1-&gt;hasTitle())
    . " / タイトル = " . $jstr($chart1-&gt;getChartTitle()-&gt;getText()) . "\n";
echo "    データ源 = {$chart1-&gt;getDataSource()}"
    . " / plotBy = {$chart1-&gt;getPlotBy()}"
    . " / hasLegend = " . $bol($chart1-&gt;hasLegend()) . "\n";

// ---------- 3. 系列・軸の読み戻し ----------
// 系列名を直接返す API は無い。系列名は chartWizard のヘッダ行なので、
// 個数は getSeriesCollection で数え、名前はヘッダセルから対応付けて確認する。
$n1 = $countSeries($chart1);
echo "[3] 系列数 = {$n1}\n";
for ($i = 1; $i &lt;= $n1; $i++) {
    $s = $chart1-&gt;getSeriesCollection($i);
    $hv = $ws-&gt;getCells(1, $i + 1)-&gt;getValue();   // 系列 i → 列 i+1 のヘッダ
    $name = array_values($hv)[0];
    echo "    系列{$i}「{$name}」値 = " . $jarr($s-&gt;getValues())
        . " / ポイント数 = {$s-&gt;getPointSize()}\n";
}
$catAxis = $chart1-&gt;getAxes(\phposbxl\enums\XlAxisType::Category);
echo "    カテゴリ名 = " . $jarr($catAxis-&gt;getCategoryNames()) . "\n";

// ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
$co2 = $ws-&gt;addChartObject("SalesLine", 320, 250, 360, 220);
$chart2 = $co2-&gt;getChart();
$chart2-&gt;chartWizard(
    "Sheet1!A1:C7",
    \phposbxl\enums\XlChartType::ChartTypeLineMarkers,
    \phposbxl\enums\XlRowCol::Columns, 1, 1,
    1, \phposbxl\enums\XlChartColorPalette::XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
// タイトルは setTitle(true) → getChartTitle()-&gt;setText() でも設定できる
$chart2-&gt;setTitle(true);
$chart2-&gt;getChartTitle()-&gt;setText("月次売上推移");
// 値軸: 最大/最小/主単位を固定（読み戻しは getMajorUnit と ...IsAuto 系）
$valAxis = $chart2-&gt;getAxes(\phposbxl\enums\XlAxisType::Value);
$valAxis-&gt;setMaximumScale(200);
$valAxis-&gt;setMinimumScale(0);
$valAxis-&gt;setMajorUnit(50);
$valAxis-&gt;setTitle(true);
$valAxis-&gt;getTitle()-&gt;setText("売上(万円)");
// 凡例を右へ
$chart2-&gt;getLegend()-&gt;setPosition(\phposbxl\enums\XlLegendPosition::LegendPositionRight);

$right = \phposbxl\enums\XlLegendPosition::LegendPositionRight-&gt;value;
echo "[4] 型 = " . $chartTypeName($chart2-&gt;getChartType())
    . " / タイトル = " . $jstr($chart2-&gt;getChartTitle()-&gt;getText()) . "\n";
echo "    値軸 主単位 = {$valAxis-&gt;getMajorUnit()}"
    . " / 最大自動 = " . $bol($valAxis-&gt;getMaximumScaleIsAuto())
    . " / 最小自動 = " . $bol($valAxis-&gt;getMinimumScaleIsAuto()) . "\n";
echo "    値軸タイトル = " . $jstr($valAxis-&gt;getTitle()-&gt;getText())
    . " / 凡例位置 = {$chart2-&gt;getLegend()-&gt;getPosition()} (Right={$right})\n";

// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
echo "[5] チャート個数 = " . $countCharts()
    . " / SalesColumn(left,top,width,height) = "
    . implode(",", [$co1-&gt;getLeft(), $co1-&gt;getTop(), $co1-&gt;getWidth(), $co1-&gt;getHeight()]) . "\n";

// ---------- 6. チャートの削除 ----------
$ws-&gt;deleteChartObject("SalesLine");   // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
echo "[完了] 削除後のチャート個数 = " . $countCharts()
    . " / 残 = {$ws-&gt;getChartObject(1)-&gt;getName()}\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// charts.wasm.js — チャート（グラフ）の基本: node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("charts-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// 補助: チャートタイプ数値 → enum 名
const chartTypeName = (code) =&gt;
    Object.keys(enums.XlChartType).find((k) =&gt; enums.XlChartType[k] === code) || String(code);
// 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
const countCharts = () =&gt; {
    let n = 0;
    for (;;) { try { ws.getChartObject(n + 1); n++; } catch (e) { break; } }
    return n;
};
// 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
const countSeries = (chart) =&gt; {
    let n = 0;
    for (;;) { try { chart.getSeriesCollection(n + 1); n++; } catch (e) { break; } }
    return n;
};

// ---------- 1. データ表（2商品の月次売上）を作る ----------
["月", "商品A", "商品B"].forEach((h, c) =&gt; ws.getCells(1, c + 1).setValue(h));
const sales = [
    ["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
    ["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
];
sales.forEach((row, i) =&gt; {
    ws.getCells(i + 2, 1).setValue(row[0]);
    ws.getCells(i + 2, 2).setNumberValue(row[1]);
    ws.getCells(i + 2, 3).setNumberValue(row[2]);
});
console.log("[1] データ表 A1:C7 を作成 / 商品A合計 =",
    sales.reduce((s, r) =&gt; s + r[1], 0));

// ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
const co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
const chart1 = co1.getChart();
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                        // seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                        // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                        // chartStyleNo（1 始まり）
    enums.XlChartColorPalette.XlPaletteColorful1,
    true,                     // hasLegend: 凡例を表示
    "月次売上"                // title: 設定するとタイトルが表示される
);
console.log("[2] 型 =", chartTypeName(chart1.getChartType()),
    "/ hasTitle =", chart1.hasTitle(),
    "/ タイトル =", JSON.stringify(chart1.getChartTitle().getText()));
console.log("    データ源 =", chart1.getDataSource(),
    "/ plotBy =", chart1.getPlotBy(),
    "/ hasLegend =", chart1.hasLegend());

// ---------- 3. 系列・軸の読み戻し ----------
const n1 = countSeries(chart1);
console.log("[3] 系列数 =", n1);
for (let i = 1; i &lt;= n1; i++) {
    const s = chart1.getSeriesCollection(i);
    const name = Object.values(ws.getCells(1, i + 1).getValue())[0]; // 系列 i → 列 i+1 のヘッダ
    console.log(`    系列${i}「${name}」値 =`, JSON.stringify(s.getValues()),
        "/ ポイント数 =", s.getPointSize());
}
console.log("    カテゴリ名 =",
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
console.log("    値軸 主単位 =", valAxis.getMajorUnit(),
    "/ 最大自動 =", valAxis.getMaximumScaleIsAuto(),
    "/ 最小自動 =", valAxis.getMinimumScaleIsAuto());
console.log("    値軸タイトル =", JSON.stringify(valAxis.getTitle().getText()),
    "/ 凡例位置 =", chart2.getLegend().getPosition(),
    `(Right=${enums.XlLegendPosition.LegendPositionRight})`);

// ---------- 5. チャート個数・位置・サイズの読み戻し ----------
console.log("[5] チャート個数 =", countCharts(),
    "/ SalesColumn(left,top,width,height) =",
    [co1.getLeft(), co1.getTop(), co1.getWidth(), co1.getHeight()].join(","));

// ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine");   // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
console.log("[完了] 削除後のチャート個数 =", countCharts(),
    "/ 残 =", ws.getChartObject(1).getName());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] データ表 A1:C7 を作成 / 商品A合計 = 882
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
[完了] 削除後のチャート個数 = 1 / 残 = SalesColumn</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>作成は 2 段階: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject()</code> で枠を作り、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChart().chartWizard()</code> で中身を構築する。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject(name, left, top, width, height)</code> が返すのは配置枠（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartObject</code>）だけで、 この時点のチャートは<strong>未構築</strong>です。<br><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartObject.getChart()</code> で得た <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Chart</code> に対して <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> を呼んで初めて型・データが設定されます。<br>位置・サイズの単位は<strong>ポイント(pt)</strong>です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> の引数順は Excel マクロと違うので注意。</strong> シグネチャは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard(source, chartType, plotBy, seriesLabelLines, categoryLabelLines, styleNo?, colorNo?, hasLegend?, title?, ...)</code>。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">source</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"Sheet1!A1:C7"</code> のような A1C1 形式、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">seriesLabelLines=1</code> は「先頭 1 行が系列名」、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">categoryLabelLines=1</code> は「先頭 1 列がカテゴリ名」を意味します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">plotBy</code> は「系列をどの方向に取るか」。</strong> 上のように月を行・商品を列（B/C 列）に並べた表で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlRowCol.Columns</code>(=2) を渡すと、<strong>列ごとに 1 系列</strong>（商品A・商品B）になります（実測で 2 系列を確認）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>タイトルは 2 通りで設定できる。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">title</code> 引数に文字列を渡す（→ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">hasTitle()</code> が <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> になる）か、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.setTitle(true)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.getChartTitle().setText("...")</code> を呼ぶ方法です。<br>どちらも <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartTitle().getText()</code> で読み戻せます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>軸は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAxes(enums.XlAxisType.Category|Value)</code> で取る。</strong> 値軸の最大/最小は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMaximumScale()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMinimumScale()</code>、 目盛りは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMajorUnit()</code> で設定します。<br>軸タイトルは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">axis.setTitle(true)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">axis.getTitle().setText()</code> です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>系列の値は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSeriesCollection(i).getValues()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getXValues()</code>（どちらも文字列配列）で読み戻せる。</strong> 系列数そのものを返す API は無いため、上のサンプルは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSeriesCollection()</code> が例外を投げるまで数えています。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>チャートの見た目（色・凡例の実表示・折れ線の形など）はコンソールから検証できません。</strong> このサンプルは型・タイトル文字列・系列の値・軸の設定値・個数といった<strong>プロパティの読み戻し</strong>で 正しく作れたかを確認しています。<br><strong>最終的な見栄えは必ず Excel で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">charts.xlsx</code> を開いて目視確認してください。</strong></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>未構築チャートへの <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setChartType()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataSource()</code> は例外になります。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject()</code> 直後（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> 前）のチャートに型やデータ源を設定しようとすると <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart is not built</code> 例外が送出されます（セグメントフォルトはしない）。<br>必ず <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> で構築してください。<br> 無効なチャートタイプ値（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">-2</code> や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0</code>、範囲外）も <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartType is invalid</code> 例外で拒否されます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getLegend()</code> は凡例が有効なチャートでのみ取得できます。</strong> 凡例が無効のチャートでは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">legend not found</code> 例外になります。<br>有効化する方法は 2 つ: 構築時に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard(…, hasLegend=true, …)</code> を渡すか、後から <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart.setLegend(true)</code></strong> を呼ぶ （実測: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setLegend(true)</code> 後は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getLegend()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPosition()</code> とも正常。<br><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setLegend(false)</code> で再び例外）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>値軸の最大/最小を読み戻す getter はありません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMaximumScale()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMinimumScale()</code> に対応する <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumScale()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumScale()</code> は存在せず、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumScaleIsAuto()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumScaleIsAuto()</code>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">boolean</code>）で 「自動かどうか」しか確認できません。<br>手動設定後は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code> になります。<br>主単位は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMajorUnit()</code> で読み戻せます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>系列名（凡例の項目名）を返す API はありません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Series</code> に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> 相当が無く、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">LegendEntry</code> にもテキスト取得が無いため、 サンプルでは系列名を<strong>元表のヘッダセル（B1/C1）から読み出して対応付け</strong>ています。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValues()</code> は数値セルを数値文字列で返し、文字列セルは空文字列 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">""</code> を返します（実測）。</strong> グラフの系列値は数値が前提のため、文字列しか入っていないセルは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">""</code> になります。<br>系列データは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue()</code> で数値として入れてください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>チャート個数を返す API は無く、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getShapes().getCount()</code> も使えません（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">typeof</code> が <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code>）。</strong> 個数は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject(1), getChartObject(2), ...</code> が例外を投げるまで数える必要があります（サンプルの <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">countCharts()</code>）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> の省略可能引数は途中を飛ばせません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">title</code> や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">hasLegend</code> を渡したいなら、 その手前の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">styleNo</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colorNo</code> にも値を与える必要があります（中間位置に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">undefined</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">null</code> を渡すと例外）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>一部のチャート型は未対応。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlChartType.ChartTypeRegionMap</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chartWizard()</code> で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">unsupported</code> 例外になります。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">条件付き書式</a> / 次: <a href="general-11-pivottables.php" class="text-primary-600 hover:underline">ピボットテーブル</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/chart.ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/nodejs/classes/chart.Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/nodejs/classes/chart.ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/nodejs/classes/chart.Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/nodejs/classes/chart.Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/nodejs/classes/chart.Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/nodejs/enums/enums.XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/nodejs/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/nodejs/enums/enums.XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/nodejs/enums/enums.XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.chart.ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/python/pyosbxl.chart.Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/python/pyosbxl.chart.ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/python/pyosbxl.chart.Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/python/pyosbxl.chart.Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/python/pyosbxl.chart.Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/python/pyosbxl.enums.XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/python/pyosbxl.enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/python/pyosbxl.enums.XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/python/pyosbxl.enums.XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/chart/ChartObjectWrapper.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/chart/ChartWrapper.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/java/com/osboffice/osbxl/chart/ChartTitleWrapper.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/java/com/osboffice/osbxl/chart/AxisWrapper.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/java/com/osboffice/osbxl/chart/SeriesWrapper.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/java/com/osboffice/osbxl/chart/LegendWrapper.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-chart-ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/php/classes/phposbxl-chart-Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/php/classes/phposbxl-chart-ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/php/classes/phposbxl-chart-Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/php/classes/phposbxl-chart-Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/php/classes/phposbxl-chart-Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/chart.ChartObject.html" class="text-primary-600 hover:underline">ChartObject</a> / <a href="/api-docs/wasm/classes/chart.Chart.html" class="text-primary-600 hover:underline">Chart</a> / <a href="/api-docs/wasm/classes/chart.ChartTitle.html" class="text-primary-600 hover:underline">ChartTitle</a> / <a href="/api-docs/wasm/classes/chart.Axis.html" class="text-primary-600 hover:underline">Axis</a> / <a href="/api-docs/wasm/classes/chart.Series.html" class="text-primary-600 hover:underline">Series</a> / <a href="/api-docs/wasm/classes/chart.Legend.html" class="text-primary-600 hover:underline">Legend</a> / <a href="/api-docs/wasm/enums/enums.XlChartType.html" class="text-primary-600 hover:underline">XlChartType</a> / <a href="/api-docs/wasm/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a> / <a href="/api-docs/wasm/enums/enums.XlAxisType.html" class="text-primary-600 hover:underline">XlAxisType</a> / <a href="/api-docs/wasm/enums/enums.XlLegendPosition.html" class="text-primary-600 hover:underline">XlLegendPosition</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除",
   "description": "チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-10-charts.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-10-charts.php",
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
     "name": "チャート（グラフ）の基本 — 追加・タイトル・型・系列/軸の読み戻し・削除",
     "item": "https://osb-office.com/sample/general-10-charts.php"
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
