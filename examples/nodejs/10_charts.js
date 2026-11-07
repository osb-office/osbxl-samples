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
    for (;;) { try { ws.getChartObject(n); n++; } catch (e) { break; } }
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
//             styleNo, colorNo, hasLegend, title)
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                        // seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                        // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                        // chartStyleNo（0 始まり (zeroBase)）
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
for (let i = 1; i <= n1; i++) {
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
    "/ 残 =", ws.getChartObject(0).getName());

wb.save();
wb.close();