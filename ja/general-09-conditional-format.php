<?php
$page_title = '条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件 | osbxl サンプル';
$no_tagline = true;
$meta_description = '条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,formatconditions,cellvalue,colorscale,databar,top10,aboveaverage,iconset,uniquevalues,apply,removecondition,colorobject,patternfill';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-09-conditional-format.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">「60 より大きいセルだけ赤くする」「値の大きさをグラデーションやバーで表す」といった <strong>条件付き書式のルール</strong>をプログラムから設定する。 あわせて、登録したルールの<strong>個数・種類・しきい値・適用範囲の読み戻し</strong>と<strong>削除</strong>を行う。</p>
<p class="mb-4 text-slate-700 leading-7">条件付き書式は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getFormatConditions()</code> で取得する <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditions</code> クラスで扱います。 流れは必ず次の 3 ステップです。</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> の条件オブジェクト（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CellValueObject</code> など）を作り、<strong>適用範囲と書式をセット</strong>する</li>
  <li><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">fc.set〜Condition(priority, obj)</code> で登録する（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">priority</code> は適用順序。1 始まりでシート単位に一意）</li>
  <li><strong>全条件を登録し終えたら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">fc.apply()</code> を呼ぶ</strong>（必須）</li>
</ol>
<h2 id="対象クラス" class="text-2xl font-bold mt-10 mb-4">対象クラス</h2>
<h3 id="formatconditions-主要メソッド" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditions</code>（主要メソッド）</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">条件の種類</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">設定</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">読み戻し</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">条件オブジェクト</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">セル値（大于・小于 など）</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCellValueCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCellValueCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CellValueObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">セル値の範囲（between）</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCellValueRangeCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCellValueRangeCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CellValueRangeObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">カラースケール</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColorScaleFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColorScaleFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.ColorScaleObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">データバー</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataBarFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDataBarFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.DataBarObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">上位 / 下位 N 件</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTop10FormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTop10FormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.Top10Object</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">平均比較</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAboveAverageCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAboveAverageCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.AboveAverageObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">アイコンセット</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setIconsetFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getIconsetFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.IconsetConditionObject</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">ユニーク値</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesFormatCondition(p, obj)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getUniqueValuesFormatCondition(p)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.UniqueValuesObject</code></td></tr>
  </tbody></table></div>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">その他のメソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code></td><td class="border border-surface-200 px-3 py-2 align-top">条件の計算・反映。<strong>全条件設定後に必ず呼ぶ</strong>。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code></td><td class="border border-surface-200 px-3 py-2 align-top">全ルールの一覧。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ priority: { セル範囲: 条件タイプ } }</code> の入れ子マップ。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeCondition(p)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeAllCondition()</code></td><td class="border border-surface-200 px-3 py-2 align-top">ルールの個別削除 / 全削除。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAppliedFontObject(range)</code> など</td><td class="border border-surface-200 px-3 py-2 align-top">条件を満たして実際に書式が付いたセルの取得（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> 後に使う）。</td></tr>
  </tbody></table></div>
<h3 id="条件オブジェクトの共通メソッド" class="text-xl font-bold mt-8 mb-3">条件オブジェクトの共通メソッド</h3>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1(適用範囲)</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFont(dto.FontObject)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPatternFill(dto.PatternFillObject)</code> を持ちます。 色は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.ColorObject</code> に <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#</code> なしの 6 桁 hex</strong>（例 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"9C0006"</code>）を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHexColor()</code> で設定します。</p>
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
        <pre><code class="language-javascript">// conditional-format.js — 条件付き書式ルールの追加・読み戻し・削除
const { nodeosbxl, enums, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("condfmt.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
const color = (hex) =&gt; {
    const c = new dto.ColorObject();
    c.setHexColor(hex);
    return c;
};

// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"].forEach((h, i) =&gt; {
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

// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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

// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);

// ---------- (3) データバー: D列 ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);

// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true);                       // true = 上位（false なら下位）
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

fc.apply();                                  // 全条件の設定後に apply() が必須
console.log("[1] 5 件のルールを登録し apply() しました");

// ---------- ルール一覧の読み戻し ----------
const typeName = Object.fromEntries(
    Object.entries(enums.XlFormatConditionType).map(([name, val]) =&gt; [val, name]));
console.log(`[2] ルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
    for (const [range, type] of Object.entries(ranges)) {
        console.log(`    priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
    }
}

// ---------- 各ルールの内容の読み戻し ----------
console.log("[3] ルール内容の読み戻し");
const g1 = fc.getCellValueCondition(1);
console.log(`    セル値  : ${g1.getA1C1()} 値=${g1.getConditionValue()}` +
    ` 演算子=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
    ` 文字色=${g1.getFont().getColorObject().getHexColor()} 太字=${g1.getFont().isBold()}` +
    ` 背景=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(`    カラースケール: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
    ` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, 閾値=${g2.getMediumValue()})` +
    ` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(`    データバー: 色=${g3.getDataBarColor().getHexColor()}` +
    ` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(`    上位 N 件: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(`    平均比較: aboveBelow=${g5.getAboveBelow()}` +
    ` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);

// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
const byRow = (a, b) =&gt; Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] 書式が適用されたセル");
console.log(`    B列 (&gt;60)     : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(`    E列 (上位/平均): ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(`    D列 データバー : ${Object.values(appliedBar).filter(Boolean).length} セル`);

// ---------- 削除 ----------
fc.removeCondition(4);                       // 上位 3 件のルールだけ削除
console.log(`[5] removeCondition(4) 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition();                     // 全ルールを削除
console.log(`[6] removeAllCondition() 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);

// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal);
fc.apply();

wb.save();
wb.close();
console.log("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）");</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># conditional-format.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

e = pyosbxl.enums


def color(hex):  # ColorObject は "#" なしの 6 桁 hex で作る
    c = pyosbxl.dto.ColorObject()
    c.setHexColor(hex)
    return c


def num(v):  # JS の数値と同じ表記にする（60.0 → "60"）
    return str(int(v)) if float(v) == int(v) else str(v)


def bol(v):  # JS の bool と同じ表記にする（True → "true"）
    return str(bool(v)).lower()


app = pyosbxl.App()
wb = app.createWorkBook("condfmt-py.xlsx")  # 出力文言は node 版に合わせ condfmt.xlsx と表示
ws = wb.openWorkSheet("Sheet1")

# ---------- データ表 ----------
# B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
for i, h in enumerate(["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"]):
    col = "ABCDE"[i]
    ws.getRange(f"{col}1:{col}1").setValue(h)
sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49]
for i, v in enumerate(sales):
    ws.getRange(f"A{i + 2}:A{i + 2}").setNumberValue(i + 1)
    for col in ["B", "C", "D", "E"]:
        ws.getRange(f"{col}{i + 2}:{col}{i + 2}").setNumberValue(v)

fc = ws.getFormatConditions()

# ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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

# ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
colorScale = pyosbxl.dto.ColorScaleObject()
colorScale.setA1C1("C2:C11")
colorScale.setMinimumCondition(color("F8696B"), e.XlConditionValueType.ConditionValueTypeMinimum)
colorScale.setMediumCondition(color("FFEB84"), e.XlConditionValueType.ConditionValueTypePercentile, 50)
colorScale.setMaximumCondition(color("63BE7B"), e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setColorScaleFormatCondition(2, colorScale)

# ---------- (3) データバー: D列 ----------
dataBar = pyosbxl.dto.DataBarObject()
dataBar.setA1C1("D2:D11")
dataBar.setDataBarColor(color("638EC6"))
dataBar.setMinimumCondition(e.XlConditionValueType.ConditionValueTypeMinimum)
dataBar.setMaximumCondition(e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setDataBarFormatCondition(3, dataBar)

# ---------- (4) 上位 3 件 + 平均より上: E列 ----------
top3 = pyosbxl.dto.Top10Object()
top3.setRank(3, True)                        # True = 上位（False なら下位）
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

fc.apply()                                   # 全条件の設定後に apply() が必須
print("[1] 5 件のルールを登録し apply() しました")

# ---------- ルール一覧の読み戻し ----------
types = fc.getConditionTypes()
print(f"[2] ルール数 = {len(types)}")
for prio in sorted(types):
    for rng, t in types[prio].items():
        print(f"    priority {prio}: {rng} → {t.name} (type={int(t)})")

# ---------- 各ルールの内容の読み戻し ----------
print("[3] ルール内容の読み戻し")
g1 = fc.getCellValueCondition(1)
print(f"    セル値  : {g1.getA1C1()} 値={num(g1.getConditionValue())}"
      f" 演算子={int(g1.getConditionOperetor())}"
      f" (Greater={int(e.XlFormatConditionOperator.FormatConditionGreater)})"
      f" 文字色={g1.getFont().getColorObject().getHexColor()} 太字={bol(g1.getFont().isBold())}"
      f" 背景={g1.getPatternFill().getCellColor().getHexColor()}")
g2 = fc.getColorScaleFormatCondition(2)
print(f"    カラースケール: min={g2.getMinimumColor().getHexColor()} (type={int(g2.getMinimumValueType())})"
      f" mid={g2.getMediumColor().getHexColor()} (type={int(g2.getMediumValueType())}, 閾値={num(g2.getMediumValue())})"
      f" max={g2.getMaximumColor().getHexColor()} (type={int(g2.getMaximumValueType())})")
g3 = fc.getDataBarFormatCondition(3)
print(f"    データバー: 色={g3.getDataBarColor().getHexColor()}"
      f" minType={int(g3.getMinimumCondition())} maxType={int(g3.getMaximumCondition())}")
g4 = fc.getTop10FormatCondition(4)
print(f"    上位 N 件: rank={g4.getRank()} isTop={bol(g4.getTopBottom())}")
g5 = fc.getAboveAverageCondition(5)
print(f"    平均比較: aboveBelow={int(g5.getAboveBelow())}"
      f" (XlAboveAverage={int(e.XlAboveBelow.XlAboveAverage)})")

# ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
byRow = lambda a: int(a[1:])
print("[4] 書式が適用されたセル")
appliedB = sorted(fc.getAppliedFontObject("B2:B11").keys(), key=byRow)
print(f"    B列 (&gt;60)     : {' '.join(appliedB)}")
appliedE = sorted(fc.getAppliedFontObject("E2:E11").keys(), key=byRow)
print(f"    E列 (上位/平均): {' '.join(appliedE)}")
appliedBar = fc.getAppliedDataBar("D2:D11")
print(f"    D列 データバー : {sum(1 for v in appliedBar.values() if v)} セル")

# ---------- 削除 ----------
fc.removeCondition(4)                        # 上位 3 件のルールだけ削除
print(f"[5] removeCondition(4) 後のルール数 = {len(fc.getConditionTypes())}")
fc.removeAllCondition()                      # 全ルールを削除
print(f"[6] removeAllCondition() 後のルール数 = {len(fc.getConditionTypes())}")

# セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal)
fc.apply()

wb.save()
wb.close()
print("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）")</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// conditional-format.java — node 版サンプルの javaosbxl 対応状況を検証する
//
// 【API 差】javaosbxl 1.5.0 の条件オブジェクト Wrapper（CellValueObjectWrapper など）には
// 基底クラス相当の setA1C1() / setFont() / setPatternFill() が公開されていません。
// 適用範囲(A1C1)を設定する手段が無いため setCellValueCondition() は
// "apply range(A1C1) must be set" 例外になり、条件付き書式ルールを登録できません。
// ここでは node 版と同等のルール登録を試みて例外になること、
// apply() / getConditionTypes() / removeCondition() / removeAllCondition() が
// 例外なく呼べることを確認し、その差分を明示します。
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("condfmt-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- データ表（node 版と同一） ----------
        String[] headers = {"月", "セル値ルール", "カラースケール", "データバー", "上位/平均"};
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

        // ---------- (1) セル値ルール: 60 より大きい ----------
        // node 版では cellVal.setA1C1("B2:B11") / setFont(...) / setPatternFill(...) を行うが、
        // javaosbxl の CellValueObjectWrapper にはこれらのメソッドが無い（API 差）
        CellValueObjectWrapper cellVal = new CellValueObjectWrapper();
        cellVal.setValueCondition(60, XlFormatConditionOperator.FormatConditionGreater);
        try {
            fc.setCellValueCondition(1, cellVal);
            System.out.println("[1] セル値ルールを登録できました");
        } catch (RuntimeException e) {
            // A1C1 未設定のため登録は必ず例外になる（javaosbxl 1.5.0 の既知の API 差）
            System.out.println("[1] セル値ルール登録は例外: " + e.getMessage());
        }
        fc.apply();  // ルール 0 件でも apply() は例外なく呼べる

        // ---------- ルール一覧の読み戻し（登録できていないので 0 件） ----------
        Map&lt;Integer, Map&lt;String, XlFormatConditionType&gt;&gt; types = fc.getConditionTypes();
        System.out.println("[2] ルール数 = " + types.size());

        // ---------- 削除系もルール 0 件で例外なく呼べる ----------
        fc.removeCondition(4);
        System.out.println("[5] removeCondition(4) 後のルール数 = " + fc.getConditionTypes().size());
        fc.removeAllCondition();
        System.out.println("[6] removeAllCondition() 後のルール数 = " + fc.getConditionTypes().size());

        wb.save();
        wb.close();
        System.out.println("[done] condfmt-java.xlsx を保存しました（javaosbxl は条件付き書式ルールを登録できないため 0 件）");
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// conditional-format.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)

// 出力ヘルパー: node 版 stdout と同じ表記にする (bool → "true"/"false")
$bol = fn($v) =&gt; $v ? "true" : "false";
// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
$color = function (string $hex): \phposbxl\dto\ColorObject {
    $c = new \phposbxl\dto\ColorObject();
    $c-&gt;setHexColor($hex);
    return $c;
};

$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("condfmt-php.xlsx"); // 出力文言は node 版に合わせ condfmt.xlsx と表示
$ws = $wb-&gt;openWorkSheet("Sheet1");

// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
foreach (["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"] as $i =&gt; $h) {
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

// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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

// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
$colorScale = new \phposbxl\dto\ColorScaleObject();
$colorScale-&gt;setA1C1("C2:C11");
$colorScale-&gt;setMinimumCondition($color("F8696B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$colorScale-&gt;setMediumCondition($color("FFEB84"), \phposbxl\enums\XlConditionValueType::ConditionValueTypePercentile, 50);
$colorScale-&gt;setMaximumCondition($color("63BE7B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc-&gt;setColorScaleFormatCondition(2, $colorScale);

// ---------- (3) データバー: D列 ----------
$dataBar = new \phposbxl\dto\DataBarObject();
$dataBar-&gt;setA1C1("D2:D11");
$dataBar-&gt;setDataBarColor($color("638EC6"));
$dataBar-&gt;setMinimumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$dataBar-&gt;setMaximumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc-&gt;setDataBarFormatCondition(3, $dataBar);

// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
$top3 = new \phposbxl\dto\Top10Object();
$top3-&gt;setRank(3, true);                      // true = 上位（false なら下位）
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

$fc-&gt;apply();                                 // 全条件の設定後に apply() が必須
echo "[1] 5 件のルールを登録し apply() しました\n";

// ---------- ルール一覧の読み戻し ----------
$types = $fc-&gt;getConditionTypes();            // array&lt;int, array&lt;string, int&gt;&gt;
ksort($types);                                // node の Object.entries と同じく priority 昇順で走査する
echo "[2] ルール数 = " . count($types) . "\n";
foreach ($types as $prio =&gt; $ranges) {
    foreach ($ranges as $range =&gt; $type) {
        $name = \phposbxl\enums\XlFormatConditionType::from($type)-&gt;name;
        echo "    priority {$prio}: {$range} → {$name} (type={$type})\n";
    }
}

// ---------- 各ルールの内容の読み戻し ----------
echo "[3] ルール内容の読み戻し\n";
$g1 = $fc-&gt;getCellValueCondition(1);
$greater = \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater-&gt;value;
echo "    セル値  : {$g1-&gt;getA1C1()} 値={$g1-&gt;getConditionValue()}"
    . " 演算子={$g1-&gt;getConditionOperetor()} (Greater={$greater})"
    . " 文字色={$g1-&gt;getFont()-&gt;getColorObject()-&gt;getHexColor()} 太字=" . $bol($g1-&gt;getFont()-&gt;isBold())
    . " 背景={$g1-&gt;getPatternFill()-&gt;getCellColor()-&gt;getHexColor()}\n";
$g2 = $fc-&gt;getColorScaleFormatCondition(2);
echo "    カラースケール: min={$g2-&gt;getMinimumColor()-&gt;getHexColor()} (type={$g2-&gt;getMinimumValueType()})"
    . " mid={$g2-&gt;getMediumColor()-&gt;getHexColor()} (type={$g2-&gt;getMediumValueType()}, 閾値={$g2-&gt;getMediumValue()})"
    . " max={$g2-&gt;getMaximumColor()-&gt;getHexColor()} (type={$g2-&gt;getMaximumValueType()})\n";
$g3 = $fc-&gt;getDataBarFormatCondition(3);
echo "    データバー: 色={$g3-&gt;getDataBarColor()-&gt;getHexColor()}"
    . " minType={$g3-&gt;getMinimumCondition()} maxType={$g3-&gt;getMaximumCondition()}\n";
$g4 = $fc-&gt;getTop10FormatCondition(4);
echo "    上位 N 件: rank={$g4-&gt;getRank()} isTop=" . $bol($g4-&gt;getTopBottom()) . "\n";
$g5 = $fc-&gt;getAboveAverageCondition(5);
$xlAbove = \phposbxl\enums\XlAboveBelow::XlAboveAverage-&gt;value;
echo "    平均比較: aboveBelow={$g5-&gt;getAboveBelow()} (XlAboveAverage={$xlAbove})\n";

// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
$byRow = fn($a, $b) =&gt; (int)substr($a, 1) - (int)substr($b, 1);
echo "[4] 書式が適用されたセル\n";
$appliedB = array_keys($fc-&gt;getAppliedFontObject("B2:B11"));
usort($appliedB, $byRow);
echo "    B列 (&gt;60)     : " . implode(" ", $appliedB) . "\n";
$appliedE = array_keys($fc-&gt;getAppliedFontObject("E2:E11"));
usort($appliedE, $byRow);
echo "    E列 (上位/平均): " . implode(" ", $appliedE) . "\n";
$appliedBar = $fc-&gt;getAppliedDataBar("D2:D11");
echo "    D列 データバー : " . count(array_filter($appliedBar)) . " セル\n";

// ---------- 削除 ----------
$fc-&gt;removeCondition(4);                      // 上位 3 件のルールだけ削除
echo "[5] removeCondition(4) 後のルール数 = " . count($fc-&gt;getConditionTypes()) . "\n";
$fc-&gt;removeAllCondition();                    // 全ルールを削除
echo "[6] removeAllCondition() 後のルール数 = " . count($fc-&gt;getConditionTypes()) . "\n";

// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
$fc-&gt;setCellValueCondition(1, $cellVal);
$fc-&gt;apply();

$wb-&gt;save();
$wb-&gt;close();
echo "[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）\n";</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// conditional-format.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("condfmt-wasm.xlsx"); // 出力文言は node 版と同一（condfmt.xlsx と表示）
const ws = wb.openWorkSheet("Sheet1");

// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
const color = (hex) =&gt; {
    const c = new dto.ColorObject();
    c.setHexColor(hex);
    return c;
};

// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"].forEach((h, i) =&gt; {
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

// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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

// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);

// ---------- (3) データバー: D列 ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);

// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true);                       // true = 上位（false なら下位）
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

fc.apply();                                  // 全条件の設定後に apply() が必須
console.log("[1] 5 件のルールを登録し apply() しました");

// ---------- ルール一覧の読み戻し ----------
const typeName = Object.fromEntries(
    Object.entries(enums.XlFormatConditionType).map(([name, val]) =&gt; [val, name]));
console.log(`[2] ルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
    for (const [range, type] of Object.entries(ranges)) {
        console.log(`    priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
    }
}

// ---------- 各ルールの内容の読み戻し ----------
console.log("[3] ルール内容の読み戻し");
const g1 = fc.getCellValueCondition(1);
console.log(`    セル値  : ${g1.getA1C1()} 値=${g1.getConditionValue()}` +
    ` 演算子=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
    ` 文字色=${g1.getFont().getColorObject().getHexColor()} 太字=${g1.getFont().isBold()}` +
    ` 背景=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(`    カラースケール: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
    ` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, 閾値=${g2.getMediumValue()})` +
    ` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(`    データバー: 色=${g3.getDataBarColor().getHexColor()}` +
    ` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(`    上位 N 件: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(`    平均比較: aboveBelow=${g5.getAboveBelow()}` +
    ` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);

// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
const byRow = (a, b) =&gt; Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] 書式が適用されたセル");
console.log(`    B列 (&gt;60)     : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(`    E列 (上位/平均): ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(`    D列 データバー : ${Object.values(appliedBar).filter(Boolean).length} セル`);

// ---------- 削除 ----------
fc.removeCondition(4);                       // 上位 3 件のルールだけ削除
console.log(`[5] removeCondition(4) 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition();                     // 全ルールを削除
console.log(`[6] removeAllCondition() 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);

// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal);
fc.apply();

wb.save();
wb.close();
console.log("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）");</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] 5 件のルールを登録し apply() しました
[2] ルール数 = 5
    priority 1: B2:B11 → CellValue (type=1)
    priority 2: C2:C11 → ColorScale (type=3)
    priority 3: D2:D11 → Databar (type=4)
    priority 4: E2:E11 → Top10 (type=5)
    priority 5: E2:E11 → AboveAverageCondition (type=12)
[3] ルール内容の読み戻し
    セル値  : B2:B11 値=60 演算子=5 (Greater=5) 文字色=9C0006 太字=true 背景=FFC7CE
    カラースケール: min=F8696B (type=5) mid=FFEB84 (type=4, 閾値=50) max=63BE7B (type=6)
    データバー: 色=638EC6 minType=5 maxType=6
    上位 N 件: rank=3 isTop=true
    平均比較: aboveBelow=0 (XlAboveAverage=0)
[4] 書式が適用されたセル
    B列 (&gt;60)     : B3 B5 B8 B10
    E列 (上位/平均): E3 E5 E6 E8 E10
    D列 データバー : 10 セル
[5] removeCondition(4) 後のルール数 = 4
[6] removeAllCondition() 後のルール数 = 0
[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7">B 列の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">B3 B5 B8 B10</code> は値が 78・91・87・65（= 60 より大きい）のセル、 E 列の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">E3 E5 E6 E8 E10</code> は「上位 3 件」と「平均（54.2）より上」の和集合です。 見た目の書式は検証できないため、このように<strong>ルールの個数・種類・しきい値・適用されたセル</strong>を 読み戻して出力することで、意図どおり設定できたことを確認しています。</p>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<h3 id="priority-適用順序" class="text-xl font-bold mt-8 mb-3">priority（適用順序）</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">ルールは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">priority</code>（1 始まりの適用順序）で管理され、<strong>シート単位で一意</strong>です。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code> の返却値のキーがこの priority です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>既存の priority に登録すると、既存ルールが後ろへずれます</strong>（実測）。<br> たとえば priority 1 に 2 回登録すると、先に登録したルールは priority 2 になります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeCondition(p)</code> で削除すると、<strong>後ろの priority が繰り上がります</strong>。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">読み戻しメソッド（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCellValueCondition(p)</code> など）は <strong>priority の位置に別の種類のルールがあると例外</strong>になります。<br> まず <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code> で種類を確認してから取得してください。</li>
</ul>
<h3 id="apply-の役割" class="text-xl font-bold mt-8 mb-3">apply() の役割</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> を呼ぶまで条件は評価されません。</strong> 実測では、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> 前の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAppliedFontObject()</code> は空 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{}</code> を返し、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> 後は条件を満たすセル（例では B3 B5 B8 B10）が返りました。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">ルールを追加・削除したら、その都度 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> を呼んでください。</li>
</ul>
<h3 id="主な-enum" class="text-xl font-bold mt-8 mb-3">主な enum</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">enum</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">よく使う値</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlFormatConditionOperator</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionGreater</code>(5) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionGreaterEqual</code>(7) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionLess</code>(6) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionLessEqual</code>(8) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionEqual</code>(3) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FormatConditionBetween</code>(1)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlFormatConditionType</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CellValue</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ColorScale</code>(3) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Databar</code>(4) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Top10</code>(5) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlIconSet</code>(6) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">UniqueValues</code>(8) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">AboveAverageCondition</code>(12)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlConditionValueType</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypeMinimum</code>(5) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypeMaximum</code>(6) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypePercentile</code>(4) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ConditionValueTypeNumeric</code>(1)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlAboveBelow</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlAboveAverage</code>(0) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlBelowAverage</code>(2) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlAboveStdDev</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlBelowStdDev</code>(3)</td></tr>
  </tbody></table></div>
<h3 id="永続化" class="text-xl font-bold mt-8 mb-3">永続化</h3>
<p class="mb-4 text-slate-700 leading-7">条件付き書式は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.save()</code> で .xlsx に保存され、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.openWorkBook()</code> で開き直しても <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionTypes()</code> や各読み戻しメソッドで復元できます（実測済み）。</p>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">apply()</code> を忘れない。</strong> 忘れても例外にはならず、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getApplied*</code> 系が空を返すだけなので 気づきにくいのが厄介です（<a href="#apply-の役割" class="text-primary-600 hover:underline">解説</a>参照）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>演算子の getter 名は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getConditionOperetor()</code> です。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Operator</code> ではなく <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Operetor</code></strong>（ライブラリ側の綴りそのまま）なので注意してください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>データバーの読み戻しはメソッド名がカラースケールと違います。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DataBarObject</code> では <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumCondition()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumCondition()</code> が <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlConditionValueType</code>（評価条件の種類）</strong>を返し、しきい値は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumValue()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMaximumValue()</code>。<br> 一方 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ColorScaleObject</code> では種類が <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumValueType()</code>、色が <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMinimumColor()</code> です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>色は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#</code> なしの 6 桁 hex</strong>（例 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"9C0006"</code>）で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHexColor()</code> に渡します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Top10Object.setRank(rank, isTop)</code> の第2引数は「上位か下位か」の bool</strong> です （<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> = 上位 N 件、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code> = 下位 N 件）。<br>「Top10」という名前ですが N 件は任意です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>条件オブジェクトには <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setA1C1()</code> で適用範囲を必ず設定します。</strong> 適用範囲は条件オブジェクト側が持ち、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">set〜Condition()</code> には範囲を渡しません。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">**<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getApplied*</code> 系が返すのは「条件を満たして書式が実際に付いたセル」だけ**です。<br> ルールの適用範囲全体ではありません。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">テーブルとコメント</a> / 次: <a href="general-10-charts.php" class="text-primary-600 hover:underline">チャート</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/nodejs/classes/dto.CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/nodejs/classes/dto.ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/nodejs/classes/dto.DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/nodejs/classes/dto.Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/nodejs/enums/enums.XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/nodejs/enums/enums.XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/nodejs/enums/enums.XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/python/pyosbxl.dto.CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/python/pyosbxl.dto.ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/python/pyosbxl.dto.DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/python/pyosbxl.dto.Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/python/pyosbxl.enums.XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/python/pyosbxl.enums.XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/python/pyosbxl.enums.XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/FormatConditionsWrapper.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/CellValueObjectWrapper.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/ColorScaleObjectWrapper.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/DataBarObjectWrapper.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/Top10ObjectWrapper.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/php/classes/phposbxl-dto-CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/php/classes/phposbxl-dto-ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/php/classes/phposbxl-dto-DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/php/classes/phposbxl-dto-Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.FormatConditions.html" class="text-primary-600 hover:underline">FormatConditions</a> / <a href="/api-docs/wasm/classes/dto.CellValueObject.html" class="text-primary-600 hover:underline">CellValueObject</a> / <a href="/api-docs/wasm/classes/dto.ColorScaleObject.html" class="text-primary-600 hover:underline">ColorScaleObject</a> / <a href="/api-docs/wasm/classes/dto.DataBarObject.html" class="text-primary-600 hover:underline">DataBarObject</a> / <a href="/api-docs/wasm/classes/dto.Top10Object.html" class="text-primary-600 hover:underline">Top10Object</a> / <a href="/api-docs/wasm/enums/enums.XlFormatConditionOperator.html" class="text-primary-600 hover:underline">XlFormatConditionOperator</a> / <a href="/api-docs/wasm/enums/enums.XlConditionValueType.html" class="text-primary-600 hover:underline">XlConditionValueType</a> / <a href="/api-docs/wasm/enums/enums.XlAboveBelow.html" class="text-primary-600 hover:underline">XlAboveBelow</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件",
   "description": "条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-09-conditional-format.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-09-conditional-format.php",
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
     "name": "条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件",
     "item": "https://osb-office.com/sample/general-09-conditional-format.php"
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
