<?php
$page_title = 'ピボットテーブル — 作成・集計・集計方法変更・削除 | osbxl サンプル';
$no_tagline = true;
$meta_description = 'ピボットテーブル — 作成・集計・集計方法変更・削除。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,pivottable,pivottables,addpivottable,setfields,refreshdata,refresh,removepivottable,getfield,getfields,pivotfieldobject,subtotalsmethodtype,layouttype';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">ピボットテーブル — 作成・集計・集計方法変更・削除</h1>
<p class="text-sm text-text-light mb-6">level: advanced ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-11-pivottables.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">ワークシート上のソースデータから<strong>ピボットテーブルを作成</strong>し、 行フィールド・値フィールドを配置して集計する。 あわせて集計方法の変更（合計→平均）とピボットテーブルの削除を行う。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">手順</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">ピボットテーブルの追加</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getPivotTables().addPivotTable(name, sourceRange, destCell)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">フィールド取得</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.getField(fieldName)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">フィールド配置</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.setFields(settings, layoutType, rowFields, colFields, filterFields, dataFields)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">集計の実行</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.refreshData()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.refresh()</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">集計方法の変更</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">field.setDataFieldSubTotalsMethod(...)</code> → 再度 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">ピボットの削除</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getPivotTables().removePivotTable(name)</code></td></tr>
  </tbody></table></div>
<h2 id="対象クラス" class="text-2xl font-bold mt-10 mb-4">対象クラス</h2>
<h3 id="pivottables-worksheet-getpivottables-で取得" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">PivotTables</code> (WorkSheet.getPivotTables() で取得)</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotTable(name, sourceRange, destCell)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ピボットテーブルを追加。sourceRange は A1 形式のセル範囲、destCell は出力先の左上セル (A1 形式)。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTable(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">名前でピボットテーブルを取得。存在しない場合は例外。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removePivotTable(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ピボットテーブルを削除。</td></tr>
  </tbody></table></div>
<h3 id="pivottable" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">PivotTable</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getField(fieldName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ソースデータのヘッダ名でフィールドを取得 (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.PivotFieldObject</code>)。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFields()</code></td><td class="border border-surface-200 px-3 py-2 align-top">全フィールドの配列を返す。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTableSetting()</code></td><td class="border border-surface-200 px-3 py-2 align-top">設定オブジェクト (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.PivotTableSettingObject</code>) を返す。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields</code> の第 1 引数に使う。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields(settings, layout, rows, cols, filters, data)</code></td><td class="border border-surface-200 px-3 py-2 align-top">フィールドを配置してレイアウトを確定する。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code></td><td class="border border-surface-200 px-3 py-2 align-top">ソースデータを再読込して集計し直す。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code></td><td class="border border-surface-200 px-3 py-2 align-top">フィールド構成の変更を反映して再集計する。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getLayoutType()</code></td><td class="border border-surface-200 px-3 py-2 align-top">現在のレイアウト種別 (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlPivotTableLayoutType</code>) を返す。</td></tr>
  </tbody></table></div>
<h3 id="dto-pivotfieldobject" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.PivotFieldObject</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFieldType(type)</code></td><td class="border border-surface-200 px-3 py-2 align-top">フィールドの役割を設定 (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlPivotFieldType</code>)。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataFieldSubTotalsMethod(method)</code></td><td class="border border-surface-200 px-3 py-2 align-top">集計方法を設定 (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlPivotFieldSubtotalsMethodType</code>)。データフィールドのみ有効。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCaption()</code></td><td class="border border-surface-200 px-3 py-2 align-top">フィールド名を返す。</td></tr>
  </tbody></table></div>
<h3 id="主な-enum" class="text-xl font-bold mt-8 mb-3">主な enum</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">enum</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldType.PivotFieldTypeRow</code></td><td class="border border-surface-200 px-3 py-2 align-top">0</td><td class="border border-surface-200 px-3 py-2 align-top">行フィールド</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldType.PivotFieldTypeCol</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">列フィールド</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldType.PivotFieldTypeData</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">データ (値) フィールド</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotTableLayoutType.TabularForm</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">テーブル形式</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotTableLayoutType.CompactForm</code></td><td class="border border-surface-200 px-3 py-2 align-top">0</td><td class="border border-surface-200 px-3 py-2 align-top">コンパクト形式</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotTableLayoutType.OutlineForm</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">アウトライン形式</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">合計</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">平均</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeCount</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">個数</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMax</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">最大</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeMin</code></td><td class="border border-surface-200 px-3 py-2 align-top">5</td><td class="border border-surface-200 px-3 py-2 align-top">最小</td></tr>
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
        <pre><code class="language-javascript">// pivot-sample.js — ピボットテーブル: 作成・集計・集計方法変更・削除
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("pivot.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
const headers = ["地域", "商品", "数量"];
const srcData = [
  ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
  ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
  ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
  ["北", "りんご", 20],
];
headers.forEach((h, i) =&gt; {
  const col = String.fromCharCode(65 + i);
  ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) =&gt; {
  row.forEach((v, ci) =&gt; {
    const col = String.fromCharCode(65 + ci);
    const r = ri + 2;
    if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
    else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
  });
});
const srcTotal = srcData.reduce((s, r) =&gt; s + r[2], 0);
console.log(`[1] ソースデータ: ${srcData.length} 行, 数量の合計 = ${srcTotal}`);

// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
const pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");

const rf = pt.getField("地域");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("数量");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);

pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");

// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
const pivotRows = {};
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    pivotRows[label] = Number(val);
    console.log(`    ${label} = ${val}`);
  }
}
const grandTotal = pivotRows["総計"];
console.log(`[3] 検算: 総計 ${grandTotal} === ソース合計 ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);

// ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] 集計方法を合計→平均に変更");
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    console.log(`    ${label} = ${val}`);
  }
}

// ---------- 5. フィールド情報の取得 ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} 個 → ${fields.map(f =&gt; f.getCaption()).join(", ")}`);
console.log(`    getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);

// ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計");
try {
  ws.getPivotTables().getPivotTable("売上集計");
  console.log("[6] 削除失敗");
} catch (e) {
  console.log("[6] removePivotTable → 削除確認OK");
}

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># pivot-sample.py — ピボットテーブル: node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

enums = pyosbxl.enums

app = pyosbxl.App()
wb = app.createWorkBook("pivot-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. ソースデータ ----------
# 地域 / 商品 / 数量 の 10 行
headers = ["地域", "商品", "数量"]
src_data = [
    ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
    ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
    ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
    ["北", "りんご", 20],
]
for i, h in enumerate(headers):
    col = chr(65 + i)
    ws.getRange(f"{col}1:{col}1").setValue(h)
for ri, row in enumerate(src_data):
    for ci, v in enumerate(row):
        col = chr(65 + ci)
        r = ri + 2
        cell = ws.getRange(f"{col}{r}:{col}{r}")
        if isinstance(v, int):
            cell.setNumberValue(v)
        else:
            cell.setValue(v)
src_total = sum(row[2] for row in src_data)
print(f"[1] ソースデータ: {len(src_data)} 行, 数量の合計 = {src_total}")

# ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1")

rf = pt.getField("地域")
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow)
df = pt.getField("数量")
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData)
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum)

pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)")

# ---------- 3. 集計結果の読み戻しと検算 ----------
# ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
pivot_rows = {}
for r in range(2, 11):
    label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
    val   = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
    if label is not None and label != "":
        pivot_rows[label] = float(val)
        print(f"    {label} = {val}")
grand_total = pivot_rows["総計"]
# node の Number 表示（205）に合わせるため、整数なら int で表示する
gt = int(grand_total) if grand_total == int(grand_total) else grand_total
print(f"[3] 検算: 総計 {gt} === ソース合計 {src_total} → {'OK' if grand_total == src_total else 'NG'}")

# ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[4] 集計方法を合計→平均に変更")
for r in range(2, 11):
    label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
    val   = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
    if label is not None and label != "":
        print(f"    {label} = {val}")

# ---------- 5. フィールド情報の取得 ----------
fields = pt.getFields()
print(f"[5] getFields: {len(fields)} 個 → " + ", ".join(f.getCaption() for f in fields))
print(f"    getLayoutType = {int(pt.getLayoutType())} (TabularForm={int(enums.XlPivotTableLayoutType.TabularForm)})")

# ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計")
try:
    ws.getPivotTables().getPivotTable("売上集計")
    print("[6] 削除失敗")
except Exception:
    print("[6] removePivotTable → 削除確認OK")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// pivot-sample.java — ピボットテーブル: node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("pivot-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. ソースデータ ----------
        // 地域 / 商品 / 数量 の 10 行
        String[] headers = {"地域", "商品", "数量"};
        String[][] srcData = {
            {"東", "りんご", "10"}, {"東", "みかん", "20"}, {"西", "りんご", "30"},
            {"西", "みかん", "15"}, {"東", "りんご", "25"}, {"北", "みかん", "40"},
            {"北", "りんご",  "5"}, {"西", "りんご", "10"}, {"東", "みかん", "30"},
            {"北", "りんご", "20"},
        };
        for (int i = 0; i &lt; headers.length; i++) {
            String col = String.valueOf((char) ('A' + i));
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        int srcTotal = 0;
        for (int ri = 0; ri &lt; srcData.length; ri++) {
            for (int ci = 0; ci &lt; srcData[ri].length; ci++) {
                String col = String.valueOf((char) ('A' + ci));
                int r = ri + 2;
                String addr = col + r + ":" + col + r;
                if (ci == 2) {   // 数量列は数値として入力
                    int v = Integer.parseInt(srcData[ri][ci]);
                    ws.getRange(addr).setNumberValue(v);
                    srcTotal += v;
                } else {
                    ws.getRange(addr).setValue(srcData[ri][ci]);
                }
            }
        }
        System.out.println("[1] ソースデータ: " + srcData.length + " 行, 数量の合計 = " + srcTotal);

        // ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
        PivotTableWrapper pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");

        PivotFieldObjectWrapper rf = pt.getField("地域");
        rf.setFieldType(XlPivotFieldType.PivotFieldTypeRow);
        PivotFieldObjectWrapper df = pt.getField("数量");
        df.setFieldType(XlPivotFieldType.PivotFieldTypeData);
        df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);

        List&lt;PivotFieldObjectWrapper&gt; empty = new ArrayList&lt;&gt;();
        pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
                     Arrays.asList(rf), empty, empty, Arrays.asList(df));
        pt.refreshData();
        pt.refresh();
        System.out.println("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");

        // ---------- 3. 集計結果の読み戻しと検算 ----------
        // ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
        Map&lt;String, Double&gt; pivotRows = new LinkedHashMap&lt;&gt;();
        for (int r = 2; r &lt;= 10; r++) {
            String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
            String val   = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
            if (label != null &amp;&amp; !label.isEmpty()) {
                pivotRows.put(label, Double.parseDouble(val));
                System.out.println("    " + label + " = " + val);
            }
        }
        double grandTotal = pivotRows.get("総計");
        // node の Number 表示（205）に合わせるため、整数なら long で表示する
        String gt = (grandTotal == Math.rint(grandTotal))
            ? String.valueOf((long) grandTotal) : String.valueOf(grandTotal);
        System.out.println("[3] 検算: 総計 " + gt + " === ソース合計 " + srcTotal
            + " → " + (grandTotal == srcTotal ? "OK" : "NG"));

        // ---------- 4. 集計方法を平均に変更 ----------
        df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
        pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
                     Arrays.asList(rf), empty, empty, Arrays.asList(df));
        pt.refreshData();
        pt.refresh();
        System.out.println("[4] 集計方法を合計→平均に変更");
        for (int r = 2; r &lt;= 10; r++) {
            String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
            String val   = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
            if (label != null &amp;&amp; !label.isEmpty()) {
                System.out.println("    " + label + " = " + val);
            }
        }

        // ---------- 5. フィールド情報の取得 ----------
        List&lt;PivotFieldObjectWrapper&gt; fields = pt.getFields();
        StringBuilder caps = new StringBuilder();
        for (int i = 0; i &lt; fields.size(); i++) {
            if (i &gt; 0) caps.append(", ");
            caps.append(fields.get(i).getCaption());
        }
        System.out.println("[5] getFields: " + fields.size() + " 個 → " + caps);
        System.out.println("    getLayoutType = " + pt.getLayoutType().getCode()
            + " (TabularForm=" + XlPivotTableLayoutType.TabularForm.getCode() + ")");

        // ---------- 6. ピボットテーブルの削除 ----------
        ws.getPivotTables().removePivotTable("売上集計");
        try {
            ws.getPivotTables().getPivotTable("売上集計");
            System.out.println("[6] 削除失敗");
        } catch (Throwable t) {
            System.out.println("[6] removePivotTable → 削除確認OK");
        }

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// pivot-sample.php — ピボットテーブル: node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("pivot-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
$headers = ["地域", "商品", "数量"];
$srcData = [
  ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
  ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
  ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
  ["北", "りんご", 20],
];
foreach ($headers as $i =&gt; $h) {
  $col = chr(65 + $i);
  $ws-&gt;getRange("{$col}1:{$col}1")-&gt;setValue($h);
}
$srcTotal = 0;
foreach ($srcData as $ri =&gt; $row) {
  foreach ($row as $ci =&gt; $v) {
    $col = chr(65 + $ci);
    $r = $ri + 2;
    if (is_int($v)) $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;setNumberValue($v);
    else $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;setValue($v);
  }
  $srcTotal += $row[2];
}
echo "[1] ソースデータ: " . count($srcData) . " 行, 数量の合計 = {$srcTotal}\n";

// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
$pt = $ws-&gt;getPivotTables()-&gt;addPivotTable("売上集計", "A1:C11", "E1");

$rf = $pt-&gt;getField("地域");
$rf-&gt;setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeRow);
$df = $pt-&gt;getField("数量");
$df-&gt;setFieldType(\phposbxl\enums\XlPivotFieldType::PivotFieldTypeData);
$df-&gt;setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeSum);

$pt-&gt;setFields($pt-&gt;getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
               [$rf], [], [], [$df]);
$pt-&gt;refreshData();
$pt-&gt;refresh();
echo "[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)\n";

// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
// getValue() は連想配列（キー=セルアドレス、値=表示値の文字列）を返す
$pivotRows = [];
for ($r = 2; $r &lt;= 10; $r++) {
  $label = $ws-&gt;getRange("E{$r}:E{$r}")-&gt;getValue()["E{$r}"] ?? null;
  $val   = $ws-&gt;getRange("F{$r}:F{$r}")-&gt;getValue()["F{$r}"] ?? null;
  if ($label !== null &amp;&amp; $label !== "") {
    $pivotRows[$label] = (float)$val;
    echo "    {$label} = {$val}\n";
  }
}
$grandTotal = $pivotRows["総計"];
// node の Number 表示（205）に合わせるため、整数なら int で表示する
$gt = ($grandTotal == (int)$grandTotal) ? (string)(int)$grandTotal : (string)$grandTotal;
echo "[3] 検算: 総計 {$gt} === ソース合計 {$srcTotal} → " . ($grandTotal == $srcTotal ? "OK" : "NG") . "\n";

// ---------- 4. 集計方法を平均に変更 ----------
$df-&gt;setDataFieldSubTotalsMethod(\phposbxl\enums\XlPivotFieldSubtotalsMethodType::SubtotalsMethodTypeAverage);
$pt-&gt;setFields($pt-&gt;getPivotTableSetting(), \phposbxl\enums\XlPivotTableLayoutType::TabularForm,
               [$rf], [], [], [$df]);
$pt-&gt;refreshData();
$pt-&gt;refresh();
echo "[4] 集計方法を合計→平均に変更\n";
for ($r = 2; $r &lt;= 10; $r++) {
  $label = $ws-&gt;getRange("E{$r}:E{$r}")-&gt;getValue()["E{$r}"] ?? null;
  $val   = $ws-&gt;getRange("F{$r}:F{$r}")-&gt;getValue()["F{$r}"] ?? null;
  if ($label !== null &amp;&amp; $label !== "") {
    echo "    {$label} = {$val}\n";
  }
}

// ---------- 5. フィールド情報の取得 ----------
$fields = $pt-&gt;getFields();
$caps = array_map(fn($f) =&gt; $f-&gt;getCaption(), $fields);
echo "[5] getFields: " . count($fields) . " 個 → " . implode(", ", $caps) . "\n";
$tabular = \phposbxl\enums\XlPivotTableLayoutType::TabularForm-&gt;value;
echo "    getLayoutType = {$pt-&gt;getLayoutType()} (TabularForm={$tabular})\n";

// ---------- 6. ピボットテーブルの削除 ----------
$ws-&gt;getPivotTables()-&gt;removePivotTable("売上集計");
try {
  $ws-&gt;getPivotTables()-&gt;getPivotTable("売上集計");
  echo "[6] 削除失敗\n";
} catch (\Throwable $e) {
  echo "[6] removePivotTable → 削除確認OK\n";
}

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// pivot-sample.wasm.js — ピボットテーブル: node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("pivot-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
const headers = ["地域", "商品", "数量"];
const srcData = [
  ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
  ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
  ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
  ["北", "りんご", 20],
];
headers.forEach((h, i) =&gt; {
  const col = String.fromCharCode(65 + i);
  ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) =&gt; {
  row.forEach((v, ci) =&gt; {
    const col = String.fromCharCode(65 + ci);
    const r = ri + 2;
    if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
    else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
  });
});
const srcTotal = srcData.reduce((s, r) =&gt; s + r[2], 0);
console.log(`[1] ソースデータ: ${srcData.length} 行, 数量の合計 = ${srcTotal}`);

// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
const pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");

const rf = pt.getField("地域");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("数量");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);

pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");

// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
const pivotRows = {};
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    pivotRows[label] = Number(val);
    console.log(`    ${label} = ${val}`);
  }
}
const grandTotal = pivotRows["総計"];
console.log(`[3] 検算: 総計 ${grandTotal} === ソース合計 ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);

// ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] 集計方法を合計→平均に変更");
for (let r = 2; r &lt;= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined &amp;&amp; label !== "") {
    console.log(`    ${label} = ${val}`);
  }
}

// ---------- 5. フィールド情報の取得 ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} 個 → ${fields.map(f =&gt; f.getCaption()).join(", ")}`);
console.log(`    getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);

// ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計");
try {
  ws.getPivotTables().getPivotTable("売上集計");
  console.log("[6] 削除失敗");
} catch (e) {
  console.log("[6] removePivotTable → 削除確認OK");
}

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] ソースデータ: 10 行, 数量の合計 = 205
[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)
    北 = 65
    東 = 85
    西 = 55
    総計 = 205
[3] 検算: 総計 205 === ソース合計 205 → OK
[4] 集計方法を合計→平均に変更
    北 = 21.66666667
    東 = 21.25
    西 = 18.33333333
    総計 = 20.5
[5] getFields: 3 個 → 地域, 商品, 数量
    getLayoutType = 2 (TabularForm=2)
[6] removePivotTable → 削除確認OK</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<h3 id="ピボットテーブル作成の流れ" class="text-xl font-bold mt-8 mb-3">ピボットテーブル作成の流れ</h3>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>ソースデータを用意する</strong> — 1 行目にヘッダ（フィールド名）、2 行目以降にデータ。</li>
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotTable(name, sourceRange, destCell)</code></strong> — ソース範囲 (A1 形式) と出力先の左上セルを指定。</li>
</ol>
<p class="mb-4 text-slate-700 leading-7">同じシート内に配置する場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">destCell</code> にソースと重ならないセルを指定します。 別シートに配置する場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">sourceRange</code> に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet1!A1:C11</code> のようにシート名を付けます。</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong>フィールドを取得して役割を設定</strong> — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pt.getField("ヘッダ名")</code> でフィールドを取得し、</li>
</ol>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFieldType()</code> で 行 / 列 / データ / フィルターのいずれかを割り当てます。</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> で配置</strong> — 設定オブジェクト、レイアウト種別、各フィールド配列を渡します。</li>
</ol>
<p class="mb-4 text-slate-700 leading-7">引数の順序は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(settings, layoutType, rowFields, colFields, filterFields, dataFields)</code> です。</p>
<ol class="list-decimal pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code></strong> — ソースデータを再読込し、集計結果をシートに書き出します。</li>
</ol>
<p class="mb-4 text-slate-700 leading-7"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> だけでは結果は反映されません。</strong> 必ず <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code> を呼んでください。</p>
<h3 id="集計方法の変更" class="text-xl font-bold mt-8 mb-3">集計方法の変更</h3>
<p class="mb-4 text-slate-700 leading-7">データフィールドの集計方法は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataFieldSubTotalsMethod()</code> で変更できます。 変更後は <strong>再度 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code></strong> が必要です。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">enum 値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">数値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">集計方法</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeSum</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">合計 (既定)</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeCount</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">個数</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeAverage</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">平均</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeMax</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">最大</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeMin</code></td><td class="border border-surface-200 px-3 py-2 align-top">5</td><td class="border border-surface-200 px-3 py-2 align-top">最小</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeProduct</code></td><td class="border border-surface-200 px-3 py-2 align-top">6</td><td class="border border-surface-200 px-3 py-2 align-top">積</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeCountNums</code></td><td class="border border-surface-200 px-3 py-2 align-top">7</td><td class="border border-surface-200 px-3 py-2 align-top">数値の個数</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeStdDev</code></td><td class="border border-surface-200 px-3 py-2 align-top">8</td><td class="border border-surface-200 px-3 py-2 align-top">標本標準偏差</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeStdDevP</code></td><td class="border border-surface-200 px-3 py-2 align-top">9</td><td class="border border-surface-200 px-3 py-2 align-top">標準偏差</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeVar</code></td><td class="border border-surface-200 px-3 py-2 align-top">10</td><td class="border border-surface-200 px-3 py-2 align-top">標本分散</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SubtotalsMethodTypeVarP</code></td><td class="border border-surface-200 px-3 py-2 align-top">11</td><td class="border border-surface-200 px-3 py-2 align-top">分散</td></tr>
  </tbody></table></div>
<h3 id="行-列のクロス集計" class="text-xl font-bold mt-8 mb-3">行 × 列のクロス集計</h3>
<p class="mb-4 text-slate-700 leading-7">列フィールドも使う場合は、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colFields</code> に配列を渡します。</p>
<p class="mb-4 text-slate-700 leading-7">出力は行ラベル (E 列) × 列ラベル (F 列以降) のマトリクスになります。</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
  </div>
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">const cf = pt.getField("商品");
cf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeCol);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [cf], [], [df]);
pt.refreshData();
pt.refresh();</code></pre>
      </div>
    </div>
</div>
<h3 id="集計結果の読み戻し" class="text-xl font-bold mt-8 mb-3">集計結果の読み戻し</h3>
<p class="mb-4 text-slate-700 leading-7">ピボットテーブルの出力は通常のセル値として読み取れます。 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("F3:F3").getValue()</code> のようにして集計値を取得し、 ソースデータの合計と比較することで検算できます。</p>
<h3 id="ピボットテーブルの削除" class="text-xl font-bold mt-8 mb-3">ピボットテーブルの削除</h3>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removePivotTable(name)</code> で削除します。削除後に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTable(name)</code> を呼ぶと例外が投げられます。</p>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> の後に必ず <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code> を呼ぶ。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> は配置を登録するだけで、シートへの書き出しは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refresh()</code> が行います。<br> ソースデータを変更した場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">refreshData()</code> を先に呼びます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>列フィールドなしのピボット (行 + 値のみ) も動作します。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colFields</code> に空配列 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[]</code> を渡します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFields()</code> はソースデータの全ヘッダを返します。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFields()</code> で配置しなかったフィールドも含まれるため、 配置済みフィールドだけを取得したい場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFieldType()</code> で確認してください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotTable()</code> のソース範囲にはヘッダ行を含めます。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A1:C11</code> のように 1 行目をヘッダとした範囲を指定します。<br> ヘッダ名がそのままフィールド名になります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>集計結果の「総計」行は日本語ロケールで <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"総計"</code> というラベルが付きます。</strong> ラベル文字列は環境に依存する可能性があるため、 読み戻す際は行番号でアクセスするか、ラベルをハードコードせず検索で探すことをおすすめします。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>同じソース範囲に対して複数のピボットテーブルを作る場合、</strong> 名前 (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pivotTableName</code>) を変えれば同一シート内に複数配置できます。<br> ただし出力先セルが重ならないようにしてください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removePivotTable()</code> で削除しても、ピボットが書き出したセルの値は残ります。</strong> 値も消したい場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCell()</code> 等で別途クリアしてください。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-10-charts.php" class="text-primary-600 hover:underline">チャート</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/nodejs/classes/dto.PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/nodejs/enums/enums.XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/nodejs/enums/enums.XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/python/pyosbxl.PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/python/pyosbxl.dto.PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/python/pyosbxl.enums.XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/python/pyosbxl.enums.XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/PivotTableWrapper.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/java/com/osboffice/osbxl/PivotTablesWrapper.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/PivotFieldObjectWrapper.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/php/classes/phposbxl-PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/php/classes/phposbxl-dto-PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.PivotTable.html" class="text-primary-600 hover:underline">PivotTable</a> / <a href="/api-docs/wasm/classes/osbxl.PivotTables.html" class="text-primary-600 hover:underline">PivotTables</a> / <a href="/api-docs/wasm/classes/dto.PivotFieldObject.html" class="text-primary-600 hover:underline">PivotFieldObject</a> / <a href="/api-docs/wasm/enums/enums.XlPivotFieldType.html" class="text-primary-600 hover:underline">XlPivotFieldType</a> / <a href="/api-docs/wasm/enums/enums.XlPivotFieldSubtotalsMethodType.html" class="text-primary-600 hover:underline">XlPivotFieldSubtotalsMethodType</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "ピボットテーブル — 作成・集計・集計方法変更・削除",
   "description": "ピボットテーブル — 作成・集計・集計方法変更・削除。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-11-pivottables.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-11-pivottables.php",
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
     "name": "ピボットテーブル — 作成・集計・集計方法変更・削除",
     "item": "https://osb-office.com/sample/general-11-pivottables.php"
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
