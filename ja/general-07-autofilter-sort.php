<?php
$page_title = 'オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し | osbxl サンプル';
$no_tagline = true;
$meta_description = 'オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,autofilter,sort,setcustomfilter,settop10valuefilter,executeSortAscending,executeSortDescending,getSortConditions,hidden,filter';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-07-autofilter-sort.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">データ表に<strong>オートフィルター</strong>を設定して条件に合う行だけを表示し、 あわせて <strong>Sort</strong> で並べ替え（昇順 / 降順）を行う。 フィルターは「条件に合わない行を<strong>非表示にする</strong>」仕組みで、 非表示かどうかは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row.getHidden()</code> で読み取れる。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">やり方</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">使うもの</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">フィルター設定・条件</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getAutoFilter(range, firstRowAsHeader)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">AutoFilter</code> の各 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">set*Filter</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">条件の解除</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.resetAllFilter()</code>（条件のみ） / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.removeFilter()</code>（フィルター自体を削除）</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">並べ替え</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort(range)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortDescending</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">非表示行の判定</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRow(r).getHidden()</code></td></tr>
  </tbody></table></div>
<h2 id="対象クラス" class="text-2xl font-bold mt-10 mb-4">対象クラス</h2>
<h3 id="nodeosbxl-autofilter-ws-getautofilter-a1c1-firstrowasheader-で取得" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl.AutoFilter</code> — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getAutoFilter(A1C1, firstRowAsHeader?)</code> で取得</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAddress()</code></td><td class="border border-surface-200 px-3 py-2 align-top">対象セル範囲を返します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(op, criteria, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">演算子 + 条件式のフィルター。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">criteria</code> に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">*</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">?</code> のワイルドカードが使えます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAndCustomFilter(op1, c1, op2, c2, colId?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setOrCustomFilter(...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">2 条件の AND / OR 結合。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTop10ValueFilter(n, colId?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBottom10ValueFilter(n, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">値の上位 / 下位 n 件。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTop10PercentFilter(p, colId?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBottom10PercentFilter(p, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">値の上位 / 下位 p パーセント。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAverageFilter(aboveAverage, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">平均以上（true）/ 平均以下（false）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesNumberFilter(vals, colId?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">数値リストのいずれかと一致する行だけ表示。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesStringFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesDateTimeFilter</code></td><td class="border border-surface-200 px-3 py-2 align-top">文字列 / 日時版（<a href="#注意点" class="text-primary-600 hover:underline">注意点</a>参照）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateTimeGroupingFilter</code></td><td class="border border-surface-200 px-3 py-2 align-top">日付の期間・グルーピング条件。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFontColorFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCellColorFilter</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setIconFilter</code></td><td class="border border-surface-200 px-3 py-2 align-top">色・アイコンによる条件。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSort()</code></td><td class="border border-surface-200 px-3 py-2 align-top">ヘッダ行を保護した <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sort</code> を返します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetFilter(colId)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllFilter()</code></td><td class="border border-surface-200 px-3 py-2 align-top">条件をクリアします（フィルター自体は残る）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeFilter()</code></td><td class="border border-surface-200 px-3 py-2 align-top">オートフィルター自体を削除し、非表示行も戻します。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">op</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlAutoFilterOperator</code>:</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">数値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">意味</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Equal</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">同一</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">LessThan</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">未満</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">LessThanOrEqual</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">以下</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">NotEqual</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">同一でない</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">GreaterThanOrEqual</code></td><td class="border border-surface-200 px-3 py-2 align-top">5</td><td class="border border-surface-200 px-3 py-2 align-top">以上</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">GreaterThan</code></td><td class="border border-surface-200 px-3 py-2 align-top">6</td><td class="border border-surface-200 px-3 py-2 align-top">より大きい</td></tr>
  </tbody></table></div>
<h3 id="nodeosbxl-sort-ws-getsort-a1c1-firstrowasheader-または-af-getsort-で取得" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl.Sort</code> — <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort(A1C1, firstRowAsHeader?)</code> または <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code> で取得</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending(target, direction?, matchCase?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">昇順に並べ替えます。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">target</code> は範囲内の列番号（先頭列 = 1）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortDescending(target, direction?, matchCase?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">降順に並べ替えます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">execute(sortFieldObject)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.SortFieldObject</code>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setSortOnValues(target, ascending)</code> で条件設定）1 つで並べ替えます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeMultiple([...])</code></td><td class="border border-surface-200 px-3 py-2 align-top">複数条件の並べ替え（2026-10-01 修正により正常動作。<a href="#注意点" class="text-primary-600 hover:underline">注意点</a>参照）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortConditions()</code></td><td class="border border-surface-200 px-3 py-2 align-top">設定済み条件を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.SortFieldObject</code> 配列で読み戻します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetSort(target)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllSort()</code></td><td class="border border-surface-200 px-3 py-2 align-top">条件をクリアします。<strong>AutoFilter / Table 経由の Sort でのみ有効</strong>。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlRowCol</code>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Rows</code> = 1 / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Columns</code> = 2）。既定は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Rows</code>（行を並び替え）。</p>
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
        <pre><code class="language-javascript">// af-sort.js — オートフィルターと並べ替え
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("af-sort.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. データ準備（ヘッダ + 10 行）----------
const DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
];
["商品名", "地域", "数量", "単価"].forEach((h, i) =&gt;
    ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) =&gt; {
    const r = i + 2;
    ws.getRange(`A${r}:A${r}`).setValue(row[0]);
    ws.getRange(`B${r}:B${r}`).setValue(row[1]);
    ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
    ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] データ準備: A1:D11（ヘッダ + 10 行）");

// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
const text = (r, col) =&gt; ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) =&gt; `${text(r, "A")}(${text(r, col)})`;
const visible = (col) =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
    return out.join(" ");
};
const hiddenList = () =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
    return out.length ? out.join(",") : "なし";
};

// ---------- 2. オートフィルター設定 ----------
const af = ws.getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
console.log(`[2] getAutoFilter → 範囲 = ${af.getAddress()} / 非表示行 = ${hiddenList()}`);

// ---------- 3. 条件フィルター: 数量 &gt; 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] 数量&gt;100 の可視行: ${visible("C")}`);
console.log(`    非表示行 = ${hiddenList()} / getValue("A2:A11") の件数 = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter();                            // 条件だけ解除（フィルター自体は残る）

// ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3);                   // 数量の上位 3 件
console.log(`[4] 数量上位3件の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3);   // 数量が 120 または 40
console.log(`[5] 数量∈{120,40} の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["東京", "福岡"], 2);   // 地域が 東京 または 福岡（文字列リスト）
console.log(`[6] 地域∈{東京,福岡} の可視行: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
console.log(`[7] 地域="大阪" の可視行: ${visible("B")}`);
af.resetAllFilter();

// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
ws.getSort("A2:D11").executeSortAscending(3);   // 数量（範囲内3列目）で昇順
console.log(`[8] 数量の昇順: ${visible("C")}`);

// オートフィルター経由の Sort はヘッダ行が自動的に保護される
const sort = af.getSort();
sort.executeSortDescending(4);                  // 単価（4列目）で降順
console.log(`[9] 単価の降順: ${visible("D")}`);

// ---------- 6. 並び替え条件の読み戻し ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] 条件の読み戻し: 対象列 = ${cond.getTarget()} / 昇順 = ${cond.getSortAscending()}`);

// ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter();
console.log(`[11] removeFilter 後: 非表示行 = ${hiddenList()}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># af-sort.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("af-sort-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. データ準備（ヘッダ + 10 行）----------
DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
]
for i, h in enumerate(["商品名", "地域", "数量", "単価"]):
    col = "ABCD"[i]
    ws.getRange(f"{col}1:{col}1").setValue(h)
for i, row in enumerate(DATA):
    r = i + 2
    ws.getRange(f"A{r}:A{r}").setValue(row[0])
    ws.getRange(f"B{r}:B{r}").setValue(row[1])
    ws.getRange(f"C{r}:C{r}").setNumberValue(row[2])
    ws.getRange(f"D{r}:D{r}").setNumberValue(row[3])
print("[1] データ準備: A1:D11（ヘッダ + 10 行）")

# 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
def text(r, col):
    return ws.getRange(f"{col}{r}:{col}{r}").getValue()[f"{col}{r}"]

def label(r, col):
    return f"{text(r, 'A')}({text(r, col)})"

def visible(col):
    out = []
    for r in range(2, 12):
        if not ws.getRow(r).getHidden():
            out.append(label(r, col))
    return " ".join(out)

def hidden_list():
    out = [str(r) for r in range(2, 12) if ws.getRow(r).getHidden()]
    return ",".join(out) if out else "なし"

# ---------- 2. オートフィルター設定 ----------
af = ws.getAutoFilter("A1:D11", True)   # 先頭行はヘッダ
print(f"[2] getAutoFilter → 範囲 = {af.getAddress()} / 非表示行 = {hidden_list()}")

# ---------- 3. 条件フィルター: 数量 &gt; 100 ----------
# colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.GreaterThan, "100", 3)
print(f"[3] 数量&gt;100 の可視行: {visible('C')}")
print(f'    非表示行 = {hidden_list()} / getValue("A2:A11") の件数 = {len(ws.getRange("A2:A11").getValue())}')
af.resetAllFilter()                     # 条件だけ解除（フィルター自体は残る）

# ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3)            # 数量の上位 3 件
print(f"[4] 数量上位3件の可視行: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesNumberFilter([120, 40], 3)   # 数量が 120 または 40
print(f"[5] 数量∈{{120,40}} の可視行: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesStringFilter(["東京", "福岡"], 2)   # 地域が 東京 または 福岡（文字列リスト）
print(f"[6] 地域∈{{東京,福岡}} の可視行: {visible('B')}")
af.resetAllFilter()
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.Equal, "大阪", 2)   # 地域 = 大阪（文字列一致）
print(f'[7] 地域="大阪" の可視行: {visible("B")}')
af.resetAllFilter()

# ---------- 5. 並べ替え（単一キー）----------
# ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = False）
ws.getSort("A2:D11").executeSortAscending(3)   # 数量（範囲内3列目）で昇順
print(f"[8] 数量の昇順: {visible('C')}")

# オートフィルター経由の Sort はヘッダ行が自動的に保護される
sort = af.getSort()
sort.executeSortDescending(4)                  # 単価（4列目）で降順
print(f"[9] 単価の降順: {visible('D')}")

# ---------- 6. 並び替え条件の読み戻し ----------
cond = sort.getSortConditions()[0]
# Python の bool は True/False 表示なので lower() で JS 風に揃える
print(f"[10] 条件の読み戻し: 対象列 = {cond.getTarget()} / 昇順 = {str(cond.getSortAscending()).lower()}")

# ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter()
print(f"[11] removeFilter 後: 非表示行 = {hidden_list()}")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// af-sort.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    static WorkSheetWrapper ws;

    // 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
    static String text(int r, String col) {
        return ws.getRange(col + r + ":" + col + r).getValue().get(col + r);
    }

    static String label(int r, String col) {
        return text(r, "A") + "(" + text(r, col) + ")";
    }

    static String visible(String col) {
        List&lt;String&gt; out = new ArrayList&lt;&gt;();
        for (int r = 2; r &lt;= 11; r++) {
            if (!ws.getRow(r).getHidden()) out.add(label(r, col));
        }
        return String.join(" ", out);
    }

    static String hiddenList() {
        List&lt;String&gt; out = new ArrayList&lt;&gt;();
        for (int r = 2; r &lt;= 11; r++) {
            if (ws.getRow(r).getHidden()) out.add(String.valueOf(r));
        }
        return out.isEmpty() ? "なし" : String.join(",", out);
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("af-sort-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. データ準備（ヘッダ + 10 行）----------
        Object[][] DATA = {
            {"りんご", "東京", 120.0, 100.0},
            {"みかん", "大阪", 40.0, 80.0},
            {"バナナ", "東京", 200.0, 150.0},
            {"ぶどう", "福岡", 60.0, 300.0},
            {"なし", "大阪", 90.0, 120.0},
            {"もも", "福岡", 150.0, 400.0},
            {"キウイ", "東京", 30.0, 90.0},
            {"マンゴー", "大阪", 70.0, 500.0},
            {"いちご", "福岡", 110.0, 350.0},
            {"メロン", "東京", 25.0, 800.0},
        };
        String[] headers = {"商品名", "地域", "数量", "単価"};
        for (int i = 0; i &lt; headers.length; i++) {
            String col = String.valueOf("ABCD".charAt(i));
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        for (int i = 0; i &lt; DATA.length; i++) {
            int r = i + 2;
            ws.getRange("A" + r + ":A" + r).setValue((String) DATA[i][0]);
            ws.getRange("B" + r + ":B" + r).setValue((String) DATA[i][1]);
            ws.getRange("C" + r + ":C" + r).setNumberValue((Double) DATA[i][2]);
            ws.getRange("D" + r + ":D" + r).setNumberValue((Double) DATA[i][3]);
        }
        System.out.println("[1] データ準備: A1:D11（ヘッダ + 10 行）");

        // ---------- 2. オートフィルター設定 ----------
        AutoFilterWrapper af = ws.getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
        System.out.println("[2] getAutoFilter → 範囲 = " + af.getAddress() + " / 非表示行 = " + hiddenList());

        // ---------- 3. 条件フィルター: 数量 &gt; 100 ----------
        // colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
        af.setCustomFilter(XlAutoFilterOperator.GreaterThan, "100", 3);
        System.out.println("[3] 数量&gt;100 の可視行: " + visible("C"));
        System.out.println("    非表示行 = " + hiddenList()
                + " / getValue(\"A2:A11\") の件数 = " + ws.getRange("A2:A11").getValue().size());
        af.resetAllFilter();                            // 条件だけ解除（フィルター自体は残る）

        // ---------- 4. Top-N / 値リストのフィルター ----------
        af.setTop10ValueFilter(3, 3);                   // 数量の上位 3 件
        System.out.println("[4] 数量上位3件の可視行: " + visible("C"));
        af.resetAllFilter();
        af.setUniqueValuesNumberFilter(Arrays.asList(120.0, 40.0), 3);   // 数量が 120 または 40
        System.out.println("[5] 数量∈{120,40} の可視行: " + visible("C"));
        af.resetAllFilter();
        af.setUniqueValuesStringFilter(Arrays.asList("東京", "福岡"), 2);   // 地域が 東京 または 福岡（文字列リスト）
        System.out.println("[6] 地域∈{東京,福岡} の可視行: " + visible("B"));
        af.resetAllFilter();
        af.setCustomFilter(XlAutoFilterOperator.Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
        System.out.println("[7] 地域=\"大阪\" の可視行: " + visible("B"));
        af.resetAllFilter();

        // ---------- 5. 並べ替え（単一キー）----------
        // ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
        ws.getSort("A2:D11").executeSortAscending(3);   // 数量（範囲内3列目）で昇順
        System.out.println("[8] 数量の昇順: " + visible("C"));

        // オートフィルター経由の Sort はヘッダ行が自動的に保護される
        SortWrapper sort = af.getSort();
        sort.executeSortDescending(4);                  // 単価（4列目）で降順
        System.out.println("[9] 単価の降順: " + visible("D"));

        // ---------- 6. 並び替え条件の読み戻し ----------
        SortFieldObjectWrapper cond = sort.getSortConditions().get(0);
        System.out.println("[10] 条件の読み戻し: 対象列 = " + cond.getTarget()
                + " / 昇順 = " + cond.getSortAscending());

        // ---------- 7. オートフィルター自体の削除 ----------
        af.removeFilter();
        System.out.println("[11] removeFilter 後: 非表示行 = " + hiddenList());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// af-sort.php — オートフィルターと並べ替え (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("af-sort-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// ---------- 1. データ準備（ヘッダ + 10 行）----------
$DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
];
foreach (["商品名", "地域", "数量", "単価"] as $i =&gt; $h) {
    $col = "ABCD"[$i];
    $ws-&gt;getRange("{$col}1:{$col}1")-&gt;setValue($h);
}
foreach ($DATA as $i =&gt; $row) {
    $r = $i + 2;
    $ws-&gt;getRange("A{$r}:A{$r}")-&gt;setValue($row[0]);
    $ws-&gt;getRange("B{$r}:B{$r}")-&gt;setValue($row[1]);
    $ws-&gt;getRange("C{$r}:C{$r}")-&gt;setNumberValue($row[2]);
    $ws-&gt;getRange("D{$r}:D{$r}")-&gt;setNumberValue($row[3]);
}
echo "[1] データ準備: A1:D11（ヘッダ + 10 行）\n";

// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
$text = fn(int $r, string $col): string =&gt; $ws-&gt;getRange("{$col}{$r}:{$col}{$r}")-&gt;getValue()["{$col}{$r}"];
$label = fn(int $r, string $col): string =&gt; $text($r, "A") . "(" . $text($r, $col) . ")";
$visible = function (string $col) use ($ws, $label): string {
    $out = [];
    for ($r = 2; $r &lt;= 11; $r++) {
        if (!$ws-&gt;getRow($r)-&gt;getHidden()) {
            $out[] = $label($r, $col);
        }
    }
    return implode(" ", $out);
};
$hiddenList = function () use ($ws): string {
    $out = [];
    for ($r = 2; $r &lt;= 11; $r++) {
        if ($ws-&gt;getRow($r)-&gt;getHidden()) {
            $out[] = (string)$r;
        }
    }
    return count($out) &gt; 0 ? implode(",", $out) : "なし";
};

// ---------- 2. オートフィルター設定 ----------
$af = $ws-&gt;getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
echo "[2] getAutoFilter → 範囲 = ", $af-&gt;getAddress(), " / 非表示行 = ", $hiddenList(), "\n";

// ---------- 3. 条件フィルター: 数量 &gt; 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
$af-&gt;setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::GreaterThan, "100", 3);
echo "[3] 数量&gt;100 の可視行: ", $visible("C"), "\n";
echo "    非表示行 = ", $hiddenList(),
    " / getValue(\"A2:A11\") の件数 = ", count($ws-&gt;getRange("A2:A11")-&gt;getValue()), "\n";
$af-&gt;resetAllFilter();                      // 条件だけ解除（フィルター自体は残る）

// ---------- 4. Top-N / 値リストのフィルター ----------
$af-&gt;setTop10ValueFilter(3, 3);             // 数量の上位 3 件
echo "[4] 数量上位3件の可視行: ", $visible("C"), "\n";
$af-&gt;resetAllFilter();
$af-&gt;setUniqueValuesNumberFilter([120, 40], 3);   // 数量が 120 または 40
echo "[5] 数量∈{120,40} の可視行: ", $visible("C"), "\n";
$af-&gt;resetAllFilter();
$af-&gt;setUniqueValuesStringFilter(["東京", "福岡"], 2);   // 地域が 東京 または 福岡（文字列リスト）
echo "[6] 地域∈{東京,福岡} の可視行: ", $visible("B"), "\n";
$af-&gt;resetAllFilter();
$af-&gt;setCustomFilter(\phposbxl\enums\XlAutoFilterOperator::Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
echo "[7] 地域=\"大阪\" の可視行: ", $visible("B"), "\n";
$af-&gt;resetAllFilter();

// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
$ws-&gt;getSort("A2:D11")-&gt;executeSortAscending(3);   // 数量（範囲内3列目）で昇順
echo "[8] 数量の昇順: ", $visible("C"), "\n";

// オートフィルター経由の Sort はヘッダ行が自動的に保護される
$sort = $af-&gt;getSort();
$sort-&gt;executeSortDescending(4);                  // 単価（4列目）で降順
echo "[9] 単価の降順: ", $visible("D"), "\n";

// ---------- 6. 並び替え条件の読み戻し ----------
$cond = $sort-&gt;getSortConditions()[0];
// PHP の bool は echo すると 1/空文字になるため JS 風の true/false に変換する
echo "[10] 条件の読み戻し: 対象列 = ", $cond-&gt;getTarget(),
    " / 昇順 = ", $cond-&gt;getSortAscending() ? "true" : "false", "\n";

// ---------- 7. オートフィルター自体の削除 ----------
$af-&gt;removeFilter();
echo "[11] removeFilter 後: 非表示行 = ", $hiddenList(), "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// af-sort.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("af-sort-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. データ準備（ヘッダ + 10 行）----------
const DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
];
["商品名", "地域", "数量", "単価"].forEach((h, i) =&gt;
    ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) =&gt; {
    const r = i + 2;
    ws.getRange(`A${r}:A${r}`).setValue(row[0]);
    ws.getRange(`B${r}:B${r}`).setValue(row[1]);
    ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
    ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] データ準備: A1:D11（ヘッダ + 10 行）");

// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
const text = (r, col) =&gt; ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) =&gt; `${text(r, "A")}(${text(r, col)})`;
const visible = (col) =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
    return out.join(" ");
};
const hiddenList = () =&gt; {
    const out = [];
    for (let r = 2; r &lt;= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
    return out.length ? out.join(",") : "なし";
};

// ---------- 2. オートフィルター設定 ----------
const af = ws.getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
console.log(`[2] getAutoFilter → 範囲 = ${af.getAddress()} / 非表示行 = ${hiddenList()}`);

// ---------- 3. 条件フィルター: 数量 &gt; 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] 数量&gt;100 の可視行: ${visible("C")}`);
console.log(`    非表示行 = ${hiddenList()} / getValue("A2:A11") の件数 = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter();                            // 条件だけ解除（フィルター自体は残る）

// ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3);                   // 数量の上位 3 件
console.log(`[4] 数量上位3件の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3);   // 数量が 120 または 40
console.log(`[5] 数量∈{120,40} の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["東京", "福岡"], 2);   // 地域が 東京 または 福岡（文字列リスト）
console.log(`[6] 地域∈{東京,福岡} の可視行: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
console.log(`[7] 地域="大阪" の可視行: ${visible("B")}`);
af.resetAllFilter();

// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
ws.getSort("A2:D11").executeSortAscending(3);   // 数量（範囲内3列目）で昇順
console.log(`[8] 数量の昇順: ${visible("C")}`);

// オートフィルター経由の Sort はヘッダ行が自動的に保護される
const sort = af.getSort();
sort.executeSortDescending(4);                  // 単価（4列目）で降順
console.log(`[9] 単価の降順: ${visible("D")}`);

// ---------- 6. 並び替え条件の読み戻し ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] 条件の読み戻し: 対象列 = ${cond.getTarget()} / 昇順 = ${cond.getSortAscending()}`);

// ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter();
console.log(`[11] removeFilter 後: 非表示行 = ${hiddenList()}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] データ準備: A1:D11（ヘッダ + 10 行）
[2] getAutoFilter → 範囲 = A1:D11 / 非表示行 = なし
[3] 数量&gt;100 の可視行: りんご(120) バナナ(200) もも(150) いちご(110)
    非表示行 = 3,5,6,8,9,11 / getValue("A2:A11") の件数 = 10
[4] 数量上位3件の可視行: りんご(120) バナナ(200) もも(150)
[5] 数量∈{120,40} の可視行: りんご(120) みかん(40)
[6] 地域∈{東京,福岡} の可視行: りんご(東京) バナナ(東京) ぶどう(福岡) もも(福岡) キウイ(東京) いちご(福岡) メロン(東京)
[7] 地域="大阪" の可視行: みかん(大阪) なし(大阪) マンゴー(大阪)
[8] 数量の昇順: メロン(25) キウイ(30) みかん(40) ぶどう(60) マンゴー(70) なし(90) いちご(110) りんご(120) もも(150) バナナ(200)
[9] 単価の降順: メロン(800) マンゴー(500) もも(400) いちご(350) ぶどう(300) バナナ(150) なし(120) りんご(100) キウイ(90) みかん(80)
[10] 条件の読み戻し: 対象列 = 4 / 昇順 = false
[11] removeFilter 後: 非表示行 = なし</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<h3 id="フィルター-行の非表示" class="text-xl font-bold mt-8 mb-3">フィルター = 行の非表示</h3>
<p class="mb-4 text-slate-700 leading-7">オートフィルターは<strong>条件に合わない行を非表示にする</strong>だけで、データは削除されません。</p>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">非表示判定は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRow(r).getHidden()</code>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row</code> の使い方は<a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">行・列の操作</a>参照）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> は非表示行の値もそのまま返します</strong>（実行結果 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[3]</code>: 6 行が非表示でも件数は 10）。<br> 「見えている行だけ」が欲しい場合は、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHidden()</code> で絞り込んでから読んでください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetFilter(colId)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllFilter()</code> は<strong>条件だけ</strong>をクリアし、フィルター（枠）は残します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeFilter()</code> はフィルター自体を削除し、非表示行もすべて表示に戻します（実行結果 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[11]</code>）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">フィルター範囲・条件・非表示状態は <strong>save → 再オープン後も復元</strong>されます（実測）。</li>
</ul>
<h3 id="条件の書き方" class="text-xl font-bold mt-8 mb-3">条件の書き方</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(op, criteria, colId)</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">colId</code> と、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sort</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">target</code> は <strong>どちらも「指定範囲内の列番号」（先頭列 = 1）</strong> です。<br>シートの列番号そのものではないので、 範囲が <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">B</code> 列始まりのときはずれに注意してください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">criteria</code> は<strong>値だけ</strong>を書きます（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"100"</code>）。<br>演算子は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">op</code>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlAutoFilterOperator</code>）で指定します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">数値の大小比較（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">GreaterThan</code> など）、Top-N、平均、数値リスト一致（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesNumberFilter</code>）は 期待どおり動作します（実行結果 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[3]</code>〜<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[5]</code>）。</li>
</ul>
<h3 id="並べ替え" class="text-xl font-bold mt-8 mb-3">並べ替え</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort("A2:D11")</code> のように<strong>データ部だけを範囲にする</strong>のが基本形です。<br> ヘッダ込みの範囲を渡すなら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort("A1:D11", true)</code>（第2引数 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">firstRowAsHeader</code>）か、 <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code>（ヘッダが自動保護される）</strong> を使います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">並べ替えは<strong>行単位</strong>で動きます。キー列以外の列（商品名・地域など）も行と一緒に移動します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4">実行後は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortConditions()</code> で条件を読み戻せます（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTarget()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortAscending()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDirection()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMatchCase()</code>）。<br>読み戻した条件は save → 再オープン後も残ります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending(target, direction, matchCase)</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlRowCol.Rows</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Columns</code>(2)。<br>既定は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Rows</code> です。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">criteria</code> には値のみを渡します（演算子記号・ワイルドカードを付けない）。</strong> 演算子は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlAutoFilterOperator</code> で指定する設計です。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(GreaterThan, "&gt;100", 3)</code> のように記号を混ぜると条件が文字列扱いになり、 数値列では 1 件もマッチせず全行が非表示になります（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"100"</code> なら正常）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>文字列の一致系フィルターは正常に動作します</strong>（2026-10-01 修正。<br>旧バージョン 1.5.0 では セッション内に新規作成した文字列セルが一致せず全行非表示になるバグがありました）。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomFilter(Equal, "東京", ...)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setUniqueValuesStringFilter(["東京","福岡"], ...)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setOrCustomFilter(...)</code> いずれも実行結果 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[6]</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[7]</code> のとおり動作します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeMultiple()</code>（複数キーの並べ替え）は正常に動作します</strong>（2026-10-01 修正）。<br> 旧バージョン 1.5.0 では行の重複・欠落が起きていました。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SortFieldObject.setSortOnValues(列, 昇順?)</code> で条件を作り、優先順に配列で渡します。<br> 受け入れテスト: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">test/sort_af_fc_ext.test.js</code>「Sort executeMultiple (複数キー・A-2 修正)」。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">executeSortAscending()</code> を複数回呼んでも複数キーソートにはなりません。</strong> 後から呼んだキーだけで全体が並べ替わり、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSortConditions()</code> には条件が蓄積されていきます（実測）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>ヘッダ込み範囲を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">firstRowAsHeader</code> なし（既定 false）でソートするとヘッダ行もデータとして並びます。</strong> 実測で、昇順ソートすると 1 行目のヘッダ（文字列）が最下部へ沈みました。<br> ヘッダを含む範囲では <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSort(range, true)</code> または <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">af.getSort()</code> を使ってください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>フィルター適用中にソートすると、可視行だけが並べ替わります。</strong> 非表示行は元の位置に留まるため、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllFilter()</code> してからソートするのが安全です（実測）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetSort()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">resetAllSort()</code> が効くのは AutoFilter / Table 経由の Sort だけです。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getSort()</code> で取得した Sort では条件がクリアされないことを実測で確認しました（ドキュメントどおりの仕様）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAutoFilter(A1C1, firstRowAsHeader?)</code> は「取得」であると同時に範囲の設定でもあります。</strong> オートフィルターはシート単位で 1 範囲だけ。<br>別の範囲で呼び出すと設定範囲が置き換わります。<br> 第2引数は「取得するかしないか」ではなく <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">firstRowAsHeader</code>（先頭行をヘッダとみなすか）です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>数値列は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue()</code>（または数値として解釈される値）で入れてください。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue("120")</code> のような文字列として入れた数値でも大小比較フィルターは動作することを実測しましたが、 Top-N 系などの安定動作のためには数値型での入力を推奨します。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">大量データの一括投入</a> / 次: <a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">テーブルとコメント</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/nodejs/classes/dto.SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/nodejs/enums/enums.XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/nodejs/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/python/pyosbxl.Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/python/pyosbxl.dto.SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/python/pyosbxl.enums.XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/python/pyosbxl.enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/AutoFilterWrapper.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/java/com/osboffice/osbxl/SortWrapper.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/SortFieldObjectWrapper.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/php/classes/phposbxl-Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/php/classes/phposbxl-dto-SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.AutoFilter.html" class="text-primary-600 hover:underline">AutoFilter</a> / <a href="/api-docs/wasm/classes/osbxl.Sort.html" class="text-primary-600 hover:underline">Sort</a> / <a href="/api-docs/wasm/classes/dto.SortFieldObject.html" class="text-primary-600 hover:underline">SortFieldObject</a> / <a href="/api-docs/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/wasm/enums/enums.XlAutoFilterOperator.html" class="text-primary-600 hover:underline">XlAutoFilterOperator</a> / <a href="/api-docs/wasm/enums/enums.XlRowCol.html" class="text-primary-600 hover:underline">XlRowCol</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し",
   "description": "オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-07-autofilter-sort.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-07-autofilter-sort.php",
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
     "name": "オートフィルターと並べ替え — 条件フィルター・ソート・条件の読み戻し",
     "item": "https://osb-office.com/sample/general-07-autofilter-sort.php"
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
