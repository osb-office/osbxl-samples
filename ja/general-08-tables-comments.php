<?php
$page_title = 'テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ | osbxl サンプル';
$no_tagline = true;
$meta_description = 'テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,listobjects,table,addlist,totalsrow,structured-reference,comments,setcomment,setmemo';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ</h1>
<p class="text-sm text-text-light mb-6">level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-08-tables-comments.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">データ範囲を Excel の「テーブル」（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ListObject</code>）にして、名前・列名・スタイル・集計行を扱う。 テーブル列を数式から参照する「構造化参照」を使い、あわせてセルへのコメント（新式スレッド）と メモ（旧式吹き出し）の追加・読み取り・削除を行う。</p>
<h2 id="対象クラス" class="text-2xl font-bold mt-10 mb-4">対象クラス</h2>
<p class="mb-4 text-slate-700 leading-7">テーブルは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getListObjects()</code> から、コメントは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getComments()</code> から取得します。</p>
<h3 id="listobjects-テーブルのコレクション" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ListObjects</code>（テーブルのコレクション）</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList(name, A1C1, useFirstRowAsHeader, insertTotals)</code></td><td class="border border-surface-200 px-3 py-2 align-top">既存のセル範囲からテーブルを作成し、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Table</code> を返す。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addListFromRange(name, topA1, copyFromA1C1, useFirstRowAsHeader, insertTotals)</code></td><td class="border border-surface-200 px-3 py-2 align-top">データを <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">topA1</code> へ<strong>コピーして</strong>テーブルを作成。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyFromA1C1</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet2!A1:C1</code> のように他シートも可。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getList(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">テーブルを取得。無い場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">table is not found</code> 例外。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList(name, deleteData?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">テーブルを削除。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteData</code> 既定は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code>（<strong>データも消える</strong>）。残すなら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code>。</td></tr>
  </tbody></table></div>
<h3 id="table-テーブル-1-つ" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Table</code>（テーブル 1 つ）</h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">分類</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">範囲</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAllRange()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHeaderRowRange()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getDataBodyRange()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowRange()</code>（いずれも <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> を返す。集計行なしで <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowRange()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">totalRow is not set</code> 例外）</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">名前</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">changeTableName(name)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">列名</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColumnName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnName([...])</code>（ヘッダセルの表示も変わる）</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">スタイル</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTableStyleName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBuiltinStyleName(enums.XlDefaultTableStyle.Xxx, clearFormat)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomTableStyleName(name, clearFormat)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">集計行</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowTotals()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setShowTotals(bool)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowFunction(col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowFunction(col, enums.XlTotalsCalculation.Xxx)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomRowFunction(col, formula, isArray?)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowLabel(col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowLabel(col, text)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">フィルター・表示</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAutoFilter()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeAutoFilter()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowAutoFilter()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowHeaders()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setShowHeaders(bool)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isShowTableStyleRowStripes()</code> ほか</td></tr>
  </tbody></table></div>
<h3 id="comments-セルコメント" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Comments</code>（セルコメント）</h3>
<p class="mb-4 text-slate-700 leading-7">Excel にはコメントが 2 系統あります。<strong>新式</strong>（スレッド・返信・完了フラグ）と<strong>旧式メモ</strong>（吹き出し）です。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">系統</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment(A1, commentObject)</code></td><td class="border border-surface-200 px-3 py-2 align-top">新式</td><td class="border border-surface-200 px-3 py-2 align-top">1 件追加。既にコメントがあるセルでは<strong>追加（自動で返信化）</strong>。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread(A1, [commentObject...])</code></td><td class="border border-surface-200 px-3 py-2 align-top">新式</td><td class="border border-surface-200 px-3 py-2 align-top">スレッドを<strong>一括設定（上書き）</strong>。先頭が本体、以降が返信。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComment(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">新式</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Array&lt;dto.CommentObject&gt;</code> を返す。無ければ例外 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">comment not found</code>。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeComment(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">新式</td><td class="border border-surface-200 px-3 py-2 align-top">新式コメントを削除。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMemo(A1, author, text, fontObject, visible?, rowColumnsNum?, colColumnsNum?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">旧式</td><td class="border border-surface-200 px-3 py-2 align-top">吹き出しメモを設定。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">fontObject</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new dto.FontObject()</code> で可。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoText(A1)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoAuthor(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">旧式</td><td class="border border-surface-200 px-3 py-2 align-top">メモ本文 / 作者を取得。無ければ例外 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">memo not found</code>。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeMemo(A1)</code></td><td class="border border-surface-200 px-3 py-2 align-top">旧式</td><td class="border border-surface-200 px-3 py-2 align-top">メモを削除。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CommentObject</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAuthor</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setContent</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setParentId</code>（返信の親ID）/ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDone</code>（完了フラグ）と それぞれの getter を持ちます。ID は設定すると <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{...}</code> 形式の UUID が自動採番されます。</p>
<h3 id="集計方法-enums-xltotalscalculation" class="text-xl font-bold mt-8 mb-3">集計方法 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlTotalsCalculation</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メンバ</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">数値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">意味</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationNone</code></td><td class="border border-surface-200 px-3 py-2 align-top">0</td><td class="border border-surface-200 px-3 py-2 align-top">集計なし</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationSum</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">合計</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationAverage</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">平均</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCount</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">件数</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCountNums</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">数値の個数</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationMin</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationMax</code></td><td class="border border-surface-200 px-3 py-2 align-top">5 / 6</td><td class="border border-surface-200 px-3 py-2 align-top">最小 / 最大</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationStdDev</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationVar</code></td><td class="border border-surface-200 px-3 py-2 align-top">7 / 8</td><td class="border border-surface-200 px-3 py-2 align-top">標準偏差 / 分散</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCustom</code></td><td class="border border-surface-200 px-3 py-2 align-top">9</td><td class="border border-surface-200 px-3 py-2 align-top">カスタム数式（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomRowFunction</code> で設定）</td></tr>
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
        <pre><code class="language-javascript">// tables-comments.js — テーブル(ListObject)作成・集計行・構造化参照・セルコメント
const { nodeosbxl, enums, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("tables-comments.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- データ投入（ヘッダ1行 + データ3行）----------
const rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", 120, 3],
    ["ミカン", 80, 5],
    ["ブドウ", 300, 2],
];
const values = [];
rows.forEach((row, ri) =&gt; row.forEach((cell, ci) =&gt; {
    const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
    const o = new dto.InputValueObject();
    if (typeof cell === "number") o.setNumberValue(a1, cell);
    else o.setStringValue(a1, cell);
    values.push(o);
}));
ws.setValueArray(values);

// ---------- 1. テーブルを作成 ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] テーブル作成");
console.log(`    名前      = ${table.getName()}`);
console.log(`    列名      = ${JSON.stringify(table.getColumnName())}`);
console.log(`    スタイル  = ${table.getTableStyleName()}`);
console.log(`    全範囲    = ${table.getAllRange().getAddress()}`);
console.log(`    ヘッダ    = ${table.getHeaderRowRange().getAddress()}`);
console.log(`    データ部  = ${table.getDataBodyRange().getAddress()}`);
console.log(`    集計行    = ${table.getTotalRowRange().getAddress()}`);

// ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] 集計行");
console.log(`    ラベル(1列目) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(`    数式          = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(`    値            = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);

// ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] スタイル変更後 = ${table.getTableStyleName()}`);

// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
console.log("[4] 構造化参照");
console.log(`    E1 数式 = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(`    E1 値   = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);

// ---------- 5. コメント（新式）とメモ（旧式）----------
const cm = ws.getComments();
const mkComment = (author, content) =&gt; {
    const c = new dto.CommentObject();
    c.setAuthor(author);
    c.setContent(content);
    return c;
};

// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", new dto.FontObject());

const thread = cm.getComment("B2");
console.log("[5] コメント");
console.log(`    B2 スレッド件数 = ${thread.length}`);
thread.forEach((c, i) =&gt; console.log(`      [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (返信)" : ""}`));
console.log(`    D2 コメント件数 = ${cm.getComment("D2").length}`);
console.log(`    A2 メモ = ${JSON.stringify(cm.getMemoText("A2"))} / 作者 ${JSON.stringify(cm.getMemoAuthor("A2"))}`);

// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) =&gt; {
    try { fn(); console.log(`    削除後 ${label}: 例外なし`); }
    catch (e) { console.log(`    削除後 ${label}: 例外 ${e.message}`); }
};
tryRead("D2コメント", () =&gt; cm.getComment("D2"));
tryRead("A2メモ", () =&gt; cm.getMemoText("A2"));
console.log(`    B2 は残すので save 後も ${cm.getComment("B2").length}件`);

const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[確認] テーブル=${table.getName()} 集計(単価)=${total} 構造化参照(E1)=${e1}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># tables-comments.py — node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl


def jdump(obj):  # JSON.stringify 相当の出力（区切り空白なし・非 ASCII そのまま）
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


app = pyosbxl.App()
wb = app.createWorkBook("tables-comments-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- データ投入（ヘッダ1行 + データ3行）----------
rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", 120, 3],
    ["ミカン", 80, 5],
    ["ブドウ", 300, 2],
]
values = []
for ri, row in enumerate(rows):
    for ci, cell in enumerate(row):
        a1 = app.convertFromRowColNumber(ri + 1, ci + 1)
        o = pyosbxl.dto.InputValueObject()
        if isinstance(cell, (int, float)):
            o.setNumberValue(a1, cell)
        else:
            o.setStringValue(a1, cell)
        values.append(o)
ws.setValueArray(values)

# ---------- 1. テーブルを作成 ----------
lo = ws.getListObjects()
table = lo.addList("Sales", "A1:C4", True, True)  # useFirstRowAsHeader, insertTotals
print("[1] テーブル作成")
print(f"    名前      = {table.getName()}")
print(f"    列名      = {jdump(table.getColumnName())}")
print(f"    スタイル  = {table.getTableStyleName()}")
print(f"    全範囲    = {table.getAllRange().getAddress()}")
print(f"    ヘッダ    = {table.getHeaderRowRange().getAddress()}")
print(f"    データ部  = {table.getDataBodyRange().getAddress()}")
print(f"    集計行    = {table.getTotalRowRange().getAddress()}")

# ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計")
table.setTotalRowFunction(2, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
table.setTotalRowFunction(3, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
print("[2] 集計行")
print(f"    ラベル(1列目) = {jdump(table.getTotalRowLabel(1))}")
print(f"    数式          = {jdump(ws.getRange('B5:C5').getFormula())}")
print(f"    値            = {jdump(ws.getRange('B5:C5').getValue(True))}")

# ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(pyosbxl.enums.XlDefaultTableStyle.TableStyleMedium9, True)
print(f"[3] スタイル変更後 = {table.getTableStyleName()}")

# ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])")  # 先頭に = は付けない
print("[4] 構造化参照")
print(f"    E1 数式 = {jdump(ws.getRange('E1:E1').getFormula())}")
print(f"    E1 値   = {jdump(ws.getRange('E1:E1').getValue(True))}")

# ---------- 5. コメント（新式）とメモ（旧式）----------
cm = ws.getComments()


def mkComment(author, content):
    c = pyosbxl.dto.CommentObject()
    c.setAuthor(author)
    c.setContent(content)
    return c


# 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")])
# 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"))
# 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", pyosbxl.dto.FontObject())

thread = cm.getComment("B2")
print("[5] コメント")
print(f"    B2 スレッド件数 = {len(thread)}")
for i, c in enumerate(thread):
    print(f"      [{i}] {c.getAuthor()}: {c.getContent()}{' (返信)' if c.getParentId() else ''}")
print(f"    D2 コメント件数 = {len(cm.getComment('D2'))}")
print(f"    A2 メモ = {jdump(cm.getMemoText('A2'))} / 作者 {jdump(cm.getMemoAuthor('A2'))}")

# 削除: 新式は removeComment / 旧式は removeMemo
cm.removeComment("D2")
cm.removeMemo("A2")


# 削除後の読み取りは node 版と同様に "comment not found" / "memo not found" 例外になる
def tryRead(label, fn):
    try:
        fn()
        print(f"    削除後 {label}: 例外なし")
    except Exception as e:
        print(f"    削除後 {label}: 例外 {e}")


tryRead("D2コメント", lambda: cm.getComment("D2"))
tryRead("A2メモ", lambda: cm.getMemoText("A2"))
print(f"    B2 は残すので save 後も {len(cm.getComment('B2'))}件")

total = ws.getRange("B5:B5").getValue(True)["B5"]
e1 = ws.getRange("E1:E1").getValue(True)["E1"]
print(f"[確認] テーブル={table.getName()} 集計(単価)={total} 構造化参照(E1)={e1}")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// tables-comments.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class Main {
    // JSON.stringify 相当の出力ヘルパー（Map は挿入順のまま）
    static String jsonMap(Map&lt;String, String&gt; m) {
        StringBuilder sb = new StringBuilder("{");
        boolean first = true;
        for (Map.Entry&lt;String, String&gt; e : m.entrySet()) {
            if (!first) sb.append(",");
            first = false;
            sb.append("\"").append(e.getKey()).append("\":\"").append(e.getValue()).append("\"");
        }
        return sb.append("}").toString();
    }

    static String jsonList(List&lt;String&gt; l) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i &lt; l.size(); i++) {
            if (i &gt; 0) sb.append(",");
            sb.append("\"").append(l.get(i)).append("\"");
        }
        return sb.append("]").toString();
    }

    static CommentObjectWrapper mkComment(String author, String content) {
        CommentObjectWrapper c = new CommentObjectWrapper();
        c.setAuthor(author);
        c.setContent(content);
        return c;
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("tables-comments-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- データ投入（ヘッダ1行 + データ3行）----------
        Object[][] rows = {
            {"商品名", "単価", "数量"},
            {"リンゴ", 120, 3},
            {"ミカン", 80, 5},
            {"ブドウ", 300, 2},
        };
        List&lt;InputValueObjectWrapper&gt; values = new ArrayList&lt;&gt;();
        for (int ri = 0; ri &lt; rows.length; ri++) {
            for (int ci = 0; ci &lt; rows[ri].length; ci++) {
                String a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
                InputValueObjectWrapper o = new InputValueObjectWrapper();
                if (rows[ri][ci] instanceof Number) {
                    o.setNumberValue(a1, ((Number) rows[ri][ci]).doubleValue());
                } else {
                    o.setStringValue(a1, (String) rows[ri][ci]);
                }
                values.add(o);
            }
        }
        ws.setValueArray(values);

        // ---------- 1. テーブルを作成 ----------
        ListObjectsWrapper lo = ws.getListObjects();
        TableWrapper table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
        System.out.println("[1] テーブル作成");
        System.out.println("    名前      = " + table.getName());
        System.out.println("    列名      = " + jsonList(table.getColumnName()));
        System.out.println("    スタイル  = " + table.getTableStyleName());
        System.out.println("    全範囲    = " + table.getAllRange().getAddress());
        System.out.println("    ヘッダ    = " + table.getHeaderRowRange().getAddress());
        System.out.println("    データ部  = " + table.getDataBodyRange().getAddress());
        System.out.println("    集計行    = " + table.getTotalRowRange().getAddress());

        // ---------- 2. 集計行 ----------
        table.setTotalRowLabel(1, "合計");
        table.setTotalRowFunction(2, XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(3, XlTotalsCalculation.TotalsCalculationSum);
        System.out.println("[2] 集計行");
        System.out.println("    ラベル(1列目) = \"" + table.getTotalRowLabel(1) + "\"");
        System.out.println("    数式          = " + jsonMap(ws.getRange("B5:C5").getFormula()));
        System.out.println("    値            = " + jsonMap(ws.getRange("B5:C5").getValue(true)));

        // ---------- 3. テーブルスタイルを変更 ----------
        table.setBuiltinStyleName(XlDefaultTableStyle.TableStyleMedium9, true);
        System.out.println("[3] スタイル変更後 = " + table.getTableStyleName());

        // ---------- 4. 構造化参照（テーブル列を参照する数式）----------
        ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
        System.out.println("[4] 構造化参照");
        System.out.println("    E1 数式 = " + jsonMap(ws.getRange("E1:E1").getFormula()));
        System.out.println("    E1 値   = " + jsonMap(ws.getRange("E1:E1").getValue(true)));

        // ---------- 5. コメント（新式）とメモ（旧式）----------
        CommentsWrapper cm = ws.getComments();
        // 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
        cm.setCommentThread("B2", Arrays.asList(
            mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")));
        // 単一の新式コメント → D2（削除の実演用）
        cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
        // 旧式メモ（吹き出し）→ A2
        cm.setMemo("A2", "Carol", "最優先商品です", new FontObjectWrapper());

        List&lt;CommentObjectWrapper&gt; thread = cm.getComment("B2");
        System.out.println("[5] コメント");
        System.out.println("    B2 スレッド件数 = " + thread.size());
        for (int i = 0; i &lt; thread.size(); i++) {
            CommentObjectWrapper c = thread.get(i);
            String pid = c.getParentId();
            boolean isReply = pid != null &amp;&amp; !pid.isEmpty();
            System.out.println("      [" + i + "] " + c.getAuthor() + ": " + c.getContent()
                + (isReply ? " (返信)" : ""));
        }
        System.out.println("    D2 コメント件数 = " + cm.getComment("D2").size());
        System.out.println("    A2 メモ = \"" + cm.getMemoText("A2") + "\" / 作者 \"" + cm.getMemoAuthor("A2") + "\"");

        // 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
        cm.removeComment("D2");
        cm.removeMemo("A2");
        try {
            cm.getComment("D2");
            System.out.println("    削除後 D2コメント: 例外なし");
        } catch (RuntimeException e) {
            System.out.println("    削除後 D2コメント: 例外 " + e.getMessage());
        }
        try {
            cm.getMemoText("A2");
            System.out.println("    削除後 A2メモ: 例外なし");
        } catch (RuntimeException e) {
            System.out.println("    削除後 A2メモ: 例外 " + e.getMessage());
        }
        System.out.println("    B2 は残すので save 後も " + cm.getComment("B2").size() + "件");

        String total = ws.getRange("B5:B5").getValue(true).get("B5");
        String e1 = ws.getRange("E1:E1").getValue(true).get("E1");
        System.out.println("[確認] テーブル=" + table.getName() + " 集計(単価)=" + total + " 構造化参照(E1)=" + e1);

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// tables-comments.php — テーブル(ListObject)作成・集計行・構造化参照・セルコメント (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("tables-comments-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// JSON.stringify 相当の出力（非 ASCII はそのまま）
$jq = fn($x): string =&gt; json_encode($x, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);

// ---------- データ投入（ヘッダ1行 + データ3行）----------
$rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", 120, 3],
    ["ミカン", 80, 5],
    ["ブドウ", 300, 2],
];
$values = [];
foreach ($rows as $ri =&gt; $row) {
    foreach ($row as $ci =&gt; $cell) {
        $a1 = $app-&gt;convertFromRowColNumber($ri + 1, $ci + 1);
        $o = new \phposbxl\dto\InputValueObject();
        if (is_int($cell) || is_float($cell)) {
            $o-&gt;setNumberValue($a1, $cell);
        } else {
            $o-&gt;setStringValue($a1, $cell);
        }
        $values[] = $o;
    }
}
$ws-&gt;setValueArray($values);

// ---------- 1. テーブルを作成 ----------
$lo = $ws-&gt;getListObjects();
$table = $lo-&gt;addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
echo "[1] テーブル作成\n";
echo "    名前      = ", $table-&gt;getName(), "\n";
echo "    列名      = ", $jq($table-&gt;getColumnName()), "\n";
echo "    スタイル  = ", $table-&gt;getTableStyleName(), "\n";
echo "    全範囲    = ", $table-&gt;getAllRange()-&gt;getAddress(), "\n";
echo "    ヘッダ    = ", $table-&gt;getHeaderRowRange()-&gt;getAddress(), "\n";
echo "    データ部  = ", $table-&gt;getDataBodyRange()-&gt;getAddress(), "\n";
echo "    集計行    = ", $table-&gt;getTotalRowRange()-&gt;getAddress(), "\n";

// ---------- 2. 集計行 ----------
$table-&gt;setTotalRowLabel(1, "合計");
$table-&gt;setTotalRowFunction(2, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
$table-&gt;setTotalRowFunction(3, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
echo "[2] 集計行\n";
echo "    ラベル(1列目) = ", $jq($table-&gt;getTotalRowLabel(1)), "\n";
echo "    数式          = ", $jq($ws-&gt;getRange("B5:C5")-&gt;getFormula()), "\n";
echo "    値            = ", $jq($ws-&gt;getRange("B5:C5")-&gt;getValue(true)), "\n";

// ---------- 3. テーブルスタイルを変更 ----------
$table-&gt;setBuiltinStyleName(\phposbxl\enums\XlDefaultTableStyle::TableStyleMedium9, true);
echo "[3] スタイル変更後 = ", $table-&gt;getTableStyleName(), "\n";

// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
$ws-&gt;getRange("E1:E1")-&gt;setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
echo "[4] 構造化参照\n";
echo "    E1 数式 = ", $jq($ws-&gt;getRange("E1:E1")-&gt;getFormula()), "\n";
echo "    E1 値   = ", $jq($ws-&gt;getRange("E1:E1")-&gt;getValue(true)), "\n";

// ---------- 5. コメント（新式）とメモ（旧式）----------
$cm = $ws-&gt;getComments();
$mkComment = function (string $author, string $content): \phposbxl\dto\CommentObject {
    $c = new \phposbxl\dto\CommentObject();
    $c-&gt;setAuthor($author);
    $c-&gt;setContent($content);
    return $c;
};

// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
$cm-&gt;setCommentThread("B2", [$mkComment("alice", "この数値を確認"), $mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
$cm-&gt;setComment("D2", $mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
$cm-&gt;setMemo("A2", "Carol", "最優先商品です", new \phposbxl\dto\FontObject());

$thread = $cm-&gt;getComment("B2");
echo "[5] コメント\n";
echo "    B2 スレッド件数 = ", count($thread), "\n";
foreach ($thread as $i =&gt; $c) {
    echo "      [{$i}] ", $c-&gt;getAuthor(), ": ", $c-&gt;getContent(), $c-&gt;getParentId() ? " (返信)" : "", "\n";
}
echo "    D2 コメント件数 = ", count($cm-&gt;getComment("D2")), "\n";
echo "    A2 メモ = ", $jq($cm-&gt;getMemoText("A2")), " / 作者 ", $jq($cm-&gt;getMemoAuthor("A2")), "\n";

// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
$cm-&gt;removeComment("D2");
$cm-&gt;removeMemo("A2");
$tryRead = function (string $label, callable $fn): void {
    try {
        $fn();
        echo "    削除後 {$label}: 例外なし\n";
    } catch (\Exception $e) {
        echo "    削除後 {$label}: 例外 ", $e-&gt;getMessage(), "\n";
    }
};
$tryRead("D2コメント", fn() =&gt; $cm-&gt;getComment("D2"));
$tryRead("A2メモ", fn() =&gt; $cm-&gt;getMemoText("A2"));
echo "    B2 は残すので save 後も ", count($cm-&gt;getComment("B2")), "件\n";

$total = $ws-&gt;getRange("B5:B5")-&gt;getValue(true)["B5"];
$e1 = $ws-&gt;getRange("E1:E1")-&gt;getValue(true)["E1"];
echo "[確認] テーブル=", $table-&gt;getName(), " 集計(単価)=", $total, " 構造化参照(E1)=", $e1, "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// tables-comments.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("tables-comments-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- データ投入（ヘッダ1行 + データ3行）----------
const rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", 120, 3],
    ["ミカン", 80, 5],
    ["ブドウ", 300, 2],
];
const values = [];
rows.forEach((row, ri) =&gt; row.forEach((cell, ci) =&gt; {
    const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
    const o = new dto.InputValueObject();
    if (typeof cell === "number") o.setNumberValue(a1, cell);
    else o.setStringValue(a1, cell);
    values.push(o);
}));
ws.setValueArray(values);

// ---------- 1. テーブルを作成 ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] テーブル作成");
console.log(`    名前      = ${table.getName()}`);
console.log(`    列名      = ${JSON.stringify(table.getColumnName())}`);
console.log(`    スタイル  = ${table.getTableStyleName()}`);
console.log(`    全範囲    = ${table.getAllRange().getAddress()}`);
console.log(`    ヘッダ    = ${table.getHeaderRowRange().getAddress()}`);
console.log(`    データ部  = ${table.getDataBodyRange().getAddress()}`);
console.log(`    集計行    = ${table.getTotalRowRange().getAddress()}`);

// ---------- 2. 集計行 ----------
table.setTotalRowLabel(1, "合計");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] 集計行");
console.log(`    ラベル(1列目) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(`    数式          = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(`    値            = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);

// ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] スタイル変更後 = ${table.getTableStyleName()}`);

// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
console.log("[4] 構造化参照");
console.log(`    E1 数式 = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(`    E1 値   = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);

// ---------- 5. コメント（新式）とメモ（旧式）----------
const cm = ws.getComments();
const mkComment = (author, content) =&gt; {
    const c = new dto.CommentObject();
    c.setAuthor(author);
    c.setContent(content);
    return c;
};

// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", new dto.FontObject());

const thread = cm.getComment("B2");
console.log("[5] コメント");
console.log(`    B2 スレッド件数 = ${thread.length}`);
thread.forEach((c, i) =&gt; console.log(`      [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (返信)" : ""}`));
console.log(`    D2 コメント件数 = ${cm.getComment("D2").length}`);
console.log(`    A2 メモ = ${JSON.stringify(cm.getMemoText("A2"))} / 作者 ${JSON.stringify(cm.getMemoAuthor("A2"))}`);

// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) =&gt; {
    try { fn(); console.log(`    削除後 ${label}: 例外なし`); }
    catch (e) {
        // embind の例外メッセージは "Comments.getComment: " のような前置が付くため、
        // node 版と同一の出力にするため前置を除去する
        const msg = String(e.message || e).replace(/^[A-Za-z_$][\w$]*\.[\w$]+: /, "");
        console.log(`    削除後 ${label}: 例外 ${msg}`);
    }
};
tryRead("D2コメント", () =&gt; cm.getComment("D2"));
tryRead("A2メモ", () =&gt; cm.getMemoText("A2"));
console.log(`    B2 は残すので save 後も ${cm.getComment("B2").length}件`);

const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[確認] テーブル=${table.getName()} 集計(単価)=${total} 構造化参照(E1)=${e1}`);

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] テーブル作成
    名前      = Sales
    列名      = ["商品名","単価","数量"]
    スタイル  = TableStyleMedium2
    全範囲    = A1:C5
    ヘッダ    = A1:C1
    データ部  = A2:C4
    集計行    = A5:C5
[2] 集計行
    ラベル(1列目) = "合計"
    数式          = {"B5":"SUBTOTAL(109,Sales[単価])","C5":"SUBTOTAL(109,Sales[数量])"}
    値            = {"B5":"500","C5":"10"}
[3] スタイル変更後 = TableStyleMedium9
[4] 構造化参照
    E1 数式 = {"E1":"SUM(Sales[単価])"}
    E1 値   = {"E1":"500"}
[5] コメント
    B2 スレッド件数 = 2
      [0] alice: この数値を確認
      [1] bob: 確認しました (返信)
    D2 コメント件数 = 1
    A2 メモ = "最優先商品です" / 作者 "Carol"
    削除後 D2コメント: 例外 comment not found
    削除後 A2メモ: 例外 memo not found
    B2 は残すので save 後も 2件
[確認] テーブル=Sales 集計(単価)=500 構造化参照(E1)=500</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>テーブルは「範囲」ではなく「名前」で扱う。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList()</code> の戻り値（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Table</code>）や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getList("Sales")</code> から、ヘッダ・データ部・集計行の範囲を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> として取り出せます。<br> 集計行を付けたので <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAllRange()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A1:C5</code>（データ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A1:C4</code> + 集計行 5 行目）になります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>集計行は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL</code> 数式として入る。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowFunction(2, ...Sum)</code> を設定すると、 B5 には <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(109,Sales[単価])</code>（109 = 無視しない合計）が自動生成されます。<br> 平均は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(101,...)</code>、件数は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(103,...)</code> に対応します（実測）。<br> 任意の数式を入れたいときは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomRowFunction(col, "MAX(C2:C4)")</code> を使い、 この場合 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowFunction()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">TotalsCalculationCustom</code>(9) を返します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>構造化参照は「テーブル名[列名]」で書く。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUM(Sales[単価])</code> は 500 になります。<br> 列名だけ（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUM(単価)</code>）では <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#NAME?</code> エラーになります。<br>数式に先頭の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> は付けません。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>列名を変えるとテーブル内の構造化参照も追従する。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnName([...])</code> で列名を変更すると、 集計行や既存数式の中の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sales[旧名]</code> が <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sales[新名]</code> に自動で置き換わります（実測）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>コメントは 2 系統。</strong> スレッド/返信/完了フラグを持つ<strong>新式</strong>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment</code>・<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread</code>・ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComment</code>）と、Excel 旧来の吹き出し<strong>メモ</strong>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMemo</code>・<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoText</code>）があります。<br> 用途に合わせて使い分けます。<br>どちらも <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> すればファイルに永続化されます。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>テーブル名にセル参照と紛らわしい名前は使えません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList("A1", ...)</code> や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList("S2", ...)</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">table name is invalid format</code> 例外になります。<br>同名の再作成は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">table name is already exists</code> 例外です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addList(insertTotals=true)</code> の既定集計は全列「なし」（Excel 準拠、2026-10-01 修正）。</strong> 先頭列にラベル <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">集計</code> だけが自動で入り、データ列には集計関数が入りません （旧バージョンは末尾列だけ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Count</code>（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(103)</code>）が既定で入りました）。<br> 集計したい列は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setTotalRowFunction()</code> で明示します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList()</code> は既定でデータも削除します。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList("Sales")</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteData=true</code> 扱いで セルの値も消えます。<br>テーブル枠だけ外して値を残すには <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">removeList("Sales", false)</code> を使います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>新式コメントは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.CommentObject</code> インスタンスが必須。</strong> ドキュメントの <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">@example</code> にある <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment("A1", { text: "..." })</code> のような<strong>プレーンなリテラルは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Invalid argument</code> 例外</strong>になります（実測）。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new dto.CommentObject()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAuthor()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setContent()</code> で組み立てて渡します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment</code> は追加、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread</code> は上書き。</strong> 既にコメントのあるセルに <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setComment</code> すると 件数が増え、新件は自動的に既存スレッドへの返信（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">parentId</code> 付き）になります。<br> スレッドを丸ごと差し替えたいときは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCommentThread</code> を使います（実測）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>無いコメント/メモを読むと例外になります。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComment()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">comment not found</code>、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoText()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getMemoAuthor()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">memo not found</code> を投げます（2026-10-01 修正）。<br> 削除後の確認は try/catch で例外を捕捉して行います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>集計行の平均などは浮動小数点のまま返ります。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> は表示丸め前の値を文字列で返すため、 平均 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUBTOTAL(101,...)</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">3.3333333333333335</code> のような値になります（表示書式で丸めてください）。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">オートフィルターと並べ替え</a> / 次: <a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">条件付き書式</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/nodejs/classes/dto.CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/nodejs/enums/enums.XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/nodejs/enums/enums.XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/python/pyosbxl.Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/python/pyosbxl.Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/python/pyosbxl.dto.CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/python/pyosbxl.enums.XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/python/pyosbxl.enums.XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/ListObjectsWrapper.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/java/com/osboffice/osbxl/TableWrapper.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/java/com/osboffice/osbxl/CommentsWrapper.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/CommentObjectWrapper.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/php/classes/phposbxl-Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/php/classes/phposbxl-Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/php/classes/phposbxl-dto-CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.ListObjects.html" class="text-primary-600 hover:underline">ListObjects</a> / <a href="/api-docs/wasm/classes/osbxl.Table.html" class="text-primary-600 hover:underline">Table</a> / <a href="/api-docs/wasm/classes/osbxl.Comments.html" class="text-primary-600 hover:underline">Comments</a> / <a href="/api-docs/wasm/classes/dto.CommentObject.html" class="text-primary-600 hover:underline">CommentObject</a> / <a href="/api-docs/wasm/enums/enums.XlTotalsCalculation.html" class="text-primary-600 hover:underline">XlTotalsCalculation</a> / <a href="/api-docs/wasm/enums/enums.XlDefaultTableStyle.html" class="text-primary-600 hover:underline">XlDefaultTableStyle</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ",
   "description": "テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-08-tables-comments.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-08-tables-comments.php",
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
     "name": "テーブル（ListObject）とコメント — 作成・集計行・構造化参照・メモ",
     "item": "https://osb-office.com/sample/general-08-tables-comments.php"
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
