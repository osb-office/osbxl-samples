<?php
$page_title = 'サンプル集 | osbxl';
$no_tagline = true;
$meta_description = '実行検証済みの osbxl サンプル集。Node.js / Python / Java / PHP / WebAssembly の5言語タブ切り替え対応。';
$meta_keywords = 'osbxl,サンプル,サンプル集,Excel,Excel生成ライブラリ,Excel入出力ライブラリ,Node.js,Python,Java,PHP,WebAssembly';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<p class="text-sm mb-4 text-right"><a href="/en/sample/" class="text-primary-600 hover:underline">English</a></p><h1 id="osbxl-サンプル集" class="text-3xl font-bold mb-2 mt-2">osbxl サンプル集</h1>
<p class="mb-4 text-slate-700 leading-7">実行して確認済みの osbxl サンプルです。各ページのコードはそのまま動きます。 コードは <strong>Node.js / Python / Java / PHP / WebAssembly の5言語タブ</strong>で切り替えられ、 タブごとに表示されるインストール手順も言語別です。</p>
<p class="mb-4 text-slate-700 leading-7"><strong>はじめての方は <a href="general-01-overview.php" class="text-primary-600 hover:underline">general-01 概要 — オブジェクトモデルと基本の流れ</a> から。</strong></p>
<h2 id="やりたいことから引く" class="text-2xl font-bold mt-10 mb-4">やりたいことから引く</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">やりたいこと</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">サンプル</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">osbxl の全体像・クラスの階層を知りたい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php" class="text-primary-600 hover:underline">general-01 概要</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">新規作成 / 既存ブックを開く / save / saveAs / close の流れ</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php#基本の流れ" class="text-primary-600 hover:underline">general-01 概要 — 基本の流れ</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">App / WorkBook / WorkSheet / Range にはどんなメソッドがあるか</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php#主要メソッド" class="text-primary-600 hover:underline">general-01 概要 — 主要メソッド</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">表を作って値と数式を入れたい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">general-02 値と数式</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">行ごとにずれる数式（相対参照）を一括で入れたい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-02-values-formulas.php#コード" class="text-primary-600 hover:underline">general-02 値と数式</a> / <a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">general-06 一括投入</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">シートを追加・削除・移動・リネームしたい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-03-sheets.php" class="text-primary-600 hover:underline">general-03 シート操作</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">シート名が分からないブックの全シートを列挙したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-03-sheets.php#コード" class="text-primary-600 hover:underline">general-03 シート操作</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">日付・時刻・真偽値を入れたい / 表示書式を指定したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-04-value-types.php" class="text-primary-600 hover:underline">general-04 セル値の型</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> の違い</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-04-value-types.php#getvalue-と-getvaluetrue" class="text-primary-600 hover:underline">general-04 セル値の型</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">行や列を挿入・削除・クリア・コピーしたい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">general-05 行・列・セル範囲</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">行の高さ・列の幅・非表示を設定したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-05-rows-cols.php#コード-2-row-col-オブジェクト" class="text-primary-600 hover:underline">general-05 行・列・セル範囲</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">数万セルを高速に投入したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">general-06 一括投入</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">データを絞り込む・並べ替えたい（オートフィルター / Sort）</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">general-07 オートフィルターと並べ替え</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">テーブル（集計行・構造化参照）やコメントを使いたい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">general-08 テーブルとコメント</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">条件付き書式（カラースケール・データバー・上位 N）を設定したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">general-09 条件付き書式</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">グラフ（縦棒・折れ線）を作成したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-10-charts.php" class="text-primary-600 hover:underline">general-10 チャート</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">ピボットテーブルで集計したい</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-11-pivottables.php" class="text-primary-600 hover:underline">general-11 ピボットテーブル</a></td></tr>
  </tbody></table></div>
<h2 id="general-概要・基本操作" class="text-2xl font-bold mt-10 mb-4">general — 概要・基本操作</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">#</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">タイトル</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">level</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">主な API</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">01</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php" class="text-primary-600 hover:underline">概要 — オブジェクトモデルと基本の流れ</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">App / WorkBook / WorkSheet / Range</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">02</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">値と数式を入れて表を作る</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">Range / InputFormulaObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">03</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-03-sheets.php" class="text-primary-600 hover:underline">ブックとワークシートの操作 — 追加・削除・移動・リネーム</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook / WorkSheet</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">04</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-04-value-types.php" class="text-primary-600 hover:underline">セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">Range / DateTimeObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">05</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">行・列・セル範囲の操作 — 挿入・削除・クリア・コピー</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet / Row / Col</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">06</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">大量データの一括投入 — setValueArray / setFormulaArray</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet / InputValueObject / InputFormulaObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">07</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">オートフィルターと並べ替え — 条件フィルター・ソート</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">AutoFilter / Sort</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">08</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">テーブルとコメント — ListObjects・集計行・構造化参照</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">ListObjects / Table / Comments</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">09</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">条件付き書式 — セル値ルール・カラースケール・データバー</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">FormatConditions / ColorScaleObject / DataBarObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">10</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-10-charts.php" class="text-primary-600 hover:underline">チャートの基本 — 縦棒・折れ線・タイトル・軸・凡例</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">Chart / chart.* / enums.XlChartType</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">11</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-11-pivottables.php" class="text-primary-600 hover:underline">ピボットテーブル — 作成・集計・集計方法の変更</a></td><td class="border border-surface-200 px-3 py-2 align-top">advanced</td><td class="border border-surface-200 px-3 py-2 align-top">PivotTable / PivotField</td></tr>
  </tbody></table></div>
<h2 id="予定-未作成" class="text-2xl font-bold mt-10 mb-4">予定（未作成）</h2>
<h3 id="general" class="text-xl font-bold mt-8 mb-3">general</h3>
<p class="mb-4 text-slate-700 leading-7">ブックの属性（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCompanyName</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCreateAuthor</code> 他） / 名前付き範囲（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNames</code>） / カスタムスタイル（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomStyle</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomTableStyle</code>） / 外部ブック連携（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addExternalWorkBook</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">updateExternalWorkBook</code>） / パスワード付きブック（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openPasswordWorkBook</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save(password)</code>） / アウトラインの詳細</p>
<h3 id="format" class="text-xl font-bold mt-8 mb-3">format</h3>
<p class="mb-4 text-slate-700 leading-7">フォント・色 / 罫線 / 塗りつぶし・グラデーション / 配置と折り返し / 表示形式（数値・日付・通貨） / 組み込みスタイル・カスタムスタイル （条件付き書式は <a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">general-09</a> 参照）</p>
<h3 id="data" class="text-xl font-bold mt-8 mb-3">data</h3>
<p class="mb-4 text-slate-700 leading-7">セル結合 / ハイパーリンク （テーブル・並べ替え・オートフィルター・コメントは <a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">general-07</a> / <a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">general-08</a> 参照）</p>
<h3 id="formula" class="text-xl font-bold mt-8 mb-3">formula</h3>
<p class="mb-4 text-slate-700 leading-7">数式の基本と配列数式 / 財務関数（ACCINT 他） / 検索関数（XLOOKUP） / 予測（FORECAST.ETS） / WEBSERVICE + FILTERXML</p>
<h3 id="pivot" class="text-xl font-bold mt-8 mb-3">pivot</h3>
<p class="mb-4 text-slate-700 leading-7">集計とグループ化（作成・集計の基本は <a href="general-11-pivottables.php" class="text-primary-600 hover:underline">general-11</a> 参照）</p>
<h3 id="chart" class="text-xl font-bold mt-8 mb-3">chart</h3>
<p class="mb-4 text-slate-700 leading-7">グラフの種類 / データラベル / 書式（色・線・影） （基本作成・軸・凡例は <a href="general-10-charts.php" class="text-primary-600 hover:underline">general-10</a> 参照）</p>
<h3 id="advanced" class="text-xl font-bold mt-8 mb-3">advanced</h3>
<p class="mb-4 text-slate-700 leading-7">ページ設定と印刷 / 保護 / ウィンドウ・枠線 / Shapes</p>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "CollectionPage",
   "name": "サンプル集 | osbxl",
   "description": "実行検証済みの osbxl サンプル集。Node.js / Python / Java / PHP / WebAssembly の5言語タブ切り替え対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-06",
   "url": "https://osb-office.com/sample/",
   "publisher": {
    "@type": "Organization",
    "name": "osb-office.inc",
    "url": "https://osb-office.com/",
    "logo": "https://osb-office.com/og-image.png"
   }
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
