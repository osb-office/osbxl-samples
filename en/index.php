<?php
$lang = 'en';
$page_title = 'Sample Collection | osbxl';
$no_tagline = true;
$meta_description = 'Execution-verified osbxl samples. Switchable code tabs for Node.js / Python / Java / PHP / WebAssembly.';
$meta_keywords = 'osbxl,sample,samples,Excel,Excel manipulation library,Excel generation library,Excel I/O library,Node.js,Python,Java,PHP,WebAssembly';
include __DIR__ . '/../../includes/header.php';
include __DIR__ . '/../../includes/sidebar.php';
?>

<main class="main-content">
<p class="text-sm mb-4 text-right"><a href="/sample/" class="text-primary-600 hover:underline">日本語</a></p><h1 id="osbxl-sample-collection" class="text-3xl font-bold mb-2 mt-2">osbxl Sample Collection</h1>
<p class="mb-4 text-slate-700 leading-7">These osbxl samples have all been executed and verified. The code on each page runs as-is. Every code block switches between <strong>five language tabs — Node.js / Python / Java / PHP / WebAssembly</strong> — and the install instructions shown per tab are language-specific too.</p>
<p class="mb-4 text-slate-700 leading-7"><strong>If you are new here, start with <a href="general-01-overview.php" class="text-primary-600 hover:underline">general-01 Overview — object model and basic flow</a>.</strong></p>
<h2 id="find-a-sample-by-what-you-want-to-do" class="text-2xl font-bold mt-10 mb-4">Find a sample by what you want to do</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">What you want to do</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Sample</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Understand the big picture of osbxl and the class hierarchy</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php" class="text-primary-600 hover:underline">general-01 Overview</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">The flow of creating a new book / opening an existing book / save / saveAs / close</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php#basic-flow" class="text-primary-600 hover:underline">general-01 Overview — basic flow</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">What methods App / WorkBook / WorkSheet / Range have</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php#main-methods" class="text-primary-600 hover:underline">general-01 Overview — main methods</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Build a table with values and formulas</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">general-02 Values and formulas</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Enter per-row shifting formulas (relative references) in bulk</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-02-values-formulas.php#code" class="text-primary-600 hover:underline">general-02 Values and formulas</a> / <a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">general-06 Bulk input</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Add, delete, move or rename sheets</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-03-sheets.php" class="text-primary-600 hover:underline">general-03 Sheet operations</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">List all sheets of a book whose sheet names you don't know</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-03-sheets.php#code" class="text-primary-600 hover:underline">general-03 Sheet operations</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Enter dates, times and booleans / set number formats</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-04-value-types.php" class="text-primary-600 hover:underline">general-04 Cell value types</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">The difference between <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> and <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-04-value-types.php#getvalue-vs-getvalue-true" class="text-primary-600 hover:underline">general-04 Cell value types</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Insert, delete, clear or copy rows and columns</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">general-05 Rows, columns and cell ranges</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Set row height, column width and hidden state</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-05-rows-cols.php#code-2-row-col-objects" class="text-primary-600 hover:underline">general-05 Rows, columns and cell ranges</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Populate tens of thousands of cells fast</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">general-06 Bulk input</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Filter or sort data (AutoFilter / Sort)</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">general-07 AutoFilter and sorting</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Use tables (totals row, structured references) and comments</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">general-08 Tables and comments</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Set conditional formatting (color scales, data bars, top N)</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">general-09 Conditional formatting</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Create charts (column, line)</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-10-charts.php" class="text-primary-600 hover:underline">general-10 Charts</a></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Aggregate with pivot tables</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-11-pivottables.php" class="text-primary-600 hover:underline">general-11 Pivot tables</a></td></tr>
  </tbody></table></div>
<h2 id="general-overview-and-basic-operations" class="text-2xl font-bold mt-10 mb-4">general — overview and basic operations</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">#</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Title</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">level</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">Main APIs</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">01</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-01-overview.php" class="text-primary-600 hover:underline">Overview — object model and basic flow</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">App / WorkBook / WorkSheet / Range</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">02</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">Building a Table with Values and Formulas</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">Range / InputFormulaObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">03</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-03-sheets.php" class="text-primary-600 hover:underline">Working with books and sheets — add, delete, move, rename</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook / WorkSheet</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">04</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-04-value-types.php" class="text-primary-600 hover:underline">Cell value types and number formats — strings, numbers, dates, times, booleans</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">Range / DateTimeObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">05</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">Rows, columns and cell ranges — insert, delete, clear, copy</a></td><td class="border border-surface-200 px-3 py-2 align-top">basic</td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet / Row / Col</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">06</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">Bulk input of large data — setValueArray / setFormulaArray</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet / InputValueObject / InputFormulaObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">07</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">AutoFilter and sorting — conditional filters, sort</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">AutoFilter / Sort</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">08</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">Tables and comments — ListObjects, totals row, structured references</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">ListObjects / Table / Comments</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">09</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">Conditional formatting — cell value rules, color scales, data bars</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">FormatConditions / ColorScaleObject / DataBarObject</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">10</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-10-charts.php" class="text-primary-600 hover:underline">Charts basics — column, line, title, axes, legend</a></td><td class="border border-surface-200 px-3 py-2 align-top">intermediate</td><td class="border border-surface-200 px-3 py-2 align-top">Chart / chart.* / enums.XlChartType</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">11</td><td class="border border-surface-200 px-3 py-2 align-top"><a href="general-11-pivottables.php" class="text-primary-600 hover:underline">Pivot tables — create, aggregate, change aggregation method</a></td><td class="border border-surface-200 px-3 py-2 align-top">advanced</td><td class="border border-surface-200 px-3 py-2 align-top">PivotTable / PivotField</td></tr>
  </tbody></table></div>
<h2 id="planned-not-yet-written" class="text-2xl font-bold mt-10 mb-4">Planned (not yet written)</h2>
<h3 id="general" class="text-xl font-bold mt-8 mb-3">general</h3>
<p class="mb-4 text-slate-700 leading-7">Book properties (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCompanyName</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCreateAuthor</code> etc.) / named ranges (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNames</code>) / custom styles (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomStyle</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomTableStyle</code>) / external workbook integration (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addExternalWorkBook</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">updateExternalWorkBook</code>) / password-protected books (<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openPasswordWorkBook</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save(password)</code>) / outlining in detail</p>
<h3 id="format" class="text-xl font-bold mt-8 mb-3">format</h3>
<p class="mb-4 text-slate-700 leading-7">Fonts and colors / borders / fills and gradients / alignment and wrapping / number formats (numeric, date, currency) / built-in and custom styles (for conditional formatting see <a href="general-09-conditional-format.php" class="text-primary-600 hover:underline">general-09</a>)</p>
<h3 id="data" class="text-xl font-bold mt-8 mb-3">data</h3>
<p class="mb-4 text-slate-700 leading-7">Cell merging / hyperlinks (for tables, sorting, AutoFilter and comments see <a href="general-07-autofilter-sort.php" class="text-primary-600 hover:underline">general-07</a> / <a href="general-08-tables-comments.php" class="text-primary-600 hover:underline">general-08</a>)</p>
<h3 id="formula" class="text-xl font-bold mt-8 mb-3">formula</h3>
<p class="mb-4 text-slate-700 leading-7">Formula basics and array formulas / financial functions (ACCINT etc.) / lookup functions (XLOOKUP) / forecasting (FORECAST.ETS) / WEBSERVICE + FILTERXML</p>
<h3 id="pivot" class="text-xl font-bold mt-8 mb-3">pivot</h3>
<p class="mb-4 text-slate-700 leading-7">Aggregation and grouping (for the basics of creating and aggregating see <a href="general-11-pivottables.php" class="text-primary-600 hover:underline">general-11</a>)</p>
<h3 id="chart" class="text-xl font-bold mt-8 mb-3">chart</h3>
<p class="mb-4 text-slate-700 leading-7">Chart varieties / data labels / formatting (colors, lines, shadows) (for basic creation, axes and legend see <a href="general-10-charts.php" class="text-primary-600 hover:underline">general-10</a>)</p>
<h3 id="advanced" class="text-xl font-bold mt-8 mb-3">advanced</h3>
<p class="mb-4 text-slate-700 leading-7">Page setup and printing / protection / windows and gridlines / Shapes</p>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "CollectionPage",
   "name": "Sample Collection | osbxl",
   "description": "Execution-verified osbxl samples. Switchable code tabs for Node.js / Python / Java / PHP / WebAssembly.",
   "inLanguage": "en",
   "dateModified": "2026-10-06",
   "url": "https://osb-office.com/en/sample/",
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
     "name": "Home",
     "item": "https://osb-office.com/en/"
    },
    {
     "@type": "ListItem",
     "position": 2,
     "name": "Sample Collection",
     "item": "https://osb-office.com/en/sample/"
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
