<?php
$page_title = '概要 — オブジェクトモデルと基本の流れ | osbxl サンプル';
$no_tagline = true;
$meta_description = '概要 — オブジェクトモデルと基本の流れ。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,overview,object-model,lifecycle,app,workbook,worksheet,range,create,open,save,saveas,close';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">概要 — オブジェクトモデルと基本の流れ</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-01-overview.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">osbxl の全体像をつかむ。</p>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li>オブジェクトの階層（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code>）</li>
  <li><strong>基本の流れ</strong>（新規作成 / 既存ブックを開く → シート → Range → 値・数式 → save / saveAs / close）</li>
  <li>4 クラスの主要メソッド</li>
</ul>
<p class="mb-4 text-slate-700 leading-7">個々の機能の詳細は、このページの後続サンプルで扱います。</p>
<h2 id="前提" class="text-2xl font-bold mt-10 mb-4">前提</h2>
<p class="mb-4 text-slate-700 leading-7">各言語のモジュールを<a href="/pages/download.php" class="text-primary-600 hover:underline">ダウンロードページ</a>から入手してインストールします。</p>
<p class="mb-4 text-slate-700 leading-7">PHP の OS 別の詳しい導入手順（Linux / macOS / Windows、Composer が未インストールの場合を含む）は<a href="/pages/quickstart.php" class="text-primary-600 hover:underline">クイックスタート</a>を参照してください。</p>
<p class="mb-4 text-slate-700 leading-7">WebAssembly 版は MEMORY64（Wasm 3.0）ビルドのため、実行環境側の対応が必要です （Chrome/Edge 133〜・Firefox 134〜・Node.js 24〜。Safari は現在未対応）。</p>
<p class="mb-4 text-slate-700 leading-7">入口は言語ごとに異なります。namespace 構成（操作対象クラス / dto / chart / enums）は4言語共通です。</p>
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
        <pre><code class="language-bash">npm install ./nodeosbxl-1.5.0.tgz</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># (例) Python 3.12 以上（abi3 wheel 1 本で 3.12 以降をカバー）
pip install ./pyosbxl-1.5.0-cp312-abi3-linux_x86_64.whl</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># JAR をクラスパスに追加（javacpp-1.5.14.jar も必要）
javac -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar Main.java
java  -cp osbxl-1.5.0.jar:javacpp-1.5.14.jar:. Main</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># HP (osb-office.com/files/php/) から PHP バージョン(8.2〜8.5)と OS に合う
# tarball または composer zip をダウンロードする
# 例: phposbxl-1.5.0-php8.5-nts-linux-x86_64.tar.gz (Linux・PHP 8.5 の場合)
# 展開して composer の path リポジトリに指定し導入する
composer config repositories.phposbxl path /path/to/phposbxl-1.5.0-php8.5-nts-linux-x86_64
composer require osb-office/phposbxl
# composer zip の場合: zip を置いたディレクトリを artifact リポジトリに指定
# composer config repositories.phposbxl artifact /path/to/zip-dir &amp;&amp; composer require osb-office/phposbxl</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-bash"># パッケージを展開（osbxl_wasm.js / osbxl_em.js / license/ を含む）
tar xzf wasmosbxl-1.5.0.tgz</code></pre>
        <pre><code class="language-javascript">// 読み込み例（Node.js 24〜 必須: この wasm は MEMORY64 = Wasm 3.0 ビルドです）
const fs = require("fs");
const createModule = require("./package/osbxl_wasm.js");   // emscripten のファクトリ
const { createOsbxl } = require("./package/osbxl_em.js");  // ローダー＋enum 定義

createOsbxl(createModule).then(({ Module, osbxl, dto, chart, enums }) =&gt; {
  // ライセンスファイルを wasm 仮想 FS へ配置（new App() より前に必須）
  const lic = process.env.OSB_LICENSE_PATH_WASM;          // 例: /root/lic_node.osb
  if (lic &amp;&amp; fs.existsSync(lic)) {
    Module.FS.mkdirTree(lic.substring(0, lic.lastIndexOf("/")));
    Module.FS.writeFile(lic, fs.readFileSync(lic));
  }
  const app = new osbxl.App();   // ここから先は Node.js 版と同じ（namespace 名のみ osbxl）
  // ...
});
// ブラウザでは &lt;script&gt; で osbxl_wasm.js を読み込んで同様に createOsbxl を呼びます。
// 対応環境: Chrome/Edge 133〜・Firefox 134〜（Safari は現在未対応）</code></pre>
      </div>
    </div>
</div>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">言語</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">入口</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">主な namespace</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Node.js</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">require("nodeosbxl")</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Python</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">import pyosbxl</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl.dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl.chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">pyosbxl.enums</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">Java</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">import com.osboffice.osbxl.*</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code> 等のパッケージ</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">WebAssembly</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createOsbxl(createModule)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">osbxl</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">下表は Node.js 表記です。各 namespace の中身の数は4言語共通です。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">namespace</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">中身</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">数</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">nodeosbxl</code></td><td class="border border-surface-200 px-3 py-2 align-top">操作の対象になるクラス（App / WorkBook / WorkSheet / Range / Font / Chart 系の入口 など）</td><td class="border border-surface-200 px-3 py-2 align-top">29 クラス</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto</code></td><td class="border border-surface-200 px-3 py-2 align-top">データの受け渡しに使うオブジェクト（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InputValueObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">FontObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ColorObject</code> など）</td><td class="border border-surface-200 px-3 py-2 align-top">40 クラス</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">chart</code></td><td class="border border-surface-200 px-3 py-2 align-top">グラフ専用オブジェクト（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartObject</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Axis</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ChartTitle</code> など）</td><td class="border border-surface-200 px-3 py-2 align-top">43 クラス</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums</code></td><td class="border border-surface-200 px-3 py-2 align-top">Excel 互換の列挙型（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlFont</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlChartType</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">XlIndexColor</code> など）</td><td class="border border-surface-200 px-3 py-2 align-top">127 種</td></tr>
  </tbody></table></div>
<h2 id="オブジェクトモデル" class="text-2xl font-bold mt-10 mb-4">オブジェクトモデル</h2>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">App                     ← すべての入口。ブックを開く/作る、セルアドレスの変換
 │
 ├─ createWorkBook()  ─┐
 ├─ openWorkBook()    ─┴→ WorkBook          ← xlsx ファイル 1 つ
 │                        │
 │                        ├─ openWorkSheet()      → WorkSheet   ← シート 1 つ
 │                        │                         │
 │                        │                         ├─ getRange("A1:B2")    → Range
 │                        │                         │    ※ getCells(1,1,2,2) でも同じ Range が取れる
 │                        │                         │                       │
 │                        │                         │                       ├─ getValue() / setValue() / setFormula()
 │                        │                         │                       ├─ getFont()   → Font   → getColor() → Color
 │                        │                         │                       ├─ getBorders()→ Borders→ getBorder()  → Border
 │                        │                         │                       ├─ getFill()   → Fill
 │                        │                         │                       └─ getAlignment() → Alignment
 │                        │                         │
 │                        │                         ├─ getRow() / getCol()  → Row / Col
 │                        │                         ├─ getSort() / getAutoFilter()
 │                        │                         ├─ getPivotTables() / getListObjects()
 │                        │                         ├─ addChartObject()     → chart.ChartObject
 │                        │                         └─ getComments() / getHyperLinks() / getShapes() ...
 │                        │
 │                        ├─ addWorkSheet() / deleteWorkSheet() / moveSheet()
 │                        ├─ addCustomStyle() / addCustomTableStyle()
 │                        └─ save() / saveAs() / close()
 │
 └─ convertFromRowColNumber() などのユーティリティ</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><strong>この階層を常に上から辿る</strong>のが osbxl の基本です。 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> より下（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Font</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Color</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Borders</code> …）は「取得して、そのオブジェクトの setter を呼ぶ」形になります。</p>
<h2 id="基本の流れ" class="text-2xl font-bold mt-10 mb-4">基本の流れ</h2>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext"> [1] new App()                 App を生成する
 [2] createWorkBook(path)      新規ブックを作る
     openWorkBook(path)        既存のブックを開く
 [3] openWorkSheet(name)       ワークシートを開く
 [4] getRange("A1:B2")         Range を取得する（getCells(1,1,2,2) でも同じ）
 [5] setValue() / setFormula() 値・数式を入れる
 [6] getValue(true)            結果を読み戻す
 [7] save() / saveAs(path)     保存する
 [8] close()                   ブックを閉じる</code></pre>
</div>
<h3 id="コード" class="text-xl font-bold mt-8 mb-3">コード</h3>
<p class="mb-4 text-slate-700 leading-7">実行結果:</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="python">Python</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="java">Java</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="php">PHP</button>
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 border-transparent text-text-light hover:text-primary-600" data-tab="wasm">WebAssembly</button>
  </div>
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// basic-flow.js — 新規作成 → 編集 → 保存 → 開き直し → 別名保存 の一連の流れ
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();                  // [1] App を生成

// ---------- 新規ブックを作って編集し、保存する ----------
let wb = app.createWorkBook("flow.xlsx");         // [2] 新規作成
console.log("[2] createWorkBook -&gt; getBookPath =", wb.getBookPath());

let ws = wb.openWorkSheet("Sheet1");              // [3] ワークシートを開く
console.log("[3] openWorkSheet  -&gt; シート名 =", ws.getName());

// [4] Range を取得して [5] 値を設定
//     数値(行, 列)で取る ws.getCells(1, 1) も getRange("A1:A1") と同じ Range を返します
ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100*3");         // [5] 数式を設定（先頭の "=" は付けない）
console.log("[6] getValue(true) -&gt; B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);

wb.save();                                        // [7] 保存
console.log("[7] save() 完了");
wb.close();                                       // [8] 閉じる（WorkSheet の close は不要）
console.log("[8] close() 完了");

// ---------- 既存ブックを開いて、別名で保存する ----------
wb = app.openWorkBook("flow.xlsx");               // [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1");                  // [3']
const read = ws.getRange("A1:B1").getValue(true);
console.log("[2'] openWorkBook -&gt; A1 =", read["A1"], "/ B1 =", read["B1"]);

ws.getRange("A2:A2").setValue("開き直して追記");
wb.saveAs("flow-copy.xlsx");                      // [7'] 別名で保存
console.log("[7'] saveAs() 完了 -&gt; getBookPath =", wb.getBookPath());
wb.close();                                       // [8']
console.log("[8] close() 完了");</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># basic-flow.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()                               # [1] App を生成

wb = app.createWorkBook("flow-py.xlsx")           # [2] 新規作成
ws = wb.openWorkSheet("Sheet1")                   # [3] ワークシートを開く

ws.getRange("A1:A1").setValue("新規作成")
ws.getRange("B1:B1").setFormula("100*3")          # [5] 数式（先頭の "=" は付けない）
print("[6] getValue(True) -&gt; B1 =", ws.getRange("B1:B1").getValue(True)["B1"])

wb.save()                                         # [7] 保存
wb.close()                                        # [8] 閉じる
print("[8] close() 完了")

wb = app.openWorkBook("flow-py.xlsx")             # [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1")
print("[2'] openWorkBook -&gt; A1 =", ws.getRange("A1:A1").getValue(True)["A1"])
wb.close()
print("[8] close() 完了")</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// basic-flow.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();                          // [1] App を生成

        WorkBookWrapper wb = app.createWorkBook("flow-java.xlsx");  // [2] 新規作成
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");           // [3] ワークシートを開く

        ws.getRange("A1:A1").setValue("新規作成");
        ws.getRange("B1:B1").setFormula("100*3");                   // [5] 数式（先頭の "=" は付けない）
        Map&lt;String, String&gt; v = ws.getRange("B1:B1").getValue(true);
        System.out.println("[6] getValue(true) -&gt; B1 = " + v.get("B1"));

        wb.save();                                                  // [7] 保存
        wb.close();                                                 // [8] 閉じる
        System.out.println("[8] close() 完了");

        wb = app.openWorkBook("flow-java.xlsx");                    // [2'] 既存のブックを開く
        ws = wb.openWorkSheet("Sheet1");
        Map&lt;String, String&gt; r = ws.getRange("A1:A1").getValue(true);
        System.out.println("[2'] openWorkBook -&gt; A1 = " + r.get("A1"));
        wb.close();
        System.out.println("[8] close() 完了");
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// basic-flow.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)
$app = new \phposbxl\App();                       // [1] App を生成

// ---------- 新規ブックを作って編集し、保存する ----------
$wb = $app-&gt;createWorkBook("flow-php.xlsx");      // [2] 新規作成
echo "[2] createWorkBook -&gt; getBookPath = " . $wb-&gt;getBookPath() . "\n";

$ws = $wb-&gt;openWorkSheet("Sheet1");               // [3] ワークシートを開く
echo "[3] openWorkSheet  -&gt; シート名 = " . $ws-&gt;getName() . "\n";

// [4] Range を取得して [5] 値を設定
//     数値(行, 列)で取る $ws-&gt;getCells(1, 1) も getRange("A1:A1") と同じ Range を返します
$ws-&gt;getRange("A1:A1")-&gt;setValue("新規作成");
$ws-&gt;getRange("B1:B1")-&gt;setFormula("100*3");      // [5] 数式を設定（先頭の "=" は付けない）
$v = $ws-&gt;getRange("B1:B1")-&gt;getValue(true);
echo "[6] getValue(true) -&gt; B1 = " . $v["B1"] . "\n";

$wb-&gt;save();                                      // [7] 保存
echo "[7] save() 完了\n";
$wb-&gt;close();                                     // [8] 閉じる（WorkSheet の close は不要）
echo "[8] close() 完了\n";

// ---------- 既存ブックを開いて、別名で保存する ----------
$wb = $app-&gt;openWorkBook("flow-php.xlsx");        // [2'] 既存のブックを開く
$ws = $wb-&gt;openWorkSheet("Sheet1");               // [3']
$read = $ws-&gt;getRange("A1:B1")-&gt;getValue(true);
echo "[2'] openWorkBook -&gt; A1 = " . $read["A1"] . " / B1 = " . $read["B1"] . "\n";

$ws-&gt;getRange("A2:A2")-&gt;setValue("開き直して追記");
$wb-&gt;saveAs("flow-copy-php.xlsx");                // [7'] 別名で保存
echo "[7'] saveAs() 完了 -&gt; getBookPath = " . $wb-&gt;getBookPath() . "\n";
$wb-&gt;close();                                     // [8']
echo "[8] close() 完了\n";</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// basic-flow.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();                        // [1] App を生成

const wb = app.createWorkBook("flow-wasm.xlsx");    // [2] 新規作成
const ws = wb.openWorkSheet("Sheet1");              // [3] ワークシートを開く

ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100*3");           // [5] 数式（先頭の "=" は付けない）
console.log("[6] getValue(true) -&gt; B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);

wb.save();                                          // [7] 保存
wb.close();                                         // [8] 閉じる
console.log("[8] close() 完了");

const wb2 = app.openWorkBook("flow-wasm.xlsx");     // [2'] 既存のブックを開く
const ws2 = wb2.openWorkSheet("Sheet1");
console.log("[2'] openWorkBook -&gt; A1 =", ws2.getRange("A1:A1").getValue(true)["A1"]);
wb2.close();
console.log("[8] close() 完了");</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[2] createWorkBook -&gt; getBookPath = &lt;実行ディレクトリ&gt;/flow.xlsx
[3] openWorkSheet  -&gt; シート名 = Sheet1
[6] getValue(true) -&gt; B1 = 300
[7] save() 完了
[8] close() 完了
[2'] openWorkBook -&gt; A1 = 新規作成 / B1 = 300
[7'] saveAs() 完了 -&gt; getBookPath = &lt;実行ディレクトリ&gt;/flow.xlsx
[8] close() 完了</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">flow.xlsx</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">flow-copy.xlsx</code> の 2 つができます。</p>
<h3 id="流れのポイント" class="text-xl font-bold mt-8 mb-3">流れのポイント</h3>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> は 1 プロセスに 1 つでよい。</strong> ブックを複数扱う場合も同じ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app</code> から <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkBook()</code> を呼びます。</li>
  <li><strong>相対パスは絶対パスに解決される。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook("flow.xlsx")</code> の後、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBookPath()</code> はカレントディレクトリを含む絶対パスを返します。</li>
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> を閉じる処理は無い。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">close()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> のメソッドだけです。 シートは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.close()</code> でまとめて解放されます。</li>
  <li><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs()</code> の違いに注意</strong>（下記の「注意点」参照）。</li>
</ul>
<h2 id="主要メソッド" class="text-2xl font-bold mt-10 mb-4">主要メソッド</h2>
<h3 id="app-10-メソッド" class="text-xl font-bold mt-8 mb-3">App（10 メソッド）</h3>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new nodeosbxl.App()</code> で生成します。ブックの開閉と、アドレス変換ユーティリティを持ちます。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getVersion()</code></td><td class="border border-surface-200 px-3 py-2 align-top">バージョン番号を返します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook(path, defaultFont?, defaultFontSize?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブックを作成します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkBook(path)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブックを開きます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openPasswordWorkBook(path, password)</code></td><td class="border border-surface-200 px-3 py-2 align-top">パスワード付きワークブックを開きます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber(row, col)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セルの数値表記を A1C1 表記に変換します。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(1,1)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1"</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber2(startRow, startCol, endRow, endCol)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル範囲の数値表記を A1C1 表記に変換します。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(1,1,2,3)</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:C2"</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertToColumnNumber("B")</code></td><td class="border border-surface-200 px-3 py-2 align-top">列の A1 表記を数値表記に変換します。→ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromColumnNumber(3)</code></td><td class="border border-surface-200 px-3 py-2 align-top">列の数値表記を A1 表記に変換します。→ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"C"</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumericValue(dateTimeObject, is1904?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">日付時刻オブジェクトから Excel 内部のシリアル値を取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">calculateHexColor(bookPath, colorObj, isForeGroundColor)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブック固有の色情報を加味した HEX 値を取得します。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">convertFromRowColNumber2</code> には絶対参照（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"$A$1"</code>）を返すオーバーロードもあります。</p>
<h3 id="workbook-27-メソッド-主要-12" class="text-xl font-bold mt-8 mb-3">WorkBook（27 メソッド）— 主要 12</h3>
<p class="mb-4 text-slate-700 leading-7">xlsx ファイル 1 つに対応します。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBookPath()</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブックのファイルパスを返します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを開きます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex(sheetIndex)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートをインデックスで開きます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet(sheetName, position?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを追加します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを削除します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet(sheetName, position)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを移動します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシート数を取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetNumber(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートのシート番号を取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save(password)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブックを同一ファイルに保存します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs(path)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs(path, password)</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブックを別ファイルに保存します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">close()</code></td><td class="border border-surface-200 px-3 py-2 align-top">ワークブックを閉じます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isDate1904()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Excel の日付 1904 年形式であるかを取得します。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">その他（15）: スタイル系 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNames</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addCustomTableStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addPivotCustomTableStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCustomTableStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCustomTableStyle</code> / 外部ブック系 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addExternalWorkBook</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">updateExternalWorkBook</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getExternalWorkBookPath</code> / ファイル属性 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCompanyName</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setManagerName</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCreateAuthor</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setLastAuthor</code></p>
<h3 id="worksheet-44-メソッド-主要-14" class="text-xl font-bold mt-8 mb-3">WorkSheet（44 メソッド）— 主要 14</h3>
<p class="mb-4 text-slate-700 leading-7">シート 1 つに対応します。<strong>ここから下のオブジェクト（Range / Sort / PivotTable / Chart …）へ辿る入口</strong>です。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange("A1:B2")</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル範囲クラスインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells(row, col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCells(sr, sc, er, ec)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル範囲クラスインスタンスを数値指定で取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRow(rowNum)</code></td><td class="border border-surface-200 px-3 py-2 align-top">行クラスインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCol(colNum)</code></td><td class="border border-surface-200 px-3 py-2 align-top">列クラスインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setName(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">シート名称の取得 / 設定を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isActive()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">シートタブが起動時の表示項目かどうかの取得 / 設定を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertRow(startRow, numOfRows)</code></td><td class="border border-surface-200 px-3 py-2 align-top">行の挿入を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteRow(startRow, numOfRows)</code></td><td class="border border-surface-200 px-3 py-2 align-top">行の削除を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCol(startCol, numOfCols)</code></td><td class="border border-surface-200 px-3 py-2 align-top">列の挿入を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCol(startCol, numOfCols)</code></td><td class="border border-surface-200 px-3 py-2 align-top">列の削除を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValueArray(values)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル値の一括入力を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray(formulas)</code></td><td class="border border-surface-200 px-3 py-2 align-top">関数の一括入力を行います。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">その他（30）: 機能オブジェクト取得 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormatConditions</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAutoFilter</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSort</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getComments</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHyperLinks</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getListObjects</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPivotTables</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getShapes</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWindow</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getHPageBreaks</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getVPageBreaks</code> / グラフ <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getChartObject</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addChartObject</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteChartObject</code> / 印刷設定 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPageSetupObject</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setPageSetupObject</code> / 選択 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getActiveCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setActiveCell</code> / 行・列・セルの操作 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCol</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCol</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCell</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell</code> / アウトライン <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setRowOutline</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRowOutline</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColOutline</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearColOutline</code></p>
<h3 id="range-31-メソッド-主要-16" class="text-xl font-bold mt-8 mb-3">Range（31 メソッド）— 主要 16</h3>
<p class="mb-4 text-slate-700 leading-7">セル範囲に対応します。<strong>値・数式の読み書きと、書式オブジェクトへの入口</strong>です。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(rawValue?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セルの値の取得を行います。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> で生値（数式の結果）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code></td><td class="border border-surface-200 px-3 py-2 align-top">数式の取得を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(value, forceString?, numberFormat?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル値の設定を行います（汎用メソッド）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(value, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル値の設定を行います（数値設定メソッド）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(dateTimeObject, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル値の設定を行います（日付/時刻設定メソッド）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue(str, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル値の設定を行います（日付/時刻を文字列から入力）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(bool, ...)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル値の設定を行います（Boolean 値設定メソッド）。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula(formula, isArray?, setAllCell?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">数式の設定を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFont()</code></td><td class="border border-surface-200 px-3 py-2 align-top">フォントオブジェクトインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getBorders()</code></td><td class="border border-surface-200 px-3 py-2 align-top">枠線一括オブジェクトインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFill()</code></td><td class="border border-surface-200 px-3 py-2 align-top">フィルオブジェクトインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAlignment()</code></td><td class="border border-surface-200 px-3 py-2 align-top">セルの配置位置指定オブジェクトインスタンスを取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberFormat(format)</code></td><td class="border border-surface-200 px-3 py-2 align-top">表示書式の設定を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">merge(mergeEachRow?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">セル範囲をマージします。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearContent()</code></td><td class="border border-surface-200 px-3 py-2 align-top">セルのクリアを行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getAddress()</code></td><td class="border border-surface-200 px-3 py-2 align-top">Range の選択範囲を取得します。</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">その他（15）: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getProtection</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getCharacters</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getPhonetics</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getNumberFormat</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearNumberFormat</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearAllFormat</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDataTable</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearDataTable</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBuiltinStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setCustomStyle</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalWidth</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalColumnWidth</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getTotalRowHeight</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">unMerge</code> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">replaceValue</code></p>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> は「取得して使い捨て」でよい。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("A1:A1")</code> は呼ぶたびに新しいインスタンスを返すので、 変数に保持し続ける必要はありません。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> の戻り値はマップ。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ "A1": "100", "B1": "300" }</code> の形で、値はすべて<strong>文字列</strong>です。<br> 数値として扱うには <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Number(...)</code> で変換します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>Range の取り方は 2 通り。</strong><ul class="mt-2 space-y-1 pl-5 list-disc text-slate-700"><li><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("A1:B2")</code> — <strong>文字列アドレス</strong>で指定。単一セルでも <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:A1"</code> と範囲の形で渡します。</li><li><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getCells(row, col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getCells(startRow, startCol, endRow, endCol)</code> — <strong>数値（行, 列）</strong>で指定。 行・列とも <strong>1 始まり</strong>で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(1,1)</code> = A1、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(3,1,3,3)</code> = A3:C3 です。 ループで行列番号を回すときはアドレス文字列を組み立てずに済むので便利です。</li><li>数値 ⇔ 文字列の変換は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber(row, col)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber2(sr, sc, er, ec)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertToColumnNumber("B")</code> を使います。</li></ul></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>数式は実行時に計算される。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula()</code> の直後に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> で計算結果を読めます （<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> を待つ必要はありません）。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>ワークブックは、一意性を確保するため内部でロックされ、この処理によって同一ワークブックの同時読込・書込は起こりません。</strong></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">close()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> のメソッドだけ。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> に close 系のメソッドはありません。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs()</code> は「その時点の内容を別ファイルに書き出す」だけで、以降の保存先は切り替わりません。</strong> Excel VBA の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SaveAs</code> とは挙動が違います。<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-plaintext">wb.save()                  → base.xlsx に書かれる
wb.saveAs("copy.xlsx")     → copy.xlsx に「この時点の内容」が書かれる
wb.getBookPath()           → base.xlsx のまま（切り替わらない）
ws.getRange(...).setValue(...)
wb.save()                  → base.xlsx に書かれる（copy.xlsx ではない）</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>数式に先頭の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> を付けない。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code> です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>起動時にライセンス表示が stdout に 1 行出力されます</strong>（プロセスにつき 1 回）。<br> 標準出力をパースする処理を書く場合は注意してください。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>次のサンプル: <a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">値と数式を入れて表を作る</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/python/pyosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/python/pyosbxl.Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/AppWrapper.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/java/com/osboffice/osbxl/WorkBookWrapper.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/java/com/osboffice/osbxl/RangeWrapper.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/php/classes/phposbxl-WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/php/classes/phposbxl-Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.App.html" class="text-primary-600 hover:underline">App</a> / <a href="/api-docs/wasm/classes/osbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/wasm/classes/osbxl.Range.html" class="text-primary-600 hover:underline">Range</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "概要 — オブジェクトモデルと基本の流れ",
   "description": "概要 — オブジェクトモデルと基本の流れ。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-01-overview.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-01-overview.php",
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
     "name": "概要 — オブジェクトモデルと基本の流れ",
     "item": "https://osb-office.com/sample/general-01-overview.php"
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
