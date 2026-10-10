<?php
$page_title = '値と数式を入れて表を作る | osbxl サンプル';
$no_tagline = true;
$meta_description = '値と数式を入れて表を作る。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,value,formula,sum,table,setformulaarray,inputformulaobject,convertfromrowcolnumber';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">値と数式を入れて表を作る</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-02-values-formulas.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">新しい xlsx を作って、表（見出し・文字列・数値・数式）を書き込み、ファイルに保存する。 <a href="general-01-overview.php" class="text-primary-600 hover:underline">概要</a>の基本の流れを実際の表作成に当てはめたものです。</p>
<h2 id="完成イメージ" class="text-2xl font-bold mt-10 mb-4">完成イメージ</h2>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">sales.xlsx</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet1</code> に次の表ができます。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap"></th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">A</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">B</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">C</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">D</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>1</strong></td><td class="border border-surface-200 px-3 py-2 align-top">商品</td><td class="border border-surface-200 px-3 py-2 align-top">単価</td><td class="border border-surface-200 px-3 py-2 align-top">個数</td><td class="border border-surface-200 px-3 py-2 align-top">金額</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>2</strong></td><td class="border border-surface-200 px-3 py-2 align-top">りんご</td><td class="border border-surface-200 px-3 py-2 align-top">100</td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">300</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>3</strong></td><td class="border border-surface-200 px-3 py-2 align-top">みかん</td><td class="border border-surface-200 px-3 py-2 align-top">50</td><td class="border border-surface-200 px-3 py-2 align-top">6</td><td class="border border-surface-200 px-3 py-2 align-top">300</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>4</strong></td><td class="border border-surface-200 px-3 py-2 align-top">ぶどう</td><td class="border border-surface-200 px-3 py-2 align-top">150</td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">300</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>5</strong></td><td class="border border-surface-200 px-3 py-2 align-top"></td><td class="border border-surface-200 px-3 py-2 align-top"></td><td class="border border-surface-200 px-3 py-2 align-top">合計</td><td class="border border-surface-200 px-3 py-2 align-top">900</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">D 列は数式 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">B*C</code>、D5 は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">SUM(D2:D4)</code> です。</p>
<h2 id="前提" class="text-2xl font-bold mt-10 mb-4">前提</h2>
<p class="mb-4 text-slate-700 leading-7">各言語のモジュールを<a href="/pages/download.php" class="text-primary-600 hover:underline">ダウンロードページ</a>から入手してインストールします （入口の表記は <a href="general-01-overview.php#前提" class="text-primary-600 hover:underline">概要の前提</a> 表を参照）。</p>
<p class="mb-4 text-slate-700 leading-7">PHP の OS 別の詳しい導入手順（Linux / macOS / Windows、Composer が未インストールの場合を含む）は<a href="/pages/quickstart.php" class="text-primary-600 hover:underline">クイックスタート</a>を参照してください。</p>
<p class="mb-4 text-slate-700 leading-7">WebAssembly 版は MEMORY64（Wasm 3.0）ビルドのため、実行環境側の対応が必要です （Chrome/Edge 133〜・Firefox 134〜・Node.js 24〜。Safari は現在未対応）。</p>
<p class="mb-4 text-slate-700 leading-7">保存先パスを省略形（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"sales.xlsx"</code>）にすると、<strong>プロセスのカレントディレクトリ</strong>に書き出されます。</p>
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
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// values-formulas.js — ブックを作って、値と数式を入れて保存する
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sales.xlsx");   // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1");         // 既定のシートを開く

// --- 見出し (1行目) ---
// app.convertFromRowColNumber(行, 列) で (1,1) -&gt; "A1" のセルアドレスが得られる
["商品", "単価", "個数", "金額"].forEach((text, i) =&gt; {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- 明細 (2〜4行目) ---
const data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
data.forEach(([name, price, qty], i) =&gt; {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((_, i) =&gt; {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(i + 2, 4));  // D2, D3, D4
    f.setFormula(`B${i + 2}*C${i + 2}`);               // "B2*C2" ...
    return f;
});
ws.setFormulaArray(formulas);

// --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
const values = ws.getRange("A1:D5").getValue(true);
console.log("金額 D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("合計:", values["D5"]);

// --- 保存して閉じる ---
wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># values-formulas.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sales-py.xlsx")   # 新規ブック作成
ws = wb.openWorkSheet("Sheet1")            # 既定のシートを開く

# --- 見出し (1行目) ---
# app.convertFromRowColNumber(行, 列) で (1,1) -&gt; "A1" のセルアドレスが得られる
for i, text in enumerate(["商品", "単価", "個数", "金額"]):
    cell = app.convertFromRowColNumber(1, i + 1)
    ws.getRange(f"{cell}:{cell}").setValue(text)

# --- 明細 (2〜4行目) ---
data = [("りんご", 100, 3), ("みかん", 50, 6), ("ぶどう", 150, 2)]
for i, (name, price, qty) in enumerate(data):
    row = i + 2
    ws.getRange(f"A{row}:A{row}").setValue(name)
    ws.getRange(f"B{row}:B{row}").setNumberValue(price)
    ws.getRange(f"C{row}:C{row}").setNumberValue(qty)

# --- 金額列 (D2:D4) ---
# 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
# 数式文字列に先頭の "=" は付けない。
formulas = []
for i in range(len(data)):
    f = pyosbxl.dto.InputFormulaObject()
    f.setA1C1(app.convertFromRowColNumber(i + 2, 4))  # D2, D3, D4
    f.setFormula(f"B{i + 2}*C{i + 2}")                # "B2*C2" ...
    formulas.append(f)
ws.setFormulaArray(formulas)

# --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計")
ws.getRange("D5:D5").setFormula("SUM(D2:D4)")

# --- 読み戻し ---
# getValue(True) は「生値」= 数式の計算結果を返す
values = ws.getRange("A1:D5").getValue(True)
print("金額 D2:D4 =", ", ".join([values["D2"], values["D3"], values["D4"]]))
print("合計:", values["D5"])

# --- 保存して閉じる ---
wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// values-formulas.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("sales-java.xlsx");   // 新規ブック作成
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");             // 既定のシートを開く

        // --- 見出し (1行目) ---
        // app.convertFromRowColNumber(行, 列) で (1,1) -&gt; "A1" のセルアドレスが得られる
        String[] headers = { "商品", "単価", "個数", "金額" };
        for (int i = 0; i &lt; headers.length; i++) {
            String cell = app.convertFromRowColNumber(1, i + 1);
            ws.getRange(cell + ":" + cell).setValue(headers[i]);
        }

        // --- 明細 (2〜4行目) ---
        String[] names = { "りんご", "みかん", "ぶどう" };
        double[] prices = { 100, 50, 150 };
        double[] qtys = { 3, 6, 2 };
        for (int i = 0; i &lt; names.length; i++) {
            int row = i + 2;
            ws.getRange("A" + row + ":A" + row).setValue(names[i]);
            ws.getRange("B" + row + ":B" + row).setNumberValue(prices[i]);
            ws.getRange("C" + row + ":C" + row).setNumberValue(qtys[i]);
        }

        // --- 金額列 (D2:D4) ---
        // 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
        // 数式文字列に先頭の "=" は付けない。
        List&lt;InputFormulaObjectWrapper&gt; formulas = new ArrayList&lt;&gt;();
        for (int i = 0; i &lt; names.length; i++) {
            InputFormulaObjectWrapper f = new InputFormulaObjectWrapper();
            f.setA1C1(app.convertFromRowColNumber(i + 2, 4));   // D2, D3, D4
            f.setFormula("B" + (i + 2) + "*C" + (i + 2));       // "B2*C2" ...
            formulas.add(f);
        }
        ws.setFormulaArray(formulas);

        // --- 合計 (D5) ---
        ws.getRange("C5:C5").setValue("合計");
        ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

        // --- 読み戻し ---
        // getValue(true) は「生値」= 数式の計算結果を返す
        Map&lt;String, String&gt; values = ws.getRange("A1:D5").getValue(true);
        System.out.println("金額 D2:D4 = "
                + String.join(", ", values.get("D2"), values.get("D3"), values.get("D4")));
        System.out.println("合計: " + values.get("D5"));

        // --- 保存して閉じる ---
        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// values-formulas.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("sales-php.xlsx");   // 新規ブック作成
$ws = $wb-&gt;openWorkSheet("Sheet1");             // 既定のシートを開く

// --- 見出し (1行目) ---
// $app-&gt;convertFromRowColNumber(行, 列) で (1,1) -&gt; "A1" のセルアドレスが得られる
$headers = ["商品", "単価", "個数", "金額"];
foreach ($headers as $i =&gt; $text) {
    $cell = $app-&gt;convertFromRowColNumber(1, $i + 1);
    $ws-&gt;getRange("{$cell}:{$cell}")-&gt;setValue($text);
}

// --- 明細 (2〜4行目) ---
$data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
foreach ($data as $i =&gt; [$name, $price, $qty]) {
    $row = $i + 2;
    $ws-&gt;getRange("A{$row}:A{$row}")-&gt;setValue($name);
    $ws-&gt;getRange("B{$row}:B{$row}")-&gt;setNumberValue($price);
    $ws-&gt;getRange("C{$row}:C{$row}")-&gt;setNumberValue($qty);
}

// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
$formulas = [];
for ($i = 0; $i &lt; count($data); $i++) {
    $f = new \phposbxl\dto\InputFormulaObject();
    $row = $i + 2;
    $f-&gt;setA1C1($app-&gt;convertFromRowColNumber($i + 2, 4));   // D2, D3, D4
    $f-&gt;setFormula("B{$row}*C{$row}");                       // "B2*C2" ...
    $formulas[] = $f;
}
$ws-&gt;setFormulaArray($formulas);

// --- 合計 (D5) ---
$ws-&gt;getRange("C5:C5")-&gt;setValue("合計");
$ws-&gt;getRange("D5:D5")-&gt;setFormula("SUM(D2:D4)");

// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
$values = $ws-&gt;getRange("A1:D5")-&gt;getValue(true);
echo "金額 D2:D4 = " . implode(", ", [$values["D2"], $values["D3"], $values["D4"]]) . "\n";
echo "合計: " . $values["D5"] . "\n";

// --- 保存して閉じる ---
$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// values-formulas.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("sales-wasm.xlsx");   // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1");              // 既定のシートを開く

// --- 見出し (1行目) ---
// app.convertFromRowColNumber(行, 列) で (1,1) -&gt; "A1" のセルアドレスが得られる
["商品", "単価", "個数", "金額"].forEach((text, i) =&gt; {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- 明細 (2〜4行目) ---
const data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
data.forEach(([name, price, qty], i) =&gt; {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setFormulaArray を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((_, i) =&gt; {
    const f = new dto.InputFormulaObject();
    f.setA1C1(app.convertFromRowColNumber(i + 2, 4));  // D2, D3, D4
    f.setFormula(`B${i + 2}*C${i + 2}`);               // "B2*C2" ...
    return f;
});
ws.setFormulaArray(formulas);

// --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
const values = ws.getRange("A1:D5").getValue(true);
console.log("金額 D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("合計:", values["D5"]);

// --- 保存して閉じる ---
wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">金額 D2:D4 = 300, 300, 300
合計: 900</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">new nodeosbxl.App()</code> がすべての入口。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">App</code> から <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">createWorkBook()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkBook()</code> で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkBook</code> を取得し、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet()</code> で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code>、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getRange()</code> で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Range</code> を辿ります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>セルアドレスは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:B2"</code> 形式の文字列。</strong> 単一セルでも <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A1:A1"</code> と範囲の形で渡します。<br> 行番号・列番号（1始まり）から作る場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber(row, col)</code>、 範囲なら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertFromRowColNumber2(startRow, startCol, endRow, endCol)</code> を使います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>値の種類ごとに setter が分かれています。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(文字列)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(数値)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(DateTimeObject)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue(日付文字列)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(真偽値)</code>。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> は「生値」</strong>（数式なら計算結果）。<br>省略すると表示値（書式適用後）になります。<br> 戻り値は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ "A1": "100", ... }</code> のマップで、値はすべて文字列です。<br>数値比較する場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Number(...)</code> で変換します。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.save()</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.close()</code> で終了。</strong> 別ファイルに書き出したいときは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.saveAs(path)</code> を使いますが、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">saveAs()</code> は<strong>以降の保存先を切り替えません</strong>（<a href="general-01-overview.php#注意点" class="text-primary-600 hover:underline">概要の注意点</a>参照）。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>数式に先頭の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> を付けない。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("=SUM(A1:A3)")</code> ではなく <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code> です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>複数セル Range への <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue</code> は「全セルに同じ値」を入れます。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("A1:C1").setValue("X")</code> → A1, B1, C1 すべて <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"X"</code> になります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula()</code> は相対参照を自動展開しません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ws.getRange("B1:B3").setFormula("A1*2")</code> は <strong>B1 だけ</strong>に数式が入り、B2 / B3 は空のままです。<br> 第3引数 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setAllCell</code> を true にした <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("A1*2", false, true)</code> も 「同じ文字列をそのまま全セルにコピーする」だけで、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A2*2</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">A3*2</code> には<strong>なりません</strong> （結果 B1:B3 がすべて同じ値になります）。<br> 行ごとにずらした数式を入れたい場合は、このサンプルのように <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormulaArray()</code> + <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">dto.InputFormulaObject</code>（セルごとに A1C1 と数式を指定）</strong> を使ってください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>起動時にライセンス表示が stdout に 1 行出力されます</strong>（プロセスにつき 1 回）。<br> 標準出力をパースする処理を書く場合は注意してください。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-01-overview.php" class="text-primary-600 hover:underline">概要 — オブジェクトモデルと基本の流れ</a> / 次: <a href="general-03-sheets.php" class="text-primary-600 hover:underline">ブックとワークシートの操作</a></li>
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
   "headline": "値と数式を入れて表を作る",
   "description": "値と数式を入れて表を作る。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-02-values-formulas.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-02-values-formulas.php",
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
     "name": "値と数式を入れて表を作る",
     "item": "https://osb-office.com/sample/general-02-values-formulas.php"
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
