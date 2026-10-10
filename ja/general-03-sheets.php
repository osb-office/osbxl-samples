<?php
$page_title = 'ブックとワークシートの操作 — 追加・削除・移動・リネーム | osbxl サンプル';
$no_tagline = true;
$meta_description = 'ブックとワークシートの操作 — 追加・削除・移動・リネーム。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,worksheet,addworksheet,deleteworksheet,movesheet,setname,openworksheetbyindex,getsheetcount,activate';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">ブックとワークシートの操作 — 追加・削除・移動・リネーム</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-03-sheets.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">1 つのブックの中でワークシートを追加・削除・移動・リネームし、 インデックスでシートを辿ったり、開いたときに表示されるシート（アクティブタブ）を制御する。</p>
<h2 id="対象メソッド" class="text-2xl font-bold mt-10 mb-4">対象メソッド</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">所属</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">説明</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet(sheetName, position?)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを追加します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを削除します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet(sheetName, position)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを移動します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートを名前で開きます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex(sheetIndex)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートをインデックスで開きます。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシート数を取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetNumber(sheetName)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkBook</td><td class="border border-surface-200 px-3 py-2 align-top">ワークシートのシート番号を取得します。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getName()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setName(name)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet</td><td class="border border-surface-200 px-3 py-2 align-top">シート名称の取得 / 設定を行います。</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isActive()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(isActive)</code></td><td class="border border-surface-200 px-3 py-2 align-top">WorkSheet</td><td class="border border-surface-200 px-3 py-2 align-top">アクティブタブ（起動時の表示項目）の取得 / 設定を行います。</td></tr>
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
        <pre><code class="language-javascript">// sheets.js — シートの追加・移動・リネーム・削除と、インデックスアクセス
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sheets.xlsx");

// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
const show = (label) =&gt; {
    const names = [];
    for (let i = 1; i &lt;= wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheetByIndex(i).getName());
    }
    console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};

show("[1] 新規作成直後        :");

wb.addWorkSheet("集計");                       // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :");

wb.addWorkSheet("受注", 1);                    // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");

wb.moveSheet("受注", 3);                       // 3 番目へ移動
show("[4] moveSheet(受注,3)   :");

const ws = wb.openWorkSheet("受注");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"));
ws.getRange("A1:A1").setValue("受注データ");

wb.openWorkSheet("Sheet1").setName("明細");     // リネーム
show("[6] setName(明細)       :");

// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
console.log("[7] isActive: 集計 =", wb.openWorkSheet("集計").isActive(),
            "/ 明細 =", wb.openWorkSheet("明細").isActive());
wb.openWorkSheet("集計").activate(true);   // 他のシートのアクティブは自動的に解除されます
console.log("    activate後: 集計 =", wb.openWorkSheet("集計").isActive(),
            "/ 明細 =", wb.openWorkSheet("明細").isActive());

// --- 削除 ---
wb.addWorkSheet("一時");
wb.deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");

wb.save();
wb.close();

// --- 開き直して確認 ---
const wb2 = app.openWorkBook("sheets.xlsx");
const names = [];
for (let i = 1; i &lt;= wb2.getSheetCount(); i++) names.push(wb2.openWorkSheetByIndex(i).getName());
console.log("[終] シート構成:", names.join(", "));
console.log("     受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"]);
wb2.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># sheets.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sheets-py.xlsx")

def b(x):
    # python の True/False を node の true/false 表記に合わせる
    return "true" if x else "false"

# インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
def show(label):
    names = [wb.openWorkSheetByIndex(i).getName() for i in range(1, wb.getSheetCount() + 1)]
    print(f"{label} count={wb.getSheetCount()} [{', '.join(names)}]")

show("[1] 新規作成直後        :")

wb.addWorkSheet("集計")                       # position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :")

wb.addWorkSheet("受注", 1)                    # position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):")

wb.moveSheet("受注", 3)                       # 3 番目へ移動
show("[4] moveSheet(受注,3)   :")

ws = wb.openWorkSheet("受注")
print("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"))
ws.getRange("A1:A1").setValue("受注データ")

wb.openWorkSheet("Sheet1").setName("明細")    # リネーム
show("[6] setName(明細)       :")

# --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
print("[7] isActive: 集計 =", b(wb.openWorkSheet("集計").isActive()),
      "/ 明細 =", b(wb.openWorkSheet("明細").isActive()))
wb.openWorkSheet("集計").activate(True)       # 他のシートのアクティブは自動的に解除されます
print("    activate後: 集計 =", b(wb.openWorkSheet("集計").isActive()),
      "/ 明細 =", b(wb.openWorkSheet("明細").isActive()))

# --- 削除 ---
wb.addWorkSheet("一時")
wb.deleteWorkSheet("一時")
show("[8] deleteWorkSheet(一時):")

wb.save()
wb.close()

# --- 開き直して確認 ---
wb2 = app.openWorkBook("sheets-py.xlsx")
names = [wb2.openWorkSheetByIndex(i).getName() for i in range(1, wb2.getSheetCount() + 1)]
print("[終] シート構成:", ", ".join(names))
print("     受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"])
wb2.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// sheets.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    static WorkBookWrapper wb;

    // インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
    static void show(String label) {
        List&lt;String&gt; names = new ArrayList&lt;&gt;();
        for (int i = 1; i &lt;= wb.getSheetCount(); i++) {
            names.add(wb.openWorkSheetByIndex(i).getName());
        }
        System.out.println(label + " count=" + wb.getSheetCount()
                + " [" + String.join(", ", names) + "]");
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        wb = app.createWorkBook("sheets-java.xlsx");

        show("[1] 新規作成直後        :");

        wb.addWorkSheet("集計");                       // position 省略 → 末尾に追加
        show("[2] addWorkSheet(集計)  :");

        wb.addWorkSheet("受注", 1);                    // position=1 → 先頭に挿入
        show("[3] addWorkSheet(受注,1):");

        wb.moveSheet("受注", 3);                       // 3 番目へ移動
        show("[4] moveSheet(受注,3)   :");

        WorkSheetWrapper ws = wb.openWorkSheet("受注");
        System.out.println("[5] getName() = " + ws.getName()
                + " / getSheetNumber(受注) = " + wb.getSheetNumber("受注"));
        ws.getRange("A1:A1").setValue("受注データ");

        wb.openWorkSheet("Sheet1").setName("明細");    // リネーム
        show("[6] setName(明細)       :");

        // --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
        System.out.println("[7] isActive: 集計 = " + wb.openWorkSheet("集計").isActive()
                + " / 明細 = " + wb.openWorkSheet("明細").isActive());
        wb.openWorkSheet("集計").activate(true);       // 他のシートのアクティブは自動的に解除されます
        System.out.println("    activate後: 集計 = " + wb.openWorkSheet("集計").isActive()
                + " / 明細 = " + wb.openWorkSheet("明細").isActive());

        // --- 削除 ---
        wb.addWorkSheet("一時");
        wb.deleteWorkSheet("一時");
        show("[8] deleteWorkSheet(一時):");

        wb.save();
        wb.close();

        // --- 開き直して確認 ---
        WorkBookWrapper wb2 = app.openWorkBook("sheets-java.xlsx");
        List&lt;String&gt; names = new ArrayList&lt;&gt;();
        for (int i = 1; i &lt;= wb2.getSheetCount(); i++) names.add(wb2.openWorkSheetByIndex(i).getName());
        System.out.println("[終] シート構成: " + String.join(", ", names));
        Map&lt;String, String&gt; v = wb2.openWorkSheet("受注").getRange("A1:A1").getValue();
        System.out.println("     受注!A1 = " + v.get("A1"));
        wb2.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// sheets.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("sheets-php.xlsx");

// PHP の bool を node の true/false 表記に合わせる
function b($x) {
    return $x ? "true" : "false";
}

// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
function show($label) {
    global $wb;
    $names = [];
    for ($i = 1; $i &lt;= $wb-&gt;getSheetCount(); $i++) {
        $names[] = $wb-&gt;openWorkSheetByIndex($i)-&gt;getName();
    }
    echo $label . " count=" . $wb-&gt;getSheetCount() . " [" . implode(", ", $names) . "]\n";
}

show("[1] 新規作成直後        :");

$wb-&gt;addWorkSheet("集計");                       // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :");

$wb-&gt;addWorkSheet("受注", 1);                    // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");

$wb-&gt;moveSheet("受注", 3);                       // 3 番目へ移動
show("[4] moveSheet(受注,3)   :");

$ws = $wb-&gt;openWorkSheet("受注");
echo "[5] getName() = " . $ws-&gt;getName() . " / getSheetNumber(受注) = " . $wb-&gt;getSheetNumber("受注") . "\n";
$ws-&gt;getRange("A1:A1")-&gt;setValue("受注データ");

$wb-&gt;openWorkSheet("Sheet1")-&gt;setName("明細");   // リネーム
show("[6] setName(明細)       :");

// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
echo "[7] isActive: 集計 = " . b($wb-&gt;openWorkSheet("集計")-&gt;isActive())
   . " / 明細 = " . b($wb-&gt;openWorkSheet("明細")-&gt;isActive()) . "\n";
$wb-&gt;openWorkSheet("集計")-&gt;activate(true);      // 他のシートのアクティブは自動的に解除されます
echo "    activate後: 集計 = " . b($wb-&gt;openWorkSheet("集計")-&gt;isActive())
   . " / 明細 = " . b($wb-&gt;openWorkSheet("明細")-&gt;isActive()) . "\n";

// --- 削除 ---
$wb-&gt;addWorkSheet("一時");
$wb-&gt;deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");

$wb-&gt;save();
$wb-&gt;close();

// --- 開き直して確認 ---
$wb2 = $app-&gt;openWorkBook("sheets-php.xlsx");
$names = [];
for ($i = 1; $i &lt;= $wb2-&gt;getSheetCount(); $i++) {
    $names[] = $wb2-&gt;openWorkSheetByIndex($i)-&gt;getName();
}
echo "[終] シート構成: " . implode(", ", $names) . "\n";
$v = $wb2-&gt;openWorkSheet("受注")-&gt;getRange("A1:A1")-&gt;getValue();
echo "     受注!A1 = " . $v["A1"] . "\n";
$wb2-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// sheets.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("sheets-wasm.xlsx");

// インデックスは 1 始まり。1..getSheetCount() を順に開いて名前を並べる
const show = (label) =&gt; {
    const names = [];
    for (let i = 1; i &lt;= wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheetByIndex(i).getName());
    }
    console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};

show("[1] 新規作成直後        :");

wb.addWorkSheet("集計");                       // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :");

wb.addWorkSheet("受注", 1);                    // position=1 → 先頭に挿入
show("[3] addWorkSheet(受注,1):");

wb.moveSheet("受注", 3);                       // 3 番目へ移動
show("[4] moveSheet(受注,3)   :");

const ws = wb.openWorkSheet("受注");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"));
ws.getRange("A1:A1").setValue("受注データ");

wb.openWorkSheet("Sheet1").setName("明細");     // リネーム
show("[6] setName(明細)       :");

// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
console.log("[7] isActive: 集計 =", wb.openWorkSheet("集計").isActive(),
            "/ 明細 =", wb.openWorkSheet("明細").isActive());
wb.openWorkSheet("集計").activate(true);   // 他のシートのアクティブは自動的に解除されます
console.log("    activate後: 集計 =", wb.openWorkSheet("集計").isActive(),
            "/ 明細 =", wb.openWorkSheet("明細").isActive());

// --- 削除 ---
wb.addWorkSheet("一時");
wb.deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");

wb.save();
wb.close();

// --- 開き直して確認 ---
const wb2 = app.openWorkBook("sheets-wasm.xlsx");
const names = [];
for (let i = 1; i &lt;= wb2.getSheetCount(); i++) names.push(wb2.openWorkSheetByIndex(i).getName());
console.log("[終] シート構成:", names.join(", "));
console.log("     受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"]);
wb2.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[1] 新規作成直後        : count=1 [Sheet1]
[2] addWorkSheet(集計)  : count=2 [Sheet1, 集計]
[3] addWorkSheet(受注,1): count=3 [受注, Sheet1, 集計]
[4] moveSheet(受注,3)   : count=3 [Sheet1, 集計, 受注]
[5] getName() = 受注 / getSheetNumber(受注) = 3
[6] setName(明細)       : count=3 [明細, 集計, 受注]
[7] isActive: 集計 = false / 明細 = true
    activate後: 集計 = true / 明細 = false
[8] deleteWorkSheet(一時): count=3 [明細, 集計, 受注]
[終] シート構成: 明細, 集計, 受注
     受注!A1 = 受注データ</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> とシート番号はどちらも 1 始まり。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet(name, position)</code> は <strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> の前に挿入</strong>します（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> = 1 なら先頭）。<br>省略（既定値 -1）で末尾に追加されます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet(name, position)</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">position</code> は「移動後の位置」</strong>で、先頭が 1 です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex(i)</code> も 1 始まり。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code> と組み合わせると 全シートを名前で列挙できます（上記の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">show()</code>）。<br> シート名が分からないブックを扱うときは、この形で辿るのが確実です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetNumber(name)</code> は現在の並び順での番号</strong>を返します。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">moveSheet()</code> や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteWorkSheet()</code> の後は値が変わります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setName()</code> は即座に反映されます。</strong> 以降は新しい名前で <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet()</code> します。<br> 古い名前で開こうとすると例外になります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(true)</code> は排他的に働きます。</strong> 目的のシートを <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(true)</code> にすると、 他のシートのアクティブは<strong>自動的に解除されます</strong>（実測）。<br> あらかじめ他を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(false)</code> にする必要はありません。<br> アクティブ状態は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">save()</code> して開き直しても保持されます。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>存在しないシート名を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheet()</code> に渡すと例外になります。</strong><div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">wb.openWorkSheet("NoSuch");   // → 例外: "NoSuch is not found"</code></pre></div>ユーザー入力や外部設定からシート名を受ける場合は、事前に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code> + <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex()</code> で名前を列挙して照合してください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">openWorkSheetByIndex()</code> の範囲外インデックスも例外になります。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1</code> 〜 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getSheetCount()</code> の範囲で呼び出してください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>新規作成直後は既定シート（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Sheet1</code>）だけが <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">isActive() === true</code></strong> です（実測）。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">addWorkSheet()</code> で追加したシートは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">false</code> になります。<br> ブックを開いたときに別のシートを手前に見せたいなら、そのシートで <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">activate(true)</code> を呼んでください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>シートを跨いだセルコピーは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCol</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell</code> を使います</strong> （コピー元をシート名で指定）。<br>詳細は <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">行・列・セル範囲の操作</a>。<br> 図形（グラフ含む）・ピボットテーブル・queryTable 形式のテーブル・データテーブルは コピー対象に含まれません。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-02-values-formulas.php" class="text-primary-600 hover:underline">値と数式を入れて表を作る</a> / 次: <a href="general-04-value-types.php" class="text-primary-600 hover:underline">セル値の型と表示書式</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/WorkBookWrapper.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.WorkBook.html" class="text-primary-600 hover:underline">WorkBook</a> / <a href="/api-docs/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "ブックとワークシートの操作 — 追加・削除・移動・リネーム",
   "description": "ブックとワークシートの操作 — 追加・削除・移動・リネーム。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-03-sheets.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-03-sheets.php",
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
     "name": "ブックとワークシートの操作 — 追加・削除・移動・リネーム",
     "item": "https://osb-office.com/sample/general-03-sheets.php"
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
