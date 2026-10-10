<?php
$page_title = '行・列・セル範囲の操作 — 挿入・削除・クリア・コピー | osbxl サンプル';
$no_tagline = true;
$meta_description = '行・列・セル範囲の操作 — 挿入・削除・クリア・コピー。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,insertrow,deleterow,clearrow,copyrow,insertcol,deletecol,insertcell,deletecell,clearcell,copycell,row,col,setheight,setcolumnwidth';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">行・列・セル範囲の操作 — 挿入・削除・クリア・コピー</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-05-rows-cols.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">ワークシート上の行・列・セル範囲を、挿入 / 削除 / クリア / コピーする。 あわせて行の高さ・列の幅・非表示などを <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col</code> オブジェクトで設定する。</p>
<h2 id="4-つの操作の違い" class="text-2xl font-bold mt-10 mb-4">4 つの操作の違い</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">操作</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">挙動</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">行・列の場合</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">セル範囲の場合</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>insert</strong></td><td class="border border-surface-200 px-3 py-2 align-top">空白を挿入し、既存データを<strong>ずらす</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertRow(start, n)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCol(start, n)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell(A1C1, direction)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>delete</strong></td><td class="border border-surface-200 px-3 py-2 align-top">削除して、既存データを<strong>ずらす</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteRow(start, n)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCol(start, n)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell(A1C1, direction)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>clear</strong></td><td class="border border-surface-200 px-3 py-2 align-top">中身と書式を初期状態に戻す。<strong>ずらさない</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow(start, n)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCol(start, n)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearCell(A1C1)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><strong>copy</strong></td><td class="border border-surface-200 px-3 py-2 align-top">別シート（または自シート）から<strong>上書きコピー</strong></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow(fromSheet, fromRow, n, toRow, copyType?)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell(fromSheet, fromA1C1, toA1, copyType?)</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlInsertDirection</code> を使います。</p>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">数値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">用途</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionRight</code></td><td class="border border-surface-200 px-3 py-2 align-top">1</td><td class="border border-surface-200 px-3 py-2 align-top">セル挿入時に右へずらす</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionDown</code></td><td class="border border-surface-200 px-3 py-2 align-top">2</td><td class="border border-surface-200 px-3 py-2 align-top">セル挿入時に下へずらす</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionLeft</code></td><td class="border border-surface-200 px-3 py-2 align-top">3</td><td class="border border-surface-200 px-3 py-2 align-top">セル削除時に左へずらす</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionUp</code></td><td class="border border-surface-200 px-3 py-2 align-top">4</td><td class="border border-surface-200 px-3 py-2 align-top">セル削除時に上へずらす</td></tr>
  </tbody></table></div>
<h2 id="コード-1-行・列・セル範囲の操作" class="text-2xl font-bold mt-10 mb-4">コード 1: 行・列・セル範囲の操作</h2>
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
        <pre><code class="language-javascript">// rows-cols.js — 行・列・セル範囲の挿入/削除/クリア/コピー
const { nodeosbxl, enums } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("rows-cols.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
const dump = (label, range) =&gt; {
    const v = ws.getRange(range).getValue();
    const keys = Object.keys(v).map((k) =&gt; /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =&gt;
        Number(a[2]) - Number(b[2]) ||
        app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
    console.log(`${label} ${keys.map((m) =&gt; `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// A1:C5 を r{行}c{列} で埋める
const fill = () =&gt; {
    for (let r = 1; r &lt;= 5; r++) for (let c = 1; c &lt;= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};

fill();
dump("[初期]       A列:", "A1:A5");

// ---------- 行 ----------
ws.insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow]  A列:", "A1:A7");
ws.deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow]  A列:", "A1:A5");
ws.clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
dump("[clearRow]   A列:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow]    A列:", "A1:A5");

fill();
// ---------- 列 ----------
ws.insertCol(2, 1);                    // B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1");
ws.deleteCol(2, 1);                    // 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1");
ws.clearCol(3, 1);                     // C 列をクリア
dump("[clearCol]  1行目:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
dump("[copyCol]   1行目:", "A1:F1");

// ---------- セル範囲 ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] B列:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] B列:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell]  F8:G9:", "F8:G9");

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># rows-cols.py — node 版サンプルと同一検証 (pyosbxl)
import re
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("rows-cols-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# 行 → 列 の順に並べて 1 行で表示する（· は空セル）
def dump(label, rng):
    v = ws.getRange(rng).getValue()
    def keyf(k):
        m = re.match(r"^([A-Z]+)(\d+)$", k)
        return (int(m.group(2)), app.convertToColumnNumber(m.group(1)))
    keys = sorted(v.keys(), key=keyf)
    print(label, " ".join(f"{k}={v[k] or '·'}" for k in keys))

# A1:C5 を r{行}c{列} で埋める
def fill():
    for r in range(1, 6):
        for c in range(1, 4):
            ws.getCells(r, c).setValue(f"r{r}c{c}")

fill()
dump("[初期]       A列:", "A1:A5")

# ---------- 行 ----------
ws.insertRow(3, 2)                     # 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow]  A列:", "A1:A7")
ws.deleteRow(3, 2)                     # 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow]  A列:", "A1:A5")
ws.clearRow(2, 1)                      # 2 行目をクリア（位置はずらさない）
dump("[clearRow]   A列:", "A1:A5")
ws.copyRow("Sheet1", 1, 1, 4)          # 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow]    A列:", "A1:A5")

fill()
# ---------- 列 ----------
ws.insertCol(2, 1)                     # B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1")
ws.deleteCol(2, 1)                     # 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1")
ws.clearCol(3, 1)                      # C 列をクリア
dump("[clearCol]  1行目:", "A1:C1")
fill()
ws.copyCol("Sheet1", 1, 2, 5)          # A,B 列を E,F 列へコピー
dump("[copyCol]   1行目:", "A1:F1")

# ---------- セル範囲 ----------
ws.insertCell("B2:B3", pyosbxl.enums.XlInsertDirection.InsertDirectionDown)
dump("[insertCell] B列:", "B1:B5")
ws.deleteCell("B2:B3", pyosbxl.enums.XlInsertDirection.DeleteDirectionUp)
dump("[deleteCell] B列:", "B1:B5")
ws.clearCell("A1:B2")
dump("[clearCell] A1:B2:", "A1:B2")
ws.copyCell("Sheet1", "A3:B4", "F8")
dump("[copyCell]  F8:G9:", "F8:G9")

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// rows-cols.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    static AppWrapper app;
    static WorkSheetWrapper ws;
    static final Pattern ADDR = Pattern.compile("^([A-Z]+)(\\d+)$");

    // 行 → 列 の順に並べて 1 行で表示する（· は空セル）
    static void dump(String label, String range) {
        Map&lt;String, String&gt; v = ws.getRange(range).getValue();
        List&lt;String&gt; keys = new ArrayList&lt;&gt;(v.keySet());
        keys.sort((a, b) -&gt; {
            Matcher ma = ADDR.matcher(a);
            Matcher mb = ADDR.matcher(b);
            ma.matches();
            mb.matches();
            int cmp = Integer.compare(Integer.parseInt(ma.group(2)), Integer.parseInt(mb.group(2)));
            if (cmp != 0) return cmp;
            return Integer.compare(app.convertToColumnNumber(ma.group(1)),
                                   app.convertToColumnNumber(mb.group(1)));
        });
        StringBuilder sb = new StringBuilder();
        for (String k : keys) {
            if (sb.length() &gt; 0) sb.append(" ");
            String val = v.get(k);
            sb.append(k).append("=").append(val == null || val.isEmpty() ? "·" : val);
        }
        System.out.println(label + " " + sb);
    }

    // A1:C5 を r{行}c{列} で埋める
    static void fill() {
        for (int r = 1; r &lt;= 5; r++)
            for (int c = 1; c &lt;= 3; c++)
                ws.getCells(r, c).setValue("r" + r + "c" + c);
    }

    public static void main(String[] args) {
        app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("rows-cols-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        fill();
        dump("[初期]       A列:", "A1:A5");

        // ---------- 行 ----------
        ws.insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
        dump("[insertRow]  A列:", "A1:A7");
        ws.deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
        dump("[deleteRow]  A列:", "A1:A5");
        ws.clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
        dump("[clearRow]   A列:", "A1:A5");
        ws.copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
        dump("[copyRow]    A列:", "A1:A5");

        fill();
        // ---------- 列 ----------
        ws.insertCol(2, 1);                    // B 列の前に 1 列挿入
        dump("[insertCol] 1行目:", "A1:D1");
        ws.deleteCol(2, 1);                    // 挿入した B 列を削除
        dump("[deleteCol] 1行目:", "A1:C1");
        ws.clearCol(3, 1);                     // C 列をクリア
        dump("[clearCol]  1行目:", "A1:C1");
        fill();
        ws.copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
        dump("[copyCol]   1行目:", "A1:F1");

        // ---------- セル範囲 ----------
        ws.insertCell("B2:B3", XlInsertDirection.InsertDirectionDown);
        dump("[insertCell] B列:", "B1:B5");
        ws.deleteCell("B2:B3", XlInsertDirection.DeleteDirectionUp);
        dump("[deleteCell] B列:", "B1:B5");
        ws.clearCell("A1:B2");
        dump("[clearCell] A1:B2:", "A1:B2");
        ws.copyCell("Sheet1", "A3:B4", "F8");
        dump("[copyCell]  F8:G9:", "F8:G9");

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// rows-cols.php — 行・列・セル範囲の挿入/削除/クリア/コピー (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("rows-cols-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
$dump = function (string $label, string $range) use ($ws, $app): void {
    $v = $ws-&gt;getRange($range)-&gt;getValue();
    $keys = array_keys($v);
    usort($keys, function (string $a, string $b) use ($app): int {
        preg_match('/^([A-Z]+)(\d+)$/', $a, $ma);
        preg_match('/^([A-Z]+)(\d+)$/', $b, $mb);
        return ((int)$ma[2] &lt;=&gt; (int)$mb[2])
            ?: ($app-&gt;convertToColumnNumber($ma[1]) &lt;=&gt; $app-&gt;convertToColumnNumber($mb[1]));
    });
    $parts = [];
    foreach ($keys as $k) {
        $val = (string)$v[$k];
        $parts[] = $k . "=" . ($val === "" ? "·" : $val);
    }
    echo $label, " ", implode(" ", $parts), "\n";
};
// A1:C5 を r{行}c{列} で埋める
$fill = function () use ($ws): void {
    for ($r = 1; $r &lt;= 5; $r++) {
        for ($c = 1; $c &lt;= 3; $c++) {
            $ws-&gt;getCells($r, $c)-&gt;setValue("r{$r}c{$c}");
        }
    }
};

$fill();
$dump("[初期]       A列:", "A1:A5");

// ---------- 行 ----------
$ws-&gt;insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
$dump("[insertRow]  A列:", "A1:A7");
$ws-&gt;deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
$dump("[deleteRow]  A列:", "A1:A5");
$ws-&gt;clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
$dump("[clearRow]   A列:", "A1:A5");
$ws-&gt;copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
$dump("[copyRow]    A列:", "A1:A5");

$fill();
// ---------- 列 ----------
$ws-&gt;insertCol(2, 1);                    // B 列の前に 1 列挿入
$dump("[insertCol] 1行目:", "A1:D1");
$ws-&gt;deleteCol(2, 1);                    // 挿入した B 列を削除
$dump("[deleteCol] 1行目:", "A1:C1");
$ws-&gt;clearCol(3, 1);                     // C 列をクリア
$dump("[clearCol]  1行目:", "A1:C1");
$fill();
$ws-&gt;copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
$dump("[copyCol]   1行目:", "A1:F1");

// ---------- セル範囲 ----------
$ws-&gt;insertCell("B2:B3", \phposbxl\enums\XlInsertDirection::InsertDirectionDown);
$dump("[insertCell] B列:", "B1:B5");
$ws-&gt;deleteCell("B2:B3", \phposbxl\enums\XlInsertDirection::DeleteDirectionUp);
$dump("[deleteCell] B列:", "B1:B5");
$ws-&gt;clearCell("A1:B2");
$dump("[clearCell] A1:B2:", "A1:B2");
$ws-&gt;copyCell("Sheet1", "A3:B4", "F8");
$dump("[copyCell]  F8:G9:", "F8:G9");

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// rows-cols.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("rows-cols-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
const dump = (label, range) =&gt; {
    const v = ws.getRange(range).getValue();
    const keys = Object.keys(v).map((k) =&gt; /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =&gt;
        Number(a[2]) - Number(b[2]) ||
        app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
    console.log(`${label} ${keys.map((m) =&gt; `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// A1:C5 を r{行}c{列} で埋める
const fill = () =&gt; {
    for (let r = 1; r &lt;= 5; r++) for (let c = 1; c &lt;= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};

fill();
dump("[初期]       A列:", "A1:A5");

// ---------- 行 ----------
ws.insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow]  A列:", "A1:A7");
ws.deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow]  A列:", "A1:A5");
ws.clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
dump("[clearRow]   A列:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow]    A列:", "A1:A5");

fill();
// ---------- 列 ----------
ws.insertCol(2, 1);                    // B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1");
ws.deleteCol(2, 1);                    // 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1");
ws.clearCol(3, 1);                     // C 列をクリア
dump("[clearCol]  1行目:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
dump("[copyCol]   1行目:", "A1:F1");

// ---------- セル範囲 ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] B列:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] B列:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell]  F8:G9:", "F8:G9");

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[初期]       A列: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[insertRow]  A列: A1=r1c1 A2=r2c1 A3=· A4=· A5=r3c1 A6=r4c1 A7=r5c1
[deleteRow]  A列: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[clearRow]   A列: A1=r1c1 A2=· A3=r3c1 A4=r4c1 A5=r5c1
[copyRow]    A列: A1=r1c1 A2=· A3=r3c1 A4=r1c1 A5=r5c1
[insertCol] 1行目: A1=r1c1 B1=· C1=r1c2 D1=r1c3
[deleteCol] 1行目: A1=r1c1 B1=r1c2 C1=r1c3
[clearCol]  1行目: A1=r1c1 B1=r1c2 C1=·
[copyCol]   1行目: A1=r1c1 B1=r1c2 C1=r1c3 D1=· E1=r1c1 F1=r1c2
[insertCell] B列: B1=r1c2 B2=· B3=· B4=r2c2 B5=r3c2
[deleteCell] B列: B1=r1c2 B2=r2c2 B3=r3c2 B4=r4c2 B5=r5c2
[clearCell] A1:B2: A1=· B1=· A2=· B2=·
[copyCell]  F8:G9: F8=r3c1 G8=r3c2 F9=r4c1 G9=r4c2</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">·</code> は空セルです。<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertRow</code> で 3・4 行目が空白になり既存データが下へずれたこと、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow</code> では位置がずれていないこと、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> で 1 行目の内容が 4 行目に上書きされたことが読み取れます。</p>
<h2 id="コード-2-row-col-オブジェクト" class="text-2xl font-bold mt-10 mb-4">コード 2: <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col</code> オブジェクト</h2>
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
        <pre><code class="language-javascript">// row-col-props.js — 行の高さ・列の幅・非表示・自動調整
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("row-col-props.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- Row ----------
const row = ws.getRow(1);                        // 1 行目
console.log("既定の高さ      =", row.getHeight());
row.setHeight(30);
console.log("setHeight(30)後 =", row.getHeight());
console.log("getHidden 既定  =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true);                    // 2 行目を非表示
console.log("2行目 hidden    =", ws.getRow(2).getHidden());

// ---------- Col ----------
const col = ws.getCol(1);                        // 1 列目（A 列）。列番号は数値で渡す
console.log("既定の列幅      = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -&gt; columnWidth=" + col.getColumnWidth(),
            "/ width=" + col.getWidth());
ws.getCol(2).setHidden(true);                    // B 列を非表示
console.log("B列 hidden      =", ws.getCol(2).getHidden());

ws.getRange("C1:C1").setValue("自動調整のサンプル文字列");
console.log("C列 bestFit 既定 =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true);                   // 内容に合わせて列幅を自動調整
console.log("setBestFit(true) -&gt; bestFit =", ws.getCol(3).getBestFit(),
            "/ columnWidth =", ws.getCol(3).getColumnWidth());

// アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1);                       // 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1);                       // B〜C 列をレベル 1 でグループ化
console.log("4行目 outlineLevel =", ws.getRow(4).getOutlineLevel(),
            "/ B列 outlineLevel =", ws.getCol(2).getOutlineLevel());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># row-col-props.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("row-col-props-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

def num(x):          # 整数値の float は int 表記にする（JS の出力と揃える）
    return int(x) if isinstance(x, float) and x.is_integer() else x

def boolstr(x):      # Python の True/False を JS 風の true/false で表示する
    return str(x).lower()

# ---------- Row ----------
row = ws.getRow(1)                           # 1 行目
print("既定の高さ      =", num(row.getHeight()))
row.setHeight(30)
print("setHeight(30)後 =", num(row.getHeight()))
print("getHidden 既定  =", boolstr(row.getHidden()), "/ getOutlineLevel =", row.getOutlineLevel())
ws.getRow(2).setHidden(True)                 # 2 行目を非表示
print("2行目 hidden    =", boolstr(ws.getRow(2).getHidden()))

# ---------- Col ----------
col = ws.getCol(1)                           # 1 列目（A 列）。列番号は数値で渡す
print("既定の列幅      = columnWidth:", col.getColumnWidth(), "/ width:", num(col.getWidth()))
col.setColumnWidth(20)
print("setColumnWidth(20) -&gt; columnWidth=" + str(num(col.getColumnWidth())),
      "/ width=" + str(num(col.getWidth())))
ws.getCol(2).setHidden(True)                 # B 列を非表示
print("B列 hidden      =", boolstr(ws.getCol(2).getHidden()))

ws.getRange("C1:C1").setValue("自動調整のサンプル文字列")
print("C列 bestFit 既定 =", boolstr(ws.getCol(3).getBestFit()))
ws.getCol(3).setBestFit(True)                # 内容に合わせて列幅を自動調整
print("setBestFit(true) -&gt; bestFit =", boolstr(ws.getCol(3).getBestFit()),
      "/ columnWidth =", ws.getCol(3).getColumnWidth())

# アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1)                    # 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1)                    # B〜C 列をレベル 1 でグループ化
print("4行目 outlineLevel =", ws.getRow(4).getOutlineLevel(),
      "/ B列 outlineLevel =", ws.getCol(2).getOutlineLevel())

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// row-col-props.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;

public class Main {
    // 整数値の double は int 表記にする（JS の出力と揃える）
    static String num(double d) {
        if (d == Math.rint(d) &amp;&amp; !Double.isInfinite(d)) return String.valueOf((long) d);
        return String.valueOf(d);
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("row-col-props-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- Row ----------
        RowWrapper row = ws.getRow(1);                   // 1 行目
        System.out.println("既定の高さ      = " + num(row.getHeight()));
        row.setHeight(30);
        System.out.println("setHeight(30)後 = " + num(row.getHeight()));
        System.out.println("getHidden 既定  = " + row.getHidden()
                + " / getOutlineLevel = " + row.getOutlineLevel());
        ws.getRow(2).setHidden(true);                    // 2 行目を非表示
        System.out.println("2行目 hidden    = " + ws.getRow(2).getHidden());

        // ---------- Col ----------
        ColWrapper col = ws.getCol(1);                   // 1 列目（A 列）。列番号は数値で渡す
        System.out.println("既定の列幅      = columnWidth: " + num(col.getColumnWidth())
                + " / width: " + num(col.getWidth()));
        col.setColumnWidth(20);
        System.out.println("setColumnWidth(20) -&gt; columnWidth=" + num(col.getColumnWidth())
                + " / width=" + num(col.getWidth()));
        ws.getCol(2).setHidden(true);                    // B 列を非表示
        System.out.println("B列 hidden      = " + ws.getCol(2).getHidden());

        ws.getRange("C1:C1").setValue("自動調整のサンプル文字列");
        System.out.println("C列 bestFit 既定 = " + ws.getCol(3).getBestFit());
        ws.getCol(3).setBestFit(true);                   // 内容に合わせて列幅を自動調整
        System.out.println("setBestFit(true) -&gt; bestFit = " + ws.getCol(3).getBestFit()
                + " / columnWidth = " + num(ws.getCol(3).getColumnWidth()));

        // アウトライン（グループ化）は WorkSheet 側で設定する
        ws.setRowOutline(4, 6, 1);                       // 4〜6 行目をレベル 1 でグループ化
        ws.setColOutline(2, 3, 1);                       // B〜C 列をレベル 1 でグループ化
        System.out.println("4行目 outlineLevel = " + ws.getRow(4).getOutlineLevel()
                + " / B列 outlineLevel = " + ws.getCol(2).getOutlineLevel());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// row-col-props.php — 行の高さ・列の幅・非表示・自動調整 (phposbxl)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("row-col-props-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// PHP の bool は echo すると 1/空文字になるため JS 風の true/false に変換する
$boolstr = fn(bool $b): string =&gt; $b ? "true" : "false";

// ---------- Row ----------
$row = $ws-&gt;getRow(1);                        // 1 行目
echo "既定の高さ      = ", $row-&gt;getHeight(), "\n";
$row-&gt;setHeight(30);
echo "setHeight(30)後 = ", $row-&gt;getHeight(), "\n";
echo "getHidden 既定  = ", $boolstr($row-&gt;getHidden()), " / getOutlineLevel = ", $row-&gt;getOutlineLevel(), "\n";
$ws-&gt;getRow(2)-&gt;setHidden(true);              // 2 行目を非表示
echo "2行目 hidden    = ", $boolstr($ws-&gt;getRow(2)-&gt;getHidden()), "\n";

// ---------- Col ----------
$col = $ws-&gt;getCol(1);                        // 1 列目（A 列）。列番号は数値で渡す
echo "既定の列幅      = columnWidth: ", $col-&gt;getColumnWidth(), " / width: ", $col-&gt;getWidth(), "\n";
$col-&gt;setColumnWidth(20);
echo "setColumnWidth(20) -&gt; columnWidth=", $col-&gt;getColumnWidth(),
    " / width=", $col-&gt;getWidth(), "\n";
$ws-&gt;getCol(2)-&gt;setHidden(true);              // B 列を非表示
echo "B列 hidden      = ", $boolstr($ws-&gt;getCol(2)-&gt;getHidden()), "\n";

$ws-&gt;getRange("C1:C1")-&gt;setValue("自動調整のサンプル文字列");
echo "C列 bestFit 既定 = ", $boolstr($ws-&gt;getCol(3)-&gt;getBestFit()), "\n";
$ws-&gt;getCol(3)-&gt;setBestFit(true);             // 内容に合わせて列幅を自動調整
echo "setBestFit(true) -&gt; bestFit = ", $boolstr($ws-&gt;getCol(3)-&gt;getBestFit()),
    " / columnWidth = ", $ws-&gt;getCol(3)-&gt;getColumnWidth(), "\n";

// アウトライン（グループ化）は WorkSheet 側で設定する
$ws-&gt;setRowOutline(4, 6, 1);                  // 4〜6 行目をレベル 1 でグループ化
$ws-&gt;setColOutline(2, 3, 1);                  // B〜C 列をレベル 1 でグループ化
echo "4行目 outlineLevel = ", $ws-&gt;getRow(4)-&gt;getOutlineLevel(),
    " / B列 outlineLevel = ", $ws-&gt;getCol(2)-&gt;getOutlineLevel(), "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g1-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// row-col-props.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("row-col-props-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- Row ----------
const row = ws.getRow(1);                        // 1 行目
console.log("既定の高さ      =", row.getHeight());
row.setHeight(30);
console.log("setHeight(30)後 =", row.getHeight());
console.log("getHidden 既定  =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true);                    // 2 行目を非表示
console.log("2行目 hidden    =", ws.getRow(2).getHidden());

// ---------- Col ----------
const col = ws.getCol(1);                        // 1 列目（A 列）。列番号は数値で渡す
console.log("既定の列幅      = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -&gt; columnWidth=" + col.getColumnWidth(),
            "/ width=" + col.getWidth());
ws.getCol(2).setHidden(true);                    // B 列を非表示
console.log("B列 hidden      =", ws.getCol(2).getHidden());

ws.getRange("C1:C1").setValue("自動調整のサンプル文字列");
console.log("C列 bestFit 既定 =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true);                   // 内容に合わせて列幅を自動調整
console.log("setBestFit(true) -&gt; bestFit =", ws.getCol(3).getBestFit(),
            "/ columnWidth =", ws.getCol(3).getColumnWidth());

// アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1);                       // 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1);                       // B〜C 列をレベル 1 でグループ化
console.log("4行目 outlineLevel =", ws.getRow(4).getOutlineLevel(),
            "/ B列 outlineLevel =", ws.getCol(2).getOutlineLevel());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">既定の高さ      = 18.75
setHeight(30)後 = 30
getHidden 既定  = false / getOutlineLevel = 0
2行目 hidden    = true
既定の列幅      = columnWidth: 8.38 / width: 72
setColumnWidth(20) -&gt; columnWidth=20 / width=165
B列 hidden      = true
C列 bestFit 既定 = false
setBestFit(true) -&gt; bestFit = true / columnWidth = 25.38
4行目 outlineLevel = 1 / B列 outlineLevel = 1</code></pre>
</div>
<p class="mb-4 text-slate-700 leading-7"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBestFit(true)</code> による列幅は、セルの内容とブックの既定フォントから算出されるため 環境によって値が変わります（上記は 25.38）。</p>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>行番号・列番号はすべて 1 始まりの数値。</strong> 列は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"A"</code> ではなく <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1</code> で渡します。<br> 文字列から変換する場合は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.convertToColumnNumber("B")</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2</code> を使います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insert</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">delete</code> はデータをずらし、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clear</code> はずらしません。</strong> 行を空にしたいだけなら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRow()</code>（既存データの位置は保たれる）、 行そのものを無くしたいなら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteRow()</code> を使います。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copy</code> 系はコピー元を「シート名」で指定します。</strong> 自シート内コピーもシート名を渡します。<br> 第5引数 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyType</code> に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">enums.XlCopyContentType</code> を渡すとコピー対象を絞れます。<br> | 値 | 数値 | コピー対象 | |---|---|---| | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeAll</code> | 0 | 値、セル書式、関数セルのセル書式（既定） | | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeValueOnly</code> | 1 | 値のみ | | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeValueAndFormat</code> | 2 | 値とセル書式 | | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeFormatOnly</code> | 3 | セル書式のみ |</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col.getWidth()</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col.getColumnWidth()</code> は単位が違います。</strong> 実測で、既定状態の A 列は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWidth()</code> = 72、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColumnWidth()</code> = 8.38 でした。<br> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnWidth(20)</code> の後は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWidth()</code> = 165、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getColumnWidth()</code> = 20 です。<br> <strong>設定できるのは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColumnWidth()</code>（文字幅）だけ</strong>で、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getWidth()</code> は読み取り専用です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBestFit(true)</code> は設定時のセル内容に合わせて列幅を再計算します。</strong> 後に内容を書き換えても自動では追従しないので、データを入れ終えてから呼んでください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>アウトライン（グループ化）は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">WorkSheet</code> のメソッド。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setRowOutline(startRow, endRow, level)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setColOutline(startCol, endCol, level)</code> で設定し、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearRowOutline(level)</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">clearColOutline(level)</code> で解除します。<br> 結果は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Row.getOutlineLevel()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Col.getOutlineLevel()</code> で確認できます。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">insertCell</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">deleteCell</code> の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">direction</code> は操作に合わせて使い分けます。</strong> 挿入は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionRight</code>(1) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">InsertDirectionDown</code>(2)、 削除は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionLeft</code>(3) / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DeleteDirectionUp</code>(4) が本来の指定です。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copy</code> 系でコピーされないものがあります。</strong> コピー対象を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">CopyContentTypeAll</code> にしても、 <strong>図形（グラフを含む）・ピボットテーブル・queryTable 形式のテーブル・データテーブルは コピーされません。</strong></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell</code> の第3引数は「コピー先の左上セル」1 つだけ。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCell("Sheet1", "A3:B4", "F8")</code> は F8:G9 に展開されます（範囲の大きさはコピー元から決まる）。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyRow</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">copyCol</code> のコピー元はシート名、位置は行番号・列番号（数値）。</strong> セル範囲の文字列（A1 形式）は使いません。<br>混同しやすいので注意してください。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-04-value-types.php" class="text-primary-600 hover:underline">セル値の型と表示書式</a> / 次: <a href="general-06-bulk-input.php" class="text-primary-600 hover:underline">大量データの一括投入</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/nodejs/classes/nodeosbxl.Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/nodejs/enums/enums.XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/nodejs/enums/enums.XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/python/pyosbxl.Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/python/pyosbxl.Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/python/pyosbxl.enums.XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/python/pyosbxl.enums.XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/WorkSheetWrapper.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/java/com/osboffice/osbxl/RowWrapper.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/java/com/osboffice/osbxl/ColWrapper.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/java/com/osboffice/osbxl/enums/XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/php/classes/phposbxl-Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/php/classes/phposbxl-Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/php/classes/phposbxl-enums-XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.WorkSheet.html" class="text-primary-600 hover:underline">WorkSheet</a> / <a href="/api-docs/wasm/classes/osbxl.Row.html" class="text-primary-600 hover:underline">Row</a> / <a href="/api-docs/wasm/classes/osbxl.Col.html" class="text-primary-600 hover:underline">Col</a> / <a href="/api-docs/wasm/enums/enums.XlInsertDirection.html" class="text-primary-600 hover:underline">XlInsertDirection</a> / <a href="/api-docs/wasm/enums/enums.XlCopyContentType.html" class="text-primary-600 hover:underline">XlCopyContentType</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "行・列・セル範囲の操作 — 挿入・削除・クリア・コピー",
   "description": "行・列・セル範囲の操作 — 挿入・削除・クリア・コピー。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-05-rows-cols.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-05-rows-cols.php",
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
     "name": "行・列・セル範囲の操作 — 挿入・削除・クリア・コピー",
     "item": "https://osb-office.com/sample/general-05-rows-cols.php"
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
