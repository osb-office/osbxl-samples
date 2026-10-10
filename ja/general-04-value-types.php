<?php
$page_title = 'セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値 | osbxl サンプル';
$no_tagline = true;
$meta_description = 'セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。';
$meta_keywords = 'osbxl,サンプル,Excel,Excel生成ライブラリ,setvalue,setnumbervalue,setdatevalue,setdatestringvalue,setbooleanvalue,datetimeobject,numberformat,getvalue,rawvalue,forcestring';
include __DIR__ . '/../includes/header.php';
include __DIR__ . '/../includes/sidebar.php';
?>

<main class="main-content">
<h1 class="text-3xl font-bold mb-2 mt-2">セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値</h1>
<p class="text-sm text-text-light mb-6">level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ <a href="/en/sample/general-04-value-types.php" class="text-primary-600 hover:underline">English</a></p><h2 id="やりたいこと" class="text-2xl font-bold mt-10 mb-4">やりたいこと</h2>
<p class="mb-4 text-slate-700 leading-7">セルに入れる値の型（文字列 / 数値 / 日付 / 時刻 / 真偽値）ごとに正しい setter を使い、 表示書式を指定し、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code> の違いを押さえる。</p>
<h2 id="型の早見表" class="text-2xl font-bold mt-10 mb-4">型の早見表</h2>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">入れたい値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">メソッド</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">例</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">文字列</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue("りんご")</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">数値</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(num)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue(1234.5)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">日付・時刻</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(dateTimeObject)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue(d)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">日付・時刻（文字列から）</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue("2024-01-15")</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">真偽値</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(bool)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setBooleanValue(true)</code></td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">型を自動判定させたい</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue("2024/1/15")</code> → 日付になる</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top">数式</td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula(str)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code></td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">すべての setter は共通の第2・第3引数 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">(forceString?, numberFormat?)</code> を持ちます （<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue</code> も同様）。</p>
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
        <pre><code class="language-javascript">// value-types.js — 型ごとの setter と表示書式、生値と表示値の違い
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("value-types.xlsx");
const ws = wb.openWorkSheet("Sheet1");

const show = (a1) =&gt; {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`  ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])}` +
                ` 生値=${JSON.stringify(r.getValue(true)[a1])}` +
                ` 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
};

// --- 日付・時刻 ---
// DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);

const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);

const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);

console.log("[日付・時刻]");
ws.getRange("A1:A1").setDateValue(date);                      // 日付のみ
ws.getRange("A2:A2").setDateValue(datetime);                  // 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time);                      // 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15");        // 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");      // 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");  // 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);

// --- 数値・文字列 ---
console.log("[数値・文字列]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123");      // 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15");  // 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0');   // 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);

// --- 真偽値 ---
console.log("[真偽値]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE");     // 数式でも真偽値になる
["C1", "C2", "C3"].forEach(show);

// --- 日付のシリアル値 ---
console.log("[シリアル値]");
console.log("  2024/1/15 のシリアル値 =", app.getNumericValue(date));
console.log("  1904 年形式の場合      =", app.getNumericValue(date, true));
console.log("  wb.isDate1904()        =", wb.isDate1904());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-python" data-lang="python">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-python"># value-types.py — node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("value-types-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# JSON.stringify 相当で引用して 表示値/生値/書式 を 1 行で出す
def show(a1):
    r = ws.getRange(f"{a1}:{a1}")
    disp = json.dumps(r.getValue()[a1], ensure_ascii=False)
    raw = json.dumps(r.getValue(True)[a1], ensure_ascii=False)      # True で生値
    fmt = json.dumps(r.getNumberFormat()[a1], ensure_ascii=False)
    print(f"  {a1}: 表示値={disp} 生値={raw} 書式={fmt}")

def num(x):          # 整数値の float は int 表記にする（JS の出力と揃える）
    return int(x) if isinstance(x, float) and x.is_integer() else x

def boolstr(x):      # Python の True/False を JS 風の true/false で表示する
    return str(x).lower()

# --- 日付・時刻 ---
# DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 1, 15)

dt = pyosbxl.dto.DateTimeObject()   # ※ 変数名 datetime は標準モジュール名を避けて dt
dt.setYMD(2024, 1, 15)
dt.setHMS(13, 45, 30)

time = pyosbxl.dto.DateTimeObject()
time.setHMS(9, 5, 0)

print("[日付・時刻]")
ws.getRange("A1:A1").setDateValue(date)                       # 日付のみ
ws.getRange("A2:A2").setDateValue(dt)                         # 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time)                       # 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15")         # 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日")       # 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, False, "yyyy-mm-dd")   # 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, False, "ggge年m月d日") # 和暦+引用符なし日本語
for a in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"]:
    show(a)

# --- 数値・文字列 ---
print("[数値・文字列]")
ws.getRange("B1:B1").setNumberValue(1234567.891)
ws.getRange("B2:B2").setNumberValue(1234567.891, False, "#,##0.00")
ws.getRange("B3:B3").setNumberValue(0.25, False, "0.0%")
ws.getRange("B4:B4").setValue("00123")       # 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15")   # 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", True) # forceString=True で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, False, '"円"#,##0')      # 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, False, "#,##0;赤-#,##0") # 裸の色トークン（[赤] へ自動変換）
for b in ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"]:
    show(b)

# --- 真偽値 ---
print("[真偽値]")
ws.getRange("C1:C1").setBooleanValue(True)
ws.getRange("C2:C2").setBooleanValue(False)
ws.getRange("C3:C3").setFormula("TRUE")      # 数式でも真偽値になる（"=" は付けない）
for c in ["C1", "C2", "C3"]:
    show(c)

# --- 日付のシリアル値 ---
print("[シリアル値]")
print("  2024/1/15 のシリアル値 =", num(app.getNumericValue(date)))
print("  1904 年形式の場合      =", num(app.getNumericValue(date, True)))
print("  wb.isDate1904()        =", boolstr(wb.isDate1904()))

wb.save()
wb.close()</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-java" data-lang="java">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-java">// value-types.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.Map;

public class Main {
    static AppWrapper app;
    static WorkSheetWrapper ws;

    // JS の JSON.stringify 相当の引用（サンプル内に制御文字は無い前提）
    static String jsonq(String s) {
        return "\"" + s.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
    }

    // 整数値の double は int 表記にする（JS の出力と揃える）
    static String num(double d) {
        if (d == Math.rint(d) &amp;&amp; !Double.isInfinite(d)) return String.valueOf((long) d);
        return String.valueOf(d);
    }

    static void show(String a1) {
        RangeWrapper r = ws.getRange(a1 + ":" + a1);
        Map&lt;String, String&gt; disp = r.getValue();
        Map&lt;String, String&gt; raw = r.getValue(true);          // true で生値
        Map&lt;String, String&gt; fmt = r.getNumberFormat();
        System.out.println("  " + a1 + ": 表示値=" + jsonq(disp.get(a1))
                + " 生値=" + jsonq(raw.get(a1))
                + " 書式=" + jsonq(fmt.get(a1)));
    }

    public static void main(String[] args) {
        app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("value-types-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        // --- 日付・時刻 ---
        // DateTimeObjectWrapper は引数なしで生成し、setYMD() / setHMS() で設定します
        DateTimeObjectWrapper date = new DateTimeObjectWrapper();
        date.setYMD(2024, 1, 15);

        DateTimeObjectWrapper datetime = new DateTimeObjectWrapper();
        datetime.setYMD(2024, 1, 15);
        datetime.setHMS(13, 45, 30);

        DateTimeObjectWrapper time = new DateTimeObjectWrapper();
        time.setHMS(9, 5, 0);

        System.out.println("[日付・時刻]");
        ws.getRange("A1:A1").setDateValue(date);                       // 日付のみ
        ws.getRange("A2:A2").setDateValue(datetime);                   // 日付 + 時刻
        ws.getRange("A3:A3").setDateValue(time);                       // 時刻のみ
        ws.getRange("A4:A4").setDateStringValue("2024-01-15");         // 文字列から
        ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");       // 和暦も認識されます
        ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");   // 表示書式を同時に指定
        ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語
        for (String a : new String[]{"A1", "A2", "A3", "A4", "A5", "A6", "A7"}) show(a);

        // --- 数値・文字列 ---
        System.out.println("[数値・文字列]");
        ws.getRange("B1:B1").setNumberValue(1234567.891);
        ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
        ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
        ws.getRange("B4:B4").setValue("00123");       // 先頭 0 は文字列として保持される
        ws.getRange("B5:B5").setValue("2024/1/15");   // 汎用 setter は日付として自動認識
        ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
        ws.getRange("B7:B7").setNumberValue(1234.5, false, "\"円\"#,##0");    // 引用符リテラル
        ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン
        for (String b : new String[]{"B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"}) show(b);

        // --- 真偽値 ---
        System.out.println("[真偽値]");
        ws.getRange("C1:C1").setBooleanValue(true);
        ws.getRange("C2:C2").setBooleanValue(false);
        ws.getRange("C3:C3").setFormula("TRUE");      // 数式でも真偽値になる（"=" は付けない）
        for (String c : new String[]{"C1", "C2", "C3"}) show(c);

        // --- 日付のシリアル値 ---
        System.out.println("[シリアル値]");
        System.out.println("  2024/1/15 のシリアル値 = " + num(app.getNumericValue(date)));
        System.out.println("  1904 年形式の場合      = " + num(app.getNumericValue(date, true)));
        System.out.println("  wb.isDate1904()        = " + wb.isDate1904());

        wb.save();
        wb.close();
    }
}</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-php" data-lang="php">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-php">// value-types.php — node 版サンプルと同一検証 (phposbxl)。ボディのみ (&lt;?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app-&gt;createWorkBook("value-types-php.xlsx");
$ws = $wb-&gt;openWorkSheet("Sheet1");

// JS の JSON.stringify 相当の引用
function jsonq($s) {
    return json_encode($s, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
}

// 整数値の float は int 表記にする（JS の出力と揃える）
function num($x) {
    return $x == floor($x) ? (string)(int)$x : (string)$x;
}

// PHP の bool を JS 風の true/false で表示する
function boolstr($x) {
    return $x ? "true" : "false";
}

function show($a1) {
    global $ws;
    $r = $ws-&gt;getRange("{$a1}:{$a1}");
    echo "  {$a1}: 表示値=" . jsonq($r-&gt;getValue()[$a1])
       . " 生値=" . jsonq($r-&gt;getValue(true)[$a1])
       . " 書式=" . jsonq($r-&gt;getNumberFormat()[$a1]) . "\n";
}

// --- 日付・時刻 ---
// DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
$date = new \phposbxl\dto\DateTimeObject();
$date-&gt;setYMD(2024, 1, 15);

$datetime = new \phposbxl\dto\DateTimeObject();
$datetime-&gt;setYMD(2024, 1, 15);
$datetime-&gt;setHMS(13, 45, 30);

$time = new \phposbxl\dto\DateTimeObject();
$time-&gt;setHMS(9, 5, 0);

echo "[日付・時刻]\n";
$ws-&gt;getRange("A1:A1")-&gt;setDateValue($date);                       // 日付のみ
$ws-&gt;getRange("A2:A2")-&gt;setDateValue($datetime);                   // 日付 + 時刻
$ws-&gt;getRange("A3:A3")-&gt;setDateValue($time);                       // 時刻のみ
$ws-&gt;getRange("A4:A4")-&gt;setDateStringValue("2024-01-15");          // 文字列から
$ws-&gt;getRange("A5:A5")-&gt;setDateStringValue("令和6年1月15日");        // 和暦も認識されます
$ws-&gt;getRange("A6:A6")-&gt;setDateValue($date, false, "yyyy-mm-dd");   // 表示書式を同時に指定
$ws-&gt;getRange("A7:A7")-&gt;setDateValue($date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
foreach (["A1", "A2", "A3", "A4", "A5", "A6", "A7"] as $a) show($a);

// --- 数値・文字列 ---
echo "[数値・文字列]\n";
$ws-&gt;getRange("B1:B1")-&gt;setNumberValue(1234567.891);
$ws-&gt;getRange("B2:B2")-&gt;setNumberValue(1234567.891, false, "#,##0.00");
$ws-&gt;getRange("B3:B3")-&gt;setNumberValue(0.25, false, "0.0%");
$ws-&gt;getRange("B4:B4")-&gt;setValue("00123");       // 先頭 0 は文字列として保持される
$ws-&gt;getRange("B5:B5")-&gt;setValue("2024/1/15");   // 汎用 setter は日付として自動認識
$ws-&gt;getRange("B6:B6")-&gt;setValue("12345", true); // forceString=true で強制的に文字列扱い
$ws-&gt;getRange("B7:B7")-&gt;setNumberValue(1234.5, false, '"円"#,##0');      // 引用符リテラル
$ws-&gt;getRange("B8:B8")-&gt;setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
foreach (["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"] as $bx) show($bx);

// --- 真偽値 ---
echo "[真偽値]\n";
$ws-&gt;getRange("C1:C1")-&gt;setBooleanValue(true);
$ws-&gt;getRange("C2:C2")-&gt;setBooleanValue(false);
$ws-&gt;getRange("C3:C3")-&gt;setFormula("TRUE");      // 数式でも真偽値になる
foreach (["C1", "C2", "C3"] as $c) show($c);

// --- 日付のシリアル値 ---
echo "[シリアル値]\n";
echo "  2024/1/15 のシリアル値 = " . num($app-&gt;getNumericValue($date)) . "\n";
echo "  1904 年形式の場合      = " . num($app-&gt;getNumericValue($date, true)) . "\n";
echo "  wb.isDate1904()        = " . boolstr($wb-&gt;isDate1904()) . "\n";

$wb-&gt;save();
$wb-&gt;close();</code></pre>
      </div>
    </div>
    <div class="tab-content hidden" id="tab-g0-wasm" data-lang="wasm">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">// value-types.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("value-types-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

const show = (a1) =&gt; {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`  ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])}` +
                ` 生値=${JSON.stringify(r.getValue(true)[a1])}` +
                ` 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
};

// --- 日付・時刻 ---
// DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);

const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);

const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);

console.log("[日付・時刻]");
ws.getRange("A1:A1").setDateValue(date);                      // 日付のみ
ws.getRange("A2:A2").setDateValue(datetime);                  // 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time);                      // 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15");        // 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");      // 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");  // 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);

// --- 数値・文字列 ---
console.log("[数値・文字列]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123");      // 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15");  // 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0');   // 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);

// --- 真偽値 ---
console.log("[真偽値]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE");     // 数式でも真偽値になる
["C1", "C2", "C3"].forEach(show);

// --- 日付のシリアル値 ---
console.log("[シリアル値]");
console.log("  2024/1/15 のシリアル値 =", app.getNumericValue(date));
console.log("  1904 年形式の場合      =", app.getNumericValue(date, true));
console.log("  wb.isDate1904()        =", wb.isDate1904());

wb.save();
wb.close();</code></pre>
      </div>
    </div>
</div>
<div class="bg-slate-800 rounded-lg p-4 overflow-x-auto mb-4">
  <pre><code class="language-plaintext">[日付・時刻]
  A1: 表示値="2024/1/15" 生値="45306" 書式="yyyy/m/d;@"
  A2: 表示値="2024/1/15 13:45:30" 生値="45306.573263888888" 書式="yyyy/m/d h:mm:ss;@"
  A3: 表示値="9:05:00" 生値="0.37847222222222221" 書式="h:mm:ss;@"
  A4: 表示値="2024/1/15" 生値="45306" 書式="yyyy/m/d;@"
  A5: 表示値="2024/1/15" 生値="45306" 書式="yyyy/m/d;@"
  A6: 表示値="2024-01-15" 生値="45306" 書式="yyyy-mm-dd"
  A7: 表示値="令和6年1月15日" 生値="45306" 書式="ggge年m月d日"
[数値・文字列]
  B1: 表示値="1234567.891" 生値="1234567.8910000001" 書式="General"
  B2: 表示値="1,234,567.89" 生値="1234567.8910000001" 書式="#,##0.00"
  B3: 表示値="25.0%" 生値="0.25" 書式="0.0%"
  B4: 表示値="00123" 生値="00123" 書式="General"
  B5: 表示値="2024/01/15" 生値="45306" 書式="yyyy/mm/dd"
  B6: 表示値="12345" 生値="12345" 書式="General"
  B7: 表示値="円1,235" 生値="1234.5" 書式="\"円\"#,##0"
  B8: 表示値="-1,235" 生値="-1234.5" 書式="#,##0;赤-#,##0"
[真偽値]
  C1: 表示値="TRUE" 生値="TRUE" 書式="General"
  C2: 表示値="FALSE" 生値="FALSE" 書式="General"
  C3: 表示値="TRUE" 生値="TRUE" 書式="General"
[シリアル値]
  2024/1/15 のシリアル値 = 45306
  1904 年形式の場合      = 43844
  wb.isDate1904()        = false</code></pre>
</div>
<h2 id="解説" class="text-2xl font-bold mt-10 mb-4">解説</h2>
<h3 id="getvalue-と-getvalue-true" class="text-xl font-bold mt-8 mb-3"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code> と <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code></h3>
<div class="overflow-x-auto mb-4"><table class="w-full text-sm border-collapse">
  <thead><tr><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">呼び出し</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">返る値</th><th class="border border-surface-200 px-3 py-2 bg-surface-100 text-left whitespace-nowrap">用途</th></tr></thead>
  <tbody>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code></td><td class="border border-surface-200 px-3 py-2 align-top"><strong>表示値</strong>（表示書式を適用した文字列）</td><td class="border border-surface-200 px-3 py-2 align-top">帳票出力、画面表示、CSV 書き出し</td></tr>
    <tr><td class="border border-surface-200 px-3 py-2 align-top"><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue(true)</code></td><td class="border border-surface-200 px-3 py-2 align-top"><strong>生値</strong>（内部保持値。日付はシリアル値、数式は計算結果）</td><td class="border border-surface-200 px-3 py-2 align-top">計算、比較、数値処理</td></tr>
  </tbody></table></div>
<p class="mb-4 text-slate-700 leading-7">どちらも戻り値は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">{ "A1": "...", ... }</code> のマップで、<strong>値はすべて文字列</strong>です。 数値として扱うには <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">Number(...)</code> で変換します。</p>
<div class="tab-group mb-6">
  <div class="flex gap-2 mb-4 border-b border-surface-200 overflow-x-auto">
    <button class="tab-btn px-4 py-2 text-sm font-bold border-b-2 active border-primary-600 text-primary-600" data-tab="node">Node.js</button>
  </div>
    <div class="tab-content" id="tab-g1-node" data-lang="node">
      <div class="bg-slate-800 rounded-lg p-4 overflow-x-auto space-y-3">
        <pre><code class="language-javascript">const raw = Number(ws.getRange("A1:A1").getValue(true)["A1"]);   // 45306
const disp = ws.getRange("A1:A1").getValue()["A1"];              // "2024/1/15"</code></pre>
      </div>
    </div>
</div>
<h3 id="日付" class="text-xl font-bold mt-8 mb-3">日付</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">DateTimeObject</code> は引数なしで生成し、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setYMD()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHMS()</code> で設定します。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setYear()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMonth()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDay()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setHour()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setMinute()</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setSecond()</code> で 個別に設定することもできます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">numberFormat</code> を指定しないと、内容に応じて既定書式が自動設定されます。</strong><ul class="mt-2 space-y-1 pl-5 list-disc text-slate-700"><li>日付のみ → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d;@</code></li><li>時刻のみ → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">h:mm:ss;@</code></li><li>日付 + 時刻 → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d h:mm:ss;@</code></li></ul></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue()</code> は Excel と同じ日付文字列を認識します。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"2024-01-15"</code>、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"2024/1/15 13:45:30"</code>、和暦の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"令和6年1月15日"</code> も日付として入ります。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">app.getNumericValue(dateTimeObject)</code> でシリアル値を取得できます。</strong> 第2引数 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">is1904</code> に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> を渡すと 1904 年形式のシリアル値になります。<br> ブックがどちらの形式かは <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">wb.isDate1904()</code> で確認できます。</li>
</ul>
<h3 id="数値・文字列" class="text-xl font-bold mt-8 mb-3">数値・文字列</h3>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setValue()</code> は型を自動判定します。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"123"</code> → 数値、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"2024/1/15"</code> → 日付、 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"00123"</code> → 文字列（先頭 0 が保持される）、<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"12A"</code> → 文字列。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>意図した型で確実に入れたければ専用 setter を使います。</strong> 数値なら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberValue()</code>、文字列として固定したいなら <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">forceString</code> を <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> にします。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">numberFormat</code> は setter の第3引数で同時に指定できます。</strong> 後から <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberFormat()</code> を呼ぶのと同じ結果になります。</li>
</ul>
<h2 id="注意点" class="text-2xl font-bold mt-10 mb-4">注意点</h2>
<ul class="space-y-2 mb-4">
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>表示書式は「値を入れてから」設定する。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setNumberFormat()</code> の後に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateValue()</code> を呼ぶと、 日付 setter が既定書式で<strong>上書き</strong>します（実測）。<div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">// ✗ 書式が消える
ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");
ws.getRange("A1:A1").setDateValue(date);      // → 書式が "yyyy/m/d;@" に上書きされる

// ✓ 書式が残る
ws.getRange("A1:A1").setDateValue(date);
ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");

// ✓ あるいは setter の第3引数で同時に指定する
ws.getRange("A1:A1").setDateValue(date, false, "yyyy-mm-dd");</code></pre></div></li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setDateStringValue()</code> は認識できない文字列で例外になります。</strong><div class="bg-slate-800 rounded-lg p-3 my-6 overflow-x-auto"><pre><code class="language-javascript">ws.getRange("A1:A1").setDateStringValue("これは日付ではない");
// → 例外: can't recognize input value as DateTime</code></pre></div>例外にしたくない場合は第2引数 <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">forceString</code> に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">true</code> を渡します。<br> 認識に失敗したときは文字列としてそのまま入力されます。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>表示書式は日本語・通貨記号・色指定も使えます</strong>（2026-09-30 検証）。<br> 引用符リテラル（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"円"#,##0</code>）も引用符なしの日本語（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy年m月d日</code>）もそのまま表示され、 和暦（<code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ggge年m月d日</code> → <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">令和6年1月15日</code>）や <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">¥#,##0</code> も動作します。<br> 色指定は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[Red]</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[赤]</code> / 裸の <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">赤</code> の3表記とも受け付け、すべて <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">[Red]</code> として保存されます。<br> 動作確認済みの書式の例: | 種別 | 書式 | 表示例 | |---|---|---| | 日付 | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/m/d</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy/mm/dd</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy-mm-dd</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yy/m/d</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">m/d</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">d-mmm-yyyy</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2024/1/15</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2024-01-15</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">15-Jan-2024</code> | | 日付（日本語） | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">yyyy年m月d日</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">ggge年m月d日</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">2024年1月15日</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">令和6年1月15日</code> | | 数値 | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">General</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0.00</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0.000</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0%</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">0.0%</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1,235</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1,234.57</code>, <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">25.0%</code> | | 数値（日本語・通貨） | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"円"#,##0</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">¥#,##0</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">円1,235</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">¥1,235</code> | | 数値（負数セクション） | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0;[Red]-#,##0</code> / <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">#,##0;赤-#,##0</code> | <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">-1,235</code>（赤字） |</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong>数値の生値は double の表現誤差を含みます。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">1234567.891</code> の生値は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"1234567.8910000001"</code> になります。<br>比較や表示には <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getValue()</code>（表示値）を使うか、 丸めてから扱ってください。</li>
  <li class="text-slate-700 leading-relaxed border-l-4 border-surface-200 pl-4"><strong><code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code> の戻り値に <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">=</code> は付きません。</strong> <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">setFormula("SUM(A1:A3)")</code> に対して <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">getFormula()</code> は <code class="bg-surface-100 px-1 py-0.5 rounded text-[0.85em]">"SUM(A1:A3)"</code> を返します。</li>
</ul>
<h2 id="関連" class="text-2xl font-bold mt-10 mb-4">関連</h2>
<ul class="list-disc pl-6 mb-4 space-y-1 text-slate-700">
  <li><a href="index.php" class="text-primary-600 hover:underline">目次</a></li>
  <li>前: <a href="general-03-sheets.php" class="text-primary-600 hover:underline">ブックとワークシートの操作</a> / 次: <a href="general-05-rows-cols.php" class="text-primary-600 hover:underline">行・列・セル範囲の操作</a></li>
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
    <div class="tab-content" id="tab-api0-node" data-lang="node"><p class="text-sm leading-7"><a href="/api-docs/nodejs/classes/nodeosbxl.Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/nodejs/classes/dto.DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-python" data-lang="python"><p class="text-sm leading-7"><a href="/api-docs/python/pyosbxl.Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/python/pyosbxl.dto.DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-java" data-lang="java"><p class="text-sm leading-7"><a href="/api-docs/java/com/osboffice/osbxl/RangeWrapper.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/java/com/osboffice/osbxl/dto/DateTimeObjectWrapper.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-php" data-lang="php"><p class="text-sm leading-7"><a href="/api-docs/php/classes/phposbxl-Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/php/classes/phposbxl-dto-DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
    <div class="tab-content hidden" id="tab-api0-wasm" data-lang="wasm"><p class="text-sm leading-7"><a href="/api-docs/wasm/classes/osbxl.Range.html" class="text-primary-600 hover:underline">Range</a> / <a href="/api-docs/wasm/classes/dto.DateTimeObject.html" class="text-primary-600 hover:underline">DateTimeObject</a></p></div>
  </div>
</div></li>
</ul>
<script type="application/ld+json">
{
 "@context": "https://schema.org",
 "@graph": [
  {
   "@type": "TechArticle",
   "headline": "セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値",
   "description": "セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値。実行検証済み osbxl サンプル（verified 2026-10-04）。Node.js/Python/Java/PHP/WebAssembly タブ対応。",
   "inLanguage": "ja",
   "dateModified": "2026-10-04",
   "url": "https://osb-office.com/sample/general-04-value-types.php",
   "mainEntityOfPage": "https://osb-office.com/sample/general-04-value-types.php",
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
     "name": "セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値",
     "item": "https://osb-office.com/sample/general-04-value-types.php"
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
