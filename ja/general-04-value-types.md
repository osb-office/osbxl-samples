# セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値

level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-04-value-types.md)

## やりたいこと

セルに入れる値の型（文字列 / 数値 / 日付 / 時刻 / 真偽値）ごとに正しい setter を使い、 表示書式を指定し、`getValue()` と `getValue(true)` の違いを押さえる。

## 型の早見表

| 入れたい値 | メソッド | 例 |
| --- | --- | --- |
| 文字列 | `setValue(str)` | `setValue("りんご")` |
| 数値 | `setNumberValue(num)` | `setNumberValue(1234.5)` |
| 日付・時刻 | `setDateValue(dateTimeObject)` | `setDateValue(d)` |
| 日付・時刻（文字列から） | `setDateStringValue(str)` | `setDateStringValue("2024-01-15")` |
| 真偽値 | `setBooleanValue(bool)` | `setBooleanValue(true)` |
| 型を自動判定させたい | `setValue(str)` | `setValue("2024/1/15")` → 日付になる |
| 数式 | `setFormula(str)` | `setFormula("SUM(A1:A3)")` |

すべての setter は共通の第2・第3引数 `(forceString?, numberFormat?)` を持ちます （`setDateStringValue` / `setValue` も同様）。

## コード

実行結果:

#### [Node.js]

```javascript
// value-types.js — 型ごとの setter と表示書式、生値と表示値の違い
const { nodeosbxl, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("value-types.xlsx");
const ws = wb.openWorkSheet("Sheet1");
const show = (a1) => {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])}` +
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
ws.getRange("A1:A1").setDateValue(date); // 日付のみ
ws.getRange("A2:A2").setDateValue(datetime); // 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time); // 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15"); // 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日"); // 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd"); // 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);
// --- 数値・文字列 ---
console.log("[数値・文字列]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123"); // 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15"); // 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0'); // 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);
// --- 真偽値 ---
console.log("[真偽値]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE"); // 数式でも真偽値になる
["C1", "C2", "C3"].forEach(show);
// --- 日付のシリアル値 ---
console.log("[シリアル値]");
console.log(" 2024/1/15 のシリアル値 =", app.getNumericValue(date));
console.log(" 1904 年形式の場合 =", app.getNumericValue(date, true));
console.log(" wb.isDate1904() =", wb.isDate1904());
wb.save();
wb.close();
```

#### [Python]

```python
# value-types.py — node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("value-types-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# JSON.stringify 相当で引用して 表示値/生値/書式 を 1 行で出す
def show(a1):
r = ws.getRange(f"{a1}:{a1}")
disp = json.dumps(r.getValue()[a1], ensure\_ascii=False)
raw = json.dumps(r.getValue(True)[a1], ensure\_ascii=False) # True で生値
fmt = json.dumps(r.getNumberFormat()[a1], ensure\_ascii=False)
print(f" {a1}: 表示値={disp} 生値={raw} 書式={fmt}")
def num(x): # 整数値の float は int 表記にする（JS の出力と揃える）
return int(x) if isinstance(x, float) and x.is\_integer() else x
def boolstr(x): # Python の True/False を JS 風の true/false で表示する
return str(x).lower()
# --- 日付・時刻 ---
# DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 1, 15)
dt = pyosbxl.dto.DateTimeObject() # ※ 変数名 datetime は標準モジュール名を避けて dt
dt.setYMD(2024, 1, 15)
dt.setHMS(13, 45, 30)
time = pyosbxl.dto.DateTimeObject()
time.setHMS(9, 5, 0)
print("[日付・時刻]")
ws.getRange("A1:A1").setDateValue(date) # 日付のみ
ws.getRange("A2:A2").setDateValue(dt) # 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time) # 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15") # 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日") # 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, False, "yyyy-mm-dd") # 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, False, "ggge年m月d日") # 和暦+引用符なし日本語
for a in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"]:
show(a)
# --- 数値・文字列 ---
print("[数値・文字列]")
ws.getRange("B1:B1").setNumberValue(1234567.891)
ws.getRange("B2:B2").setNumberValue(1234567.891, False, "#,##0.00")
ws.getRange("B3:B3").setNumberValue(0.25, False, "0.0%")
ws.getRange("B4:B4").setValue("00123") # 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15") # 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", True) # forceString=True で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, False, '"円"#,##0') # 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, False, "#,##0;赤-#,##0") # 裸の色トークン（[赤] へ自動変換）
for b in ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"]:
show(b)
# --- 真偽値 ---
print("[真偽値]")
ws.getRange("C1:C1").setBooleanValue(True)
ws.getRange("C2:C2").setBooleanValue(False)
ws.getRange("C3:C3").setFormula("TRUE") # 数式でも真偽値になる（"=" は付けない）
for c in ["C1", "C2", "C3"]:
show(c)
# --- 日付のシリアル値 ---
print("[シリアル値]")
print(" 2024/1/15 のシリアル値 =", num(app.getNumericValue(date)))
print(" 1904 年形式の場合 =", num(app.getNumericValue(date, True)))
print(" wb.isDate1904() =", boolstr(wb.isDate1904()))
wb.save()
wb.close()
```

#### [Java]

```java
// value-types.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
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
if (d == Math.rint(d) && !Double.isInfinite(d)) return String.valueOf((long) d);
return String.valueOf(d);
}
static void show(String a1) {
RangeWrapper r = ws.getRange(a1 + ":" + a1);
Map<String, String> disp = r.getValue();
Map<String, String> raw = r.getValue(true); // true で生値
Map<String, String> fmt = r.getNumberFormat();
System.out.println(" " + a1 + ": 表示値=" + jsonq(disp.get(a1))
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
ws.getRange("A1:A1").setDateValue(date); // 日付のみ
ws.getRange("A2:A2").setDateValue(datetime); // 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time); // 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15"); // 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日"); // 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd"); // 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語
for (String a : new String[]{"A1", "A2", "A3", "A4", "A5", "A6", "A7"}) show(a);
// --- 数値・文字列 ---
System.out.println("[数値・文字列]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123"); // 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15"); // 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, false, "\"円\"#,##0"); // 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン
for (String b : new String[]{"B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"}) show(b);
// --- 真偽値 ---
System.out.println("[真偽値]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE"); // 数式でも真偽値になる（"=" は付けない）
for (String c : new String[]{"C1", "C2", "C3"}) show(c);
// --- 日付のシリアル値 ---
System.out.println("[シリアル値]");
System.out.println(" 2024/1/15 のシリアル値 = " + num(app.getNumericValue(date)));
System.out.println(" 1904 年形式の場合 = " + num(app.getNumericValue(date, true)));
System.out.println(" wb.isDate1904() = " + wb.isDate1904());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// value-types.md — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("value-types-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// JS の JSON.stringify 相当の引用
function jsonq($s) {
return json\_encode($s, JSON\_UNESCAPED\_UNICODE | JSON\_UNESCAPED\_SLASHES);
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
$r = $ws->getRange("{$a1}:{$a1}");
echo " {$a1}: 表示値=" . jsonq($r->getValue()[$a1])
. " 生値=" . jsonq($r->getValue(true)[$a1])
. " 書式=" . jsonq($r->getNumberFormat()[$a1]) . "\n";
}
// --- 日付・時刻 ---
// DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
$date = new \phposbxl\dto\DateTimeObject();
$date->setYMD(2024, 1, 15);
$datetime = new \phposbxl\dto\DateTimeObject();
$datetime->setYMD(2024, 1, 15);
$datetime->setHMS(13, 45, 30);
$time = new \phposbxl\dto\DateTimeObject();
$time->setHMS(9, 5, 0);
echo "[日付・時刻]\n";
$ws->getRange("A1:A1")->setDateValue($date); // 日付のみ
$ws->getRange("A2:A2")->setDateValue($datetime); // 日付 + 時刻
$ws->getRange("A3:A3")->setDateValue($time); // 時刻のみ
$ws->getRange("A4:A4")->setDateStringValue("2024-01-15"); // 文字列から
$ws->getRange("A5:A5")->setDateStringValue("令和6年1月15日"); // 和暦も認識されます
$ws->getRange("A6:A6")->setDateValue($date, false, "yyyy-mm-dd"); // 表示書式を同時に指定
$ws->getRange("A7:A7")->setDateValue($date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
foreach (["A1", "A2", "A3", "A4", "A5", "A6", "A7"] as $a) show($a);
// --- 数値・文字列 ---
echo "[数値・文字列]\n";
$ws->getRange("B1:B1")->setNumberValue(1234567.891);
$ws->getRange("B2:B2")->setNumberValue(1234567.891, false, "#,##0.00");
$ws->getRange("B3:B3")->setNumberValue(0.25, false, "0.0%");
$ws->getRange("B4:B4")->setValue("00123"); // 先頭 0 は文字列として保持される
$ws->getRange("B5:B5")->setValue("2024/1/15"); // 汎用 setter は日付として自動認識
$ws->getRange("B6:B6")->setValue("12345", true); // forceString=true で強制的に文字列扱い
$ws->getRange("B7:B7")->setNumberValue(1234.5, false, '"円"#,##0'); // 引用符リテラル
$ws->getRange("B8:B8")->setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
foreach (["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"] as $bx) show($bx);
// --- 真偽値 ---
echo "[真偽値]\n";
$ws->getRange("C1:C1")->setBooleanValue(true);
$ws->getRange("C2:C2")->setBooleanValue(false);
$ws->getRange("C3:C3")->setFormula("TRUE"); // 数式でも真偽値になる
foreach (["C1", "C2", "C3"] as $c) show($c);
// --- 日付のシリアル値 ---
echo "[シリアル値]\n";
echo " 2024/1/15 のシリアル値 = " . num($app->getNumericValue($date)) . "\n";
echo " 1904 年形式の場合 = " . num($app->getNumericValue($date, true)) . "\n";
echo " wb.isDate1904() = " . boolstr($wb->isDate1904()) . "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// value-types.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("value-types-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
const show = (a1) => {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])}` +
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
ws.getRange("A1:A1").setDateValue(date); // 日付のみ
ws.getRange("A2:A2").setDateValue(datetime); // 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time); // 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15"); // 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日"); // 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd"); // 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);
// --- 数値・文字列 ---
console.log("[数値・文字列]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123"); // 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15"); // 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0'); // 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);
// --- 真偽値 ---
console.log("[真偽値]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE"); // 数式でも真偽値になる
["C1", "C2", "C3"].forEach(show);
// --- 日付のシリアル値 ---
console.log("[シリアル値]");
console.log(" 2024/1/15 のシリアル値 =", app.getNumericValue(date));
console.log(" 1904 年形式の場合 =", app.getNumericValue(date, true));
console.log(" wb.isDate1904() =", wb.isDate1904());
wb.save();
wb.close();
```

```plaintext
[日付・時刻]
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
1904 年形式の場合 = 43844
wb.isDate1904() = false
```

## 解説

### `getValue()` と `getValue(true)`

| 呼び出し | 返る値 | 用途 |
| --- | --- | --- |
| `getValue()` | **表示値**（表示書式を適用した文字列） | 帳票出力、画面表示、CSV 書き出し |
| `getValue(true)` | **生値**（内部保持値。日付はシリアル値、数式は計算結果） | 計算、比較、数値処理 |

どちらも戻り値は `{ "A1": "...", ... }` のマップで、**値はすべて文字列**です。 数値として扱うには `Number(...)` で変換します。

#### [Node.js]

```javascript
const raw = Number(ws.getRange("A1:A1").getValue(true)["A1"]); // 45306
const disp = ws.getRange("A1:A1").getValue()["A1"]; // "2024/1/15"
```

### 日付

* **`DateTimeObject` は引数なしで生成し、`setYMD()` / `setHMS()` で設定します。** `setYear()` / `setMonth()` / `setDay()` / `setHour()` / `setMinute()` / `setSecond()` で 個別に設定することもできます。
* **`numberFormat` を指定しないと、内容に応じて既定書式が自動設定されます。**
  + 日付のみ → `yyyy/m/d;@`
  + 時刻のみ → `h:mm:ss;@`
  + 日付 + 時刻 → `yyyy/m/d h:mm:ss;@`
* **`setDateStringValue()` は Excel と同じ日付文字列を認識します。** `"2024-01-15"`、`"2024/1/15 13:45:30"`、和暦の `"令和6年1月15日"` も日付として入ります。
* **`app.getNumericValue(dateTimeObject)` でシリアル値を取得できます。** 第2引数 `is1904` に `true` を渡すと 1904 年形式のシリアル値になります。  
   ブックがどちらの形式かは `wb.isDate1904()` で確認できます。

### 数値・文字列

* **`setValue()` は型を自動判定します。** `"123"` → 数値、`"2024/1/15"` → 日付、 `"00123"` → 文字列（先頭 0 が保持される）、`"12A"` → 文字列。
* **意図した型で確実に入れたければ専用 setter を使います。** 数値なら `setNumberValue()`、文字列として固定したいなら `forceString` を `true` にします。
* **`numberFormat` は setter の第3引数で同時に指定できます。** 後から `setNumberFormat()` を呼ぶのと同じ結果になります。

## 注意点

* **表示書式は「値を入れてから」設定する。** `setNumberFormat()` の後に `setDateValue()` を呼ぶと、 日付 setter が既定書式で**上書き**します（実測）。

  ```javascript
  // ✗ 書式が消える
  ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");
  ws.getRange("A1:A1").setDateValue(date); // → 書式が "yyyy/m/d;@" に上書きされる
  // ✓ 書式が残る
  ws.getRange("A1:A1").setDateValue(date);
  ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");
  // ✓ あるいは setter の第3引数で同時に指定する
  ws.getRange("A1:A1").setDateValue(date, false, "yyyy-mm-dd");
  ```
* **`setDateStringValue()` は認識できない文字列で例外になります。**

  ```javascript
  ws.getRange("A1:A1").setDateStringValue("これは日付ではない");
  // → 例外: can't recognize input value as DateTime
  ```

  例外にしたくない場合は第2引数 `forceString` に `true` を渡します。  
   認識に失敗したときは文字列としてそのまま入力されます。
* **表示書式は日本語・通貨記号・色指定も使えます**（2026-09-30 検証）。  
   引用符リテラル（`"円"#,##0`）も引用符なしの日本語（`yyyy年m月d日`）もそのまま表示され、 和暦（`ggge年m月d日` → `令和6年1月15日`）や `¥#,##0` も動作します。  
   色指定は `[Red]` / `[赤]` / 裸の `赤` の3表記とも受け付け、すべて `[Red]` として保存されます。  
   動作確認済みの書式の例: | 種別 | 書式 | 表示例 | |---|---|---| | 日付 | `yyyy/m/d` / `yyyy/mm/dd` / `yyyy-mm-dd` / `yy/m/d` / `m/d` / `d-mmm-yyyy` | `2024/1/15`, `2024-01-15`, `15-Jan-2024` | | 日付（日本語） | `yyyy年m月d日` / `ggge年m月d日` | `2024年1月15日` / `令和6年1月15日` | | 数値 | `General` / `#,##0` / `#,##0.00` / `0.000` / `0%` / `0.0%` | `1,235`, `1,234.57`, `25.0%` | | 数値（日本語・通貨） | `"円"#,##0` / `¥#,##0` | `円1,235` / `¥1,235` | | 数値（負数セクション） | `#,##0;[Red]-#,##0` / `#,##0;赤-#,##0` | `-1,235`（赤字） |
* **数値の生値は double の表現誤差を含みます。** `1234567.891` の生値は `"1234567.8910000001"` になります。  
  比較や表示には `getValue()`（表示値）を使うか、 丸めてから扱ってください。
* **`getFormula()` の戻り値に `=` は付きません。** `setFormula("SUM(A1:A3)")` に対して `getFormula()` は `"SUM(A1:A3)"` を返します。

## 関連

* [目次](index.md)
* 前: [ブックとワークシートの操作](general-03-sheets.md) / 次: [行・列・セル範囲の操作](general-05-rows-cols.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
