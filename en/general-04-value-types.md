# Value Types and Number Formats — text, numbers, dates

level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](/ja/general-04-value-types.md)

## What you want to do

Use the correct setter for each type of value you put into a cell (string / number / date / time / boolean), specify number formats, and understand the difference between `getValue()` and `getValue(true)`.

## Type quick reference

| Value to store | Method | Example |
| --- | --- | --- |
| String | `setValue(str)` | `setValue("Apple")` |
| Number | `setNumberValue(num)` | `setNumberValue(1234.5)` |
| Date / time | `setDateValue(dateTimeObject)` | `setDateValue(d)` |
| Date / time (from a string) | `setDateStringValue(str)` | `setDateStringValue("2024-01-15")` |
| Boolean | `setBooleanValue(bool)` | `setBooleanValue(true)` |
| Let the type be detected automatically | `setValue(str)` | `setValue("2024/1/15")` → becomes a date |
| Formula | `setFormula(str)` | `setFormula("SUM(A1:A3)")` |

All setters share the same second and third arguments `(forceString?, numberFormat?)` (`setDateStringValue` / `setValue` included).

## Code

Output:

#### [Node.js]

```javascript
// value-types.js — per-type setters and number formats, raw vs. display values
const { nodeosbxl, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("value-types.xlsx");
const ws = wb.openWorkSheet("Sheet1");
const show = (a1) => {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: disp=${JSON.stringify(r.getValue()[a1])}` +
` raw=${JSON.stringify(r.getValue(true)[a1])}` +
` fmt=${JSON.stringify(r.getNumberFormat()[a1])}`);
};
// --- dates & times ---
// Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);
const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);
const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);
console.log("[Dates & times]");
ws.getRange("A1:A1").setDateValue(date); // date only
ws.getRange("A2:A2").setDateValue(datetime); // date + time
ws.getRange("A3:A3").setDateValue(time); // time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15"); // from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日"); // Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd"); // number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals also work
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);
// --- numbers & strings ---
console.log("[Numbers & strings]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123"); // leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15"); // the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", true); // forceString=true forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0'); // quoted literal ("yen")
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red) → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);
// --- booleans ---
console.log("[Booleans]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE"); // a formula can produce a boolean too
["C1", "C2", "C3"].forEach(show);
// --- date serial values ---
console.log("[Serial values]");
console.log(" serial value of 2024/1/15 =", app.getNumericValue(date));
console.log(" with the 1904 date system =", app.getNumericValue(date, true));
console.log(" wb.isDate1904() =", wb.isDate1904());
wb.save();
wb.close();
```

#### [Python]

```python
# value-types.py — same verification as the Node.js sample (pyosbxl)
import json
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("value-types-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# Print display value / raw value / format on one line, quoted like JSON.stringify
def show(a1):
r = ws.getRange(f"{a1}:{a1}")
disp = json.dumps(r.getValue()[a1], ensure\_ascii=False)
raw = json.dumps(r.getValue(True)[a1], ensure\_ascii=False) # True for the raw value
fmt = json.dumps(r.getNumberFormat()[a1], ensure\_ascii=False)
print(f" {a1}: disp={disp} raw={raw} fmt={fmt}")
def num(x): # render integer-valued floats as ints (to match the JS output)
return int(x) if isinstance(x, float) and x.is\_integer() else x
def boolstr(x): # render Python's True/False as JS-style true/false
return str(x).lower()
# --- dates & times ---
# Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 1, 15)
dt = pyosbxl.dto.DateTimeObject() # variable named dt to avoid the stdlib module name datetime
dt.setYMD(2024, 1, 15)
dt.setHMS(13, 45, 30)
time = pyosbxl.dto.DateTimeObject()
time.setHMS(9, 5, 0)
print("[Dates & times]")
ws.getRange("A1:A1").setDateValue(date) # date only
ws.getRange("A2:A2").setDateValue(dt) # date + time
ws.getRange("A3:A3").setDateValue(time) # time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15") # from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日") # Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, False, "yyyy-mm-dd") # number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, False, "ggge年m月d日") # wareki + unquoted Japanese literals
for a in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"]:
show(a)
# --- numbers & strings ---
print("[Numbers & strings]")
ws.getRange("B1:B1").setNumberValue(1234567.891)
ws.getRange("B2:B2").setNumberValue(1234567.891, False, "#,##0.00")
ws.getRange("B3:B3").setNumberValue(0.25, False, "0.0%")
ws.getRange("B4:B4").setValue("00123") # leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15") # the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", True) # forceString=True forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, False, '"円"#,##0') # quoted literal
ws.getRange("B8:B8").setNumberValue(-1234.5, False, "#,##0;赤-#,##0") # bare color token (auto-converted to [赤])
for b in ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"]:
show(b)
# --- booleans ---
print("[Booleans]")
ws.getRange("C1:C1").setBooleanValue(True)
ws.getRange("C2:C2").setBooleanValue(False)
ws.getRange("C3:C3").setFormula("TRUE") # a formula can produce a boolean too (no leading "=")
for c in ["C1", "C2", "C3"]:
show(c)
# --- date serial values ---
print("[Serial values]")
print(" serial value of 2024/1/15 =", num(app.getNumericValue(date)))
print(" with the 1904 date system =", num(app.getNumericValue(date, True)))
print(" wb.isDate1904() =", boolstr(wb.isDate1904()))
wb.save()
wb.close()
```

#### [Java]

```java
// value-types.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import java.util.Map;
public class Main {
static AppWrapper app;
static WorkSheetWrapper ws;
// Quoting equivalent to JS JSON.stringify (assumes no control characters in the sample)
static String jsonq(String s) {
return "\"" + s.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
}
// Render integer-valued doubles as ints (to match the JS output)
static String num(double d) {
if (d == Math.rint(d) && !Double.isInfinite(d)) return String.valueOf((long) d);
return String.valueOf(d);
}
static void show(String a1) {
RangeWrapper r = ws.getRange(a1 + ":" + a1);
Map<String, String> disp = r.getValue();
Map<String, String> raw = r.getValue(true); // true for the raw value
Map<String, String> fmt = r.getNumberFormat();
System.out.println(" " + a1 + ": disp=" + jsonq(disp.get(a1))
+ " raw=" + jsonq(raw.get(a1))
+ " fmt=" + jsonq(fmt.get(a1)));
}
public static void main(String[] args) {
app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("value-types-java.xlsx");
ws = wb.openWorkSheet("Sheet1");
// --- dates & times ---
// Create a DateTimeObjectWrapper with no arguments, then set it with setYMD() / setHMS()
DateTimeObjectWrapper date = new DateTimeObjectWrapper();
date.setYMD(2024, 1, 15);
DateTimeObjectWrapper datetime = new DateTimeObjectWrapper();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);
DateTimeObjectWrapper time = new DateTimeObjectWrapper();
time.setHMS(9, 5, 0);
System.out.println("[Dates & times]");
ws.getRange("A1:A1").setDateValue(date); // date only
ws.getRange("A2:A2").setDateValue(datetime); // date + time
ws.getRange("A3:A3").setDateValue(time); // time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15"); // from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日"); // Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd"); // number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals
for (String a : new String[]{"A1", "A2", "A3", "A4", "A5", "A6", "A7"}) show(a);
// --- numbers & strings ---
System.out.println("[Numbers & strings]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123"); // leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15"); // the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", true); // forceString=true forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, false, "\"円\"#,##0"); // quoted literal ("yen")
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red)
for (String b : new String[]{"B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"}) show(b);
// --- booleans ---
System.out.println("[Booleans]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE"); // a formula can produce a boolean too (no leading "=")
for (String c : new String[]{"C1", "C2", "C3"}) show(c);
// --- date serial values ---
System.out.println("[Serial values]");
System.out.println(" serial value of 2024/1/15 = " + num(app.getNumericValue(date)));
System.out.println(" with the 1904 date system = " + num(app.getNumericValue(date, true)));
System.out.println(" wb.isDate1904() = " + wb.isDate1904());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// value-types.md — same verification as the Node.js sample (phposbxl)
// Body only (<?php and the autoload require are added by the harness)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("value-types-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// Quoting equivalent to JS JSON.stringify
function jsonq($s) {
return json\_encode($s, JSON\_UNESCAPED\_UNICODE | JSON\_UNESCAPED\_SLASHES);
}
// Render integer-valued floats as ints (to match the JS output)
function num($x) {
return $x == floor($x) ? (string)(int)$x : (string)$x;
}
// Render PHP booleans as JS-style true/false
function boolstr($x) {
return $x ? "true" : "false";
}
function show($a1) {
global $ws;
$r = $ws->getRange("{$a1}:{$a1}");
echo " {$a1}: disp=" . jsonq($r->getValue()[$a1])
. " raw=" . jsonq($r->getValue(true)[$a1])
. " fmt=" . jsonq($r->getNumberFormat()[$a1]) . "\n";
}
// --- dates & times ---
// Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
$date = new \phposbxl\dto\DateTimeObject();
$date->setYMD(2024, 1, 15);
$datetime = new \phposbxl\dto\DateTimeObject();
$datetime->setYMD(2024, 1, 15);
$datetime->setHMS(13, 45, 30);
$time = new \phposbxl\dto\DateTimeObject();
$time->setHMS(9, 5, 0);
echo "[Dates & times]\n";
$ws->getRange("A1:A1")->setDateValue($date); // date only
$ws->getRange("A2:A2")->setDateValue($datetime); // date + time
$ws->getRange("A3:A3")->setDateValue($time); // time only
$ws->getRange("A4:A4")->setDateStringValue("2024-01-15"); // from a string
$ws->getRange("A5:A5")->setDateStringValue("令和6年1月15日"); // Japanese era (wareki) dates are recognized too
$ws->getRange("A6:A6")->setDateValue($date, false, "yyyy-mm-dd"); // number format can be given at the same time
$ws->getRange("A7:A7")->setDateValue($date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals also work
foreach (["A1", "A2", "A3", "A4", "A5", "A6", "A7"] as $a) show($a);
// --- numbers & strings ---
echo "[Numbers & strings]\n";
$ws->getRange("B1:B1")->setNumberValue(1234567.891);
$ws->getRange("B2:B2")->setNumberValue(1234567.891, false, "#,##0.00");
$ws->getRange("B3:B3")->setNumberValue(0.25, false, "0.0%");
$ws->getRange("B4:B4")->setValue("00123"); // leading zeros are kept as a string
$ws->getRange("B5:B5")->setValue("2024/1/15"); // the generic setter auto-detects dates
$ws->getRange("B6:B6")->setValue("12345", true); // forceString=true forces string handling
$ws->getRange("B7:B7")->setNumberValue(1234.5, false, '"円"#,##0'); // quoted literal ("yen")
$ws->getRange("B8:B8")->setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red) → [Red]
foreach (["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"] as $bx) show($bx);
// --- booleans ---
echo "[Booleans]\n";
$ws->getRange("C1:C1")->setBooleanValue(true);
$ws->getRange("C2:C2")->setBooleanValue(false);
$ws->getRange("C3:C3")->setFormula("TRUE"); // a formula can produce a boolean too
foreach (["C1", "C2", "C3"] as $c) show($c);
// --- date serial values ---
echo "[Serial values]\n";
echo " serial value of 2024/1/15 = " . num($app->getNumericValue($date)) . "\n";
echo " with the 1904 date system = " . num($app->getNumericValue($date, true)) . "\n";
echo " wb.isDate1904() = " . boolstr($wb->isDate1904()) . "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// value-types.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("value-types-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
const show = (a1) => {
const r = ws.getRange(`${a1}:${a1}`);
console.log(` ${a1}: disp=${JSON.stringify(r.getValue()[a1])}` +
` raw=${JSON.stringify(r.getValue(true)[a1])}` +
` fmt=${JSON.stringify(r.getNumberFormat()[a1])}`);
};
// --- dates & times ---
// Create a DateTimeObject with no arguments, then set its contents with setYMD() / setHMS()
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);
const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);
const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);
console.log("[Dates & times]");
ws.getRange("A1:A1").setDateValue(date); // date only
ws.getRange("A2:A2").setDateValue(datetime); // date + time
ws.getRange("A3:A3").setDateValue(time); // time only
ws.getRange("A4:A4").setDateStringValue("2024-01-15"); // from a string
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日"); // Japanese era (wareki) dates are recognized too
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd"); // number format can be given at the same time
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // wareki + unquoted Japanese literals also work
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);
// --- numbers & strings ---
console.log("[Numbers & strings]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123"); // leading zeros are kept as a string
ws.getRange("B5:B5").setValue("2024/1/15"); // the generic setter auto-detects dates
ws.getRange("B6:B6").setValue("12345", true); // forceString=true forces string handling
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0'); // quoted literal ("yen")
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // bare color token (red) → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);
// --- booleans ---
console.log("[Booleans]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE"); // a formula can produce a boolean too
["C1", "C2", "C3"].forEach(show);
// --- date serial values ---
console.log("[Serial values]");
console.log(" serial value of 2024/1/15 =", app.getNumericValue(date));
console.log(" with the 1904 date system =", app.getNumericValue(date, true));
console.log(" wb.isDate1904() =", wb.isDate1904());
wb.save();
wb.close();
```

```plaintext
[Dates & times]
A1: disp="2024/1/15" raw="45306" fmt="yyyy/m/d;@"
A2: disp="2024/1/15 13:45:30" raw="45306.573263888888" fmt="yyyy/m/d h:mm:ss;@"
A3: disp="9:05:00" raw="0.37847222222222221" fmt="h:mm:ss;@"
A4: disp="2024/1/15" raw="45306" fmt="yyyy/m/d;@"
A5: disp="2024/1/15" raw="45306" fmt="yyyy/m/d;@"
A6: disp="2024-01-15" raw="45306" fmt="yyyy-mm-dd"
A7: disp="令和6年1月15日" raw="45306" fmt="ggge年m月d日"
[Numbers & strings]
B1: disp="1234567.891" raw="1234567.8910000001" fmt="General"
B2: disp="1,234,567.89" raw="1234567.8910000001" fmt="#,##0.00"
B3: disp="25.0%" raw="0.25" fmt="0.0%"
B4: disp="00123" raw="00123" fmt="General"
B5: disp="2024/01/15" raw="45306" fmt="yyyy/mm/dd"
B6: disp="12345" raw="12345" fmt="General"
B7: disp="円1,235" raw="1234.5" fmt="\"円\"#,##0"
B8: disp="-1,235" raw="-1234.5" fmt="#,##0;赤-#,##0"
[Booleans]
C1: disp="TRUE" raw="TRUE" fmt="General"
C2: disp="FALSE" raw="FALSE" fmt="General"
C3: disp="TRUE" raw="TRUE" fmt="General"
[Serial values]
serial value of 2024/1/15 = 45306
with the 1904 date system = 43844
wb.isDate1904() = false
```

## Notes

### `getValue()` vs. `getValue(true)`

| Call | Returns | Use it for |
| --- | --- | --- |
| `getValue()` | **Display value** (string after applying the number format) | Report output, on-screen display, CSV export |
| `getValue(true)` | **Raw value** (internally stored value; dates are serial values, formulas are results) | Calculation, comparison, numeric processing |

Both return a map like `{ "A1": "...", ... }` and **all values are strings**. Convert with `Number(...)` when treating them as numbers.

#### [Node.js]

```javascript
const raw = Number(ws.getRange("A1:A1").getValue(true)["A1"]); // 45306
const disp = ws.getRange("A1:A1").getValue()["A1"]; // "2024/1/15"
```

### Dates

* **Create a `DateTimeObject` with no arguments and set it with `setYMD()` / `setHMS()`.** You can also set the fields individually with `setYear()` / `setMonth()` / `setDay()` / `setHour()` / `setMinute()` / `setSecond()`.
* **If you do not specify `numberFormat`, a default format is set automatically depending on the contents.**
  + Date only → `yyyy/m/d;@`
  + Time only → `h:mm:ss;@`
  + Date + time → `yyyy/m/d h:mm:ss;@`
* **`setDateStringValue()` recognizes the same date strings Excel does.** `"2024-01-15"`, `"2024/1/15 13:45:30"` and even the Japanese era string `"令和6年1月15日"` go in as dates.
* **`app.getNumericValue(dateTimeObject)` returns the serial value.** Pass `true` as the second argument `is1904` to get the serial value in the 1904 date system.  
  Use `wb.isDate1904()` to check which system the book uses.

### Numbers and strings

* **`setValue()` auto-detects the type.** `"123"` → number, `"2024/1/15"` → date, `"00123"` → string (leading zero preserved), `"12A"` → string.
* **Use the dedicated setter when you want the intended type for sure.** `setNumberValue()` for numbers; set `forceString` to `true` to pin a value as a string.
* **`numberFormat` can be given as the setter's third argument at the same time.** The result is the same as calling `setNumberFormat()` afterwards.

## Caveats

* **Set the number format AFTER putting the value in.** If you call `setDateValue()` after `setNumberFormat()`, the date setter **overwrites** the format with the default (measured).

  ```javascript
  // ✗ the format is lost
  ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");
  ws.getRange("A1:A1").setDateValue(date); // → format gets overwritten with "yyyy/m/d;@"
  // ✓ the format survives
  ws.getRange("A1:A1").setDateValue(date);
  ws.getRange("A1:A1").setNumberFormat("yyyy-mm-dd");
  // ✓ or give it at the same time via the setter's third argument
  ws.getRange("A1:A1").setDateValue(date, false, "yyyy-mm-dd");
  ```
* **`setDateStringValue()` throws on strings it cannot recognize.**

  ```javascript
  ws.getRange("A1:A1").setDateStringValue("this is not a date");
  // → exception: can't recognize input value as DateTime
  ```

  To avoid the exception, pass `true` as the second argument `forceString`. When recognition fails, the value is then entered as a plain string.
* **Number formats also support Japanese text, currency symbols and color specs** (verified 2026-09-30). Quoted literals (`"円"#,##0`) and unquoted Japanese text (`yyyy年m月d日`) display as-is, and the Japanese era calendar (`ggge年m月d日` → `令和6年1月15日`) and `¥#,##0` also work.  
  Color specs accept all three spellings `[Red]` / `[赤]` / bare `赤`, and all of them are stored as `[Red]`. Examples of verified formats: | Kind | Format | Example display | |---|---|---| | Date | `yyyy/m/d` / `yyyy/mm/dd` / `yyyy-mm-dd` / `yy/m/d` / `m/d` / `d-mmm-yyyy` | `2024/1/15`, `2024-01-15`, `15-Jan-2024` | | Date (Japanese) | `yyyy年m月d日` / `ggge年m月d日` | `2024年1月15日` / `令和6年1月15日` | | Number | `General` / `#,##0` / `#,##0.00` / `0.000` / `0%` / `0.0%` | `1,235`, `1,234.57`, `25.0%` | | Number (Japanese text & currency) | `"円"#,##0` / `¥#,##0` | `円1,235` / `¥1,235` | | Number (negative section) | `#,##0;[Red]-#,##0` / `#,##0;赤-#,##0` | `-1,235` (in red) |
* **Raw number values contain double representation error.** The raw value of `1234567.891` is `"1234567.8910000001"`. For comparisons and display, use `getValue()` (the display value) or round the number first.
* **`getFormula()` results have no leading `=`.** For `setFormula("SUM(A1:A3)")`, `getFormula()` returns `"SUM(A1:A3)"`.

## See also

* [Index](index.md)
* Previous: [Working with books and worksheets](general-03-sheets.md) / Next: [Working with rows, columns and cell ranges](general-05-rows-cols.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
