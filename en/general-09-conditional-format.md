# Conditional Formatting — rules, scales, data bars

level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](../ja/general-09-conditional-format.md)

## What you want to do

Set **conditional formatting rules** programmatically, such as "make only cells greater than 60 red" or "show the magnitude of values with a gradient or a bar". Also **read back the count, type, thresholds and applied ranges of the registered rules** and **remove** them.

Conditional formatting is handled by the `FormatConditions` class obtained via `ws.getFormatConditions()`. The flow is always these 3 steps:

1. Create a condition object from `dto` (e.g. `CellValueObject`) and **set its apply range and format**
2. Register it with `fc.set…Condition(priority, obj)` (`priority` is the evaluation order,

1-based and unique per sheet)

1. **Call `fc.apply()` after all conditions are registered** (required)

## Classes used

### `FormatConditions` (main methods)

| Condition type | Set | Read back | Condition object |
| --- | --- | --- | --- |
| Cell value (greater, less, ...) | `setCellValueCondition(p, obj)` | `getCellValueCondition(p)` | `dto.CellValueObject` |
| Cell value range (between) | `setCellValueRangeCondition(p, obj)` | `getCellValueRangeCondition(p)` | `dto.CellValueRangeObject` |
| Color scale | `setColorScaleFormatCondition(p, obj)` | `getColorScaleFormatCondition(p)` | `dto.ColorScaleObject` |
| Data bar | `setDataBarFormatCondition(p, obj)` | `getDataBarFormatCondition(p)` | `dto.DataBarObject` |
| Top / bottom N | `setTop10FormatCondition(p, obj)` | `getTop10FormatCondition(p)` | `dto.Top10Object` |
| Above/below average | `setAboveAverageCondition(p, obj)` | `getAboveAverageCondition(p)` | `dto.AboveAverageObject` |
| Icon set | `setIconsetFormatCondition(p, obj)` | `getIconsetFormatCondition(p)` | `dto.IconsetConditionObject` |
| Unique values | `setUniqueValuesFormatCondition(p, obj)` | `getUniqueValuesFormatCondition(p)` | `dto.UniqueValuesObject` |

| Other methods | Description |
| --- | --- |
| `apply()` | Evaluates and reflects the conditions. **Always call after setting all conditions.** |
| `getConditionTypes()` | List of all rules; a nested map `{ priority: { cellRange: conditionType } }`. |
| `removeCondition(p)` / `removeAllCondition()` | Remove one rule / all rules. |
| `getAppliedFontObject(range)` etc. | Get the cells that met a condition and actually received formatting (use after `apply()`). |

### Methods common to all condition objects

They all have `setA1C1(applyRange)` and `setFont(dto.FontObject)` / `setPatternFill(dto.PatternFillObject)`. Colors are set on a `dto.ColorObject` via `setHexColor()` with a **6-digit hex string without `#`** (e.g. `"9C0006"`).

## Code

Output:

#### [Node.js]

```javascript
// conditional-format.js — adding, reading back and removing conditional format rules
const { nodeosbxl, enums, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("condfmt.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// color helper: ColorObject takes a 6-digit hex without "#"
const color = (hex) => {
const c = new dto.ColorObject();
c.setHexColor(hex);
return c;
};
// ---------- data table ----------
// col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"].forEach((h, i) => {
const col = "ABCDE"[i];
ws.getRange(`${col}1:${col}1`).setValue(h);
});
const sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
sales.forEach((v, i) => {
ws.getRange(`A${i + 2}:A${i + 2}`).setNumberValue(i + 1);
for (const col of ["B", "C", "D", "E"]) {
ws.getRange(`${col}${i + 2}:${col}${i + 2}`).setNumberValue(v);
}
});
const fc = ws.getFormatConditions();
// ---------- (1) cell value rule: greater than 60 -> dark red bold + light red fill ----------
const cellVal = new dto.CellValueObject();
cellVal.setValueCondition(60, enums.XlFormatConditionOperator.FormatConditionGreater);
cellVal.setA1C1("B2:B11");
const redFont = new dto.FontObject();
redFont.setBold(true);
redFont.setColorObject(color("9C0006"));
cellVal.setFont(redFont);
const pinkFill = new dto.PatternFillObject();
pinkFill.setPattern(enums.XlPatternType.PatternSolid);
pinkFill.setCellColor(color("FFC7CE"));
cellVal.setPatternFill(pinkFill);
fc.setCellValueCondition(1, cellVal);
// ---------- (2) color scale: 3-color gradient on column C ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);
// ---------- (3) data bar: column D ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);
// ---------- (4) top 3 + above average: column E ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true); // true = top (false = bottom)
top3.setA1C1("E2:E11");
const blueFont = new dto.FontObject();
blueFont.setBold(true);
blueFont.setColorObject(color("0070C0"));
top3.setFont(blueFont);
fc.setTop10FormatCondition(4, top3);
const aboveAvg = new dto.AboveAverageObject();
aboveAvg.setAboveBelow(enums.XlAboveBelow.XlAboveAverage);
aboveAvg.setA1C1("E2:E11");
const greenFont = new dto.FontObject();
greenFont.setBold(true);
greenFont.setColorObject(color("006100"));
aboveAvg.setFont(greenFont);
fc.setAboveAverageCondition(5, aboveAvg);
fc.apply(); // apply() is required after setting all conditions
console.log("[1] Registered 5 rules and called apply()");
// ---------- read back the rule list ----------
const typeName = Object.fromEntries(
Object.entries(enums.XlFormatConditionType).map(([name, val]) => [val, name]));
console.log(`[2] rule count = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
for (const [range, type] of Object.entries(ranges)) {
console.log(` priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
}
}
// ---------- read back each rule's details ----------
console.log("[3] Reading back rule details");
const g1 = fc.getCellValueCondition(1);
console.log(` cell value : ${g1.getA1C1()} value=${g1.getConditionValue()}` +
` operator=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
` fontColor=${g1.getFont().getColorObject().getHexColor()} bold=${g1.getFont().isBold()}` +
` fill=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(` color scale: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, threshold=${g2.getMediumValue()})` +
` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(` data bar: color=${g3.getDataBarColor().getHexColor()}` +
` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(` top N: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(` above average: aboveBelow=${g5.getAboveBelow()}` +
` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);
// ---------- cells that met the conditions and got formatted (available after apply()) ----------
const byRow = (a, b) => Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] Cells that received formatting");
console.log(` col B (>60) : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(` col E (top/avg) : ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(` col D data bars : ${Object.values(appliedBar).filter(Boolean).length} cells`);
// ---------- removal ----------
fc.removeCondition(4); // remove only the top-3 rule
console.log(`[5] rule count after removeCondition(4) = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition(); // remove all rules
console.log(`[6] rule count after removeAllCondition() = ${Object.keys(fc.getConditionTypes()).length}`);
// re-register the cell value rule and save (so the effect is visible when the file is opened)
fc.setCellValueCondition(1, cellVal);
fc.apply();
wb.save();
wb.close();
console.log("[done] saved condfmt.xlsx (cell-value rule re-registered)");
```

#### [Python]

```python
# conditional-format.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl
e = pyosbxl.enums
def color(hex): # ColorObject takes a 6-digit hex without "#"
c = pyosbxl.dto.ColorObject()
c.setHexColor(hex)
return c
def num(v): # print numbers like JS does (60.0 -> "60")
return str(int(v)) if float(v) == int(v) else str(v)
def bol(v): # print booleans like JS does (True -> "true")
return str(bool(v)).lower()
app = pyosbxl.App()
wb = app.createWorkBook("condfmt-py.xlsx") # output text says condfmt.xlsx to match the Node.js sample
ws = wb.openWorkSheet("Sheet1")
# ---------- data table ----------
# col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
for i, h in enumerate(["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"]):
col = "ABCDE"[i]
ws.getRange(f"{col}1:{col}1").setValue(h)
sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49]
for i, v in enumerate(sales):
ws.getRange(f"A{i + 2}:A{i + 2}").setNumberValue(i + 1)
for col in ["B", "C", "D", "E"]:
ws.getRange(f"{col}{i + 2}:{col}{i + 2}").setNumberValue(v)
fc = ws.getFormatConditions()
# ---------- (1) cell value rule: greater than 60 -> dark red bold + light red fill ----------
cellVal = pyosbxl.dto.CellValueObject()
cellVal.setValueCondition(60, e.XlFormatConditionOperator.FormatConditionGreater)
cellVal.setA1C1("B2:B11")
redFont = pyosbxl.dto.FontObject()
redFont.setBold(True)
redFont.setColorObject(color("9C0006"))
cellVal.setFont(redFont)
pinkFill = pyosbxl.dto.PatternFillObject()
pinkFill.setPattern(e.XlPatternType.PatternSolid)
pinkFill.setCellColor(color("FFC7CE"))
cellVal.setPatternFill(pinkFill)
fc.setCellValueCondition(1, cellVal)
# ---------- (2) color scale: 3-color gradient on column C ----------
colorScale = pyosbxl.dto.ColorScaleObject()
colorScale.setA1C1("C2:C11")
colorScale.setMinimumCondition(color("F8696B"), e.XlConditionValueType.ConditionValueTypeMinimum)
colorScale.setMediumCondition(color("FFEB84"), e.XlConditionValueType.ConditionValueTypePercentile, 50)
colorScale.setMaximumCondition(color("63BE7B"), e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setColorScaleFormatCondition(2, colorScale)
# ---------- (3) data bar: column D ----------
dataBar = pyosbxl.dto.DataBarObject()
dataBar.setA1C1("D2:D11")
dataBar.setDataBarColor(color("638EC6"))
dataBar.setMinimumCondition(e.XlConditionValueType.ConditionValueTypeMinimum)
dataBar.setMaximumCondition(e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setDataBarFormatCondition(3, dataBar)
# ---------- (4) top 3 + above average: column E ----------
top3 = pyosbxl.dto.Top10Object()
top3.setRank(3, True) # True = top (False = bottom)
top3.setA1C1("E2:E11")
blueFont = pyosbxl.dto.FontObject()
blueFont.setBold(True)
blueFont.setColorObject(color("0070C0"))
top3.setFont(blueFont)
fc.setTop10FormatCondition(4, top3)
aboveAvg = pyosbxl.dto.AboveAverageObject()
aboveAvg.setAboveBelow(e.XlAboveBelow.XlAboveAverage)
aboveAvg.setA1C1("E2:E11")
greenFont = pyosbxl.dto.FontObject()
greenFont.setBold(True)
greenFont.setColorObject(color("006100"))
aboveAvg.setFont(greenFont)
fc.setAboveAverageCondition(5, aboveAvg)
fc.apply() # apply() is required after setting all conditions
print("[1] Registered 5 rules and called apply()")
# ---------- read back the rule list ----------
types = fc.getConditionTypes()
print(f"[2] rule count = {len(types)}")
for prio in sorted(types):
for rng, t in types[prio].items():
print(f" priority {prio}: {rng} → {t.name} (type={int(t)})")
# ---------- read back each rule's details ----------
print("[3] Reading back rule details")
g1 = fc.getCellValueCondition(1)
print(f" cell value : {g1.getA1C1()} value={num(g1.getConditionValue())}"
f" operator={int(g1.getConditionOperetor())}"
f" (Greater={int(e.XlFormatConditionOperator.FormatConditionGreater)})"
f" fontColor={g1.getFont().getColorObject().getHexColor()} bold={bol(g1.getFont().isBold())}"
f" fill={g1.getPatternFill().getCellColor().getHexColor()}")
g2 = fc.getColorScaleFormatCondition(2)
print(f" color scale: min={g2.getMinimumColor().getHexColor()} (type={int(g2.getMinimumValueType())})"
f" mid={g2.getMediumColor().getHexColor()} (type={int(g2.getMediumValueType())}, threshold={num(g2.getMediumValue())})"
f" max={g2.getMaximumColor().getHexColor()} (type={int(g2.getMaximumValueType())})")
g3 = fc.getDataBarFormatCondition(3)
print(f" data bar: color={g3.getDataBarColor().getHexColor()}"
f" minType={int(g3.getMinimumCondition())} maxType={int(g3.getMaximumCondition())}")
g4 = fc.getTop10FormatCondition(4)
print(f" top N: rank={g4.getRank()} isTop={bol(g4.getTopBottom())}")
g5 = fc.getAboveAverageCondition(5)
print(f" above average: aboveBelow={int(g5.getAboveBelow())}"
f" (XlAboveAverage={int(e.XlAboveBelow.XlAboveAverage)})")
# ---------- cells that met the conditions and got formatted (available after apply()) ----------
byRow = lambda a: int(a[1:])
print("[4] Cells that received formatting")
appliedB = sorted(fc.getAppliedFontObject("B2:B11").keys(), key=byRow)
print(f" col B (>60) : {' '.join(appliedB)}")
appliedE = sorted(fc.getAppliedFontObject("E2:E11").keys(), key=byRow)
print(f" col E (top/avg) : {' '.join(appliedE)}")
appliedBar = fc.getAppliedDataBar("D2:D11")
print(f" col D data bars : {sum(1 for v in appliedBar.values() if v)} cells")
# ---------- removal ----------
fc.removeCondition(4) # remove only the top-3 rule
print(f"[5] rule count after removeCondition(4) = {len(fc.getConditionTypes())}")
fc.removeAllCondition() # remove all rules
print(f"[6] rule count after removeAllCondition() = {len(fc.getConditionTypes())}")
# re-register the cell value rule and save (so the effect is visible when the file is opened)
fc.setCellValueCondition(1, cellVal)
fc.apply()
wb.save()
wb.close()
print("[done] saved condfmt.xlsx (cell-value rule re-registered)")
```

#### [Java]

```java
// conditional-format.java — checks how much of the Node.js sample javaosbxl supports
//
// [API gap] In javaosbxl 1.5.0, the condition-object wrappers (CellValueObjectWrapper etc.)
// do not expose the base-class methods setA1C1() / setFont() / setPatternFill().
// With no way to set the apply range (A1C1), setCellValueCondition() throws
// "apply range(A1C1) must be set", so conditional format rules cannot be registered.
// This block demonstrates that registering the same rule as the Node.js version throws,
// while apply() / getConditionTypes() / removeCondition() / removeAllCondition() can all
// be called without exceptions, making the gap explicit.
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("condfmt-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- data table (same as the Node.js version) ----------
String[] headers = {"Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"};
for (int i = 0; i < headers.length; i++) {
String col = "ABCDE".substring(i, i + 1);
ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
}
int[] sales = {45, 78, 23, 91, 56, 34, 87, 12, 65, 49};
for (int i = 0; i < sales.length; i++) {
ws.getRange("A" + (i + 2) + ":A" + (i + 2)).setNumberValue(i + 1);
for (String col : new String[]{"B", "C", "D", "E"}) {
ws.getRange(col + (i + 2) + ":" + col + (i + 2)).setNumberValue(sales[i]);
}
}
FormatConditionsWrapper fc = ws.getFormatConditions();
// ---------- (1) cell value rule: greater than 60 ----------
// The Node.js version calls cellVal.setA1C1("B2:B11") / setFont(...) / setPatternFill(...),
// but javaosbxl's CellValueObjectWrapper does not have these methods (API gap)
CellValueObjectWrapper cellVal = new CellValueObjectWrapper();
cellVal.setValueCondition(60, XlFormatConditionOperator.FormatConditionGreater);
try {
fc.setCellValueCondition(1, cellVal);
System.out.println("[1] Registered the cell-value rule");
} catch (RuntimeException e) {
// registration always throws because A1C1 cannot be set (known API gap in javaosbxl 1.5.0)
System.out.println("[1] Registering the cell-value rule threw: " + e.getMessage());
}
fc.apply(); // apply() can be called without exceptions even with 0 rules
// ---------- read back the rule list (nothing was registered, so 0) ----------
Map<Integer, Map<String, XlFormatConditionType>> types = fc.getConditionTypes();
System.out.println("[2] rule count = " + types.size());
// ---------- removal methods also work without exceptions at 0 rules ----------
fc.removeCondition(4);
System.out.println("[5] rule count after removeCondition(4) = " + fc.getConditionTypes().size());
fc.removeAllCondition();
System.out.println("[6] rule count after removeAllCondition() = " + fc.getConditionTypes().size());
wb.save();
wb.close();
System.out.println("[done] saved condfmt-java.xlsx (0 rules because javaosbxl cannot register conditional format rules)");
}
}
```

#### [PHP]

```php
// conditional-format.md — same verification as the Node.js sample (phposbxl). Body only (<?php / autoload provided by the harness)
// Output helpers: match the Node.js stdout representation (bool -> "true"/"false")
$bol = fn($v) => $v ? "true" : "false";
// Color helper: ColorObject takes a 6-digit hex without "#"
$color = function (string $hex): \phposbxl\dto\ColorObject {
$c = new \phposbxl\dto\ColorObject();
$c->setHexColor($hex);
return $c;
};
$app = new \phposbxl\App();
$wb = $app->createWorkBook("condfmt-php.xlsx"); // output text says condfmt.xlsx to match the Node.js sample
$ws = $wb->openWorkSheet("Sheet1");
// ---------- data table ----------
// col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
foreach (["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"] as $i => $h) {
$col = "ABCDE"[$i];
$ws->getRange("{$col}1:{$col}1")->setValue($h);
}
$sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
foreach ($sales as $i => $v) {
$r = $i + 2;
$ws->getRange("A{$r}:A{$r}")->setNumberValue($i + 1);
foreach (["B", "C", "D", "E"] as $col) {
$ws->getRange("{$col}{$r}:{$col}{$r}")->setNumberValue($v);
}
}
$fc = $ws->getFormatConditions();
// ---------- (1) cell value rule: greater than 60 -> dark red bold + light red fill ----------
$cellVal = new \phposbxl\dto\CellValueObject();
$cellVal->setValueCondition(60, \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater);
$cellVal->setA1C1("B2:B11");
$redFont = new \phposbxl\dto\FontObject();
$redFont->setBold(true);
$redFont->setColorObject($color("9C0006"));
$cellVal->setFont($redFont);
$pinkFill = new \phposbxl\dto\PatternFillObject();
$pinkFill->setPattern(\phposbxl\enums\XlPatternType::PatternSolid);
$pinkFill->setCellColor($color("FFC7CE"));
$cellVal->setPatternFill($pinkFill);
$fc->setCellValueCondition(1, $cellVal);
// ---------- (2) color scale: 3-color gradient on column C ----------
$colorScale = new \phposbxl\dto\ColorScaleObject();
$colorScale->setA1C1("C2:C11");
$colorScale->setMinimumCondition($color("F8696B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$colorScale->setMediumCondition($color("FFEB84"), \phposbxl\enums\XlConditionValueType::ConditionValueTypePercentile, 50);
$colorScale->setMaximumCondition($color("63BE7B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc->setColorScaleFormatCondition(2, $colorScale);
// ---------- (3) data bar: column D ----------
$dataBar = new \phposbxl\dto\DataBarObject();
$dataBar->setA1C1("D2:D11");
$dataBar->setDataBarColor($color("638EC6"));
$dataBar->setMinimumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$dataBar->setMaximumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc->setDataBarFormatCondition(3, $dataBar);
// ---------- (4) top 3 + above average: column E ----------
$top3 = new \phposbxl\dto\Top10Object();
$top3->setRank(3, true); // true = top (false = bottom)
$top3->setA1C1("E2:E11");
$blueFont = new \phposbxl\dto\FontObject();
$blueFont->setBold(true);
$blueFont->setColorObject($color("0070C0"));
$top3->setFont($blueFont);
$fc->setTop10FormatCondition(4, $top3);
$aboveAvg = new \phposbxl\dto\AboveAverageObject();
$aboveAvg->setAboveBelow(\phposbxl\enums\XlAboveBelow::XlAboveAverage);
$aboveAvg->setA1C1("E2:E11");
$greenFont = new \phposbxl\dto\FontObject();
$greenFont->setBold(true);
$greenFont->setColorObject($color("006100"));
$aboveAvg->setFont($greenFont);
$fc->setAboveAverageCondition(5, $aboveAvg);
$fc->apply(); // apply() is required after setting all conditions
echo "[1] Registered 5 rules and called apply()\n";
// ---------- read back the rule list ----------
$types = $fc->getConditionTypes(); // array<int, array<string, int>>
ksort($types); // iterate in ascending priority order, like Object.entries in the Node.js version
echo "[2] rule count = " . count($types) . "\n";
foreach ($types as $prio => $ranges) {
foreach ($ranges as $range => $type) {
$name = \phposbxl\enums\XlFormatConditionType::from($type)->name;
echo " priority {$prio}: {$range} → {$name} (type={$type})\n";
}
}
// ---------- read back each rule's details ----------
echo "[3] Reading back rule details\n";
$g1 = $fc->getCellValueCondition(1);
$greater = \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater->value;
echo " cell value : {$g1->getA1C1()} value={$g1->getConditionValue()}"
. " operator={$g1->getConditionOperetor()} (Greater={$greater})"
. " fontColor={$g1->getFont()->getColorObject()->getHexColor()} bold=" . $bol($g1->getFont()->isBold())
. " fill={$g1->getPatternFill()->getCellColor()->getHexColor()}\n";
$g2 = $fc->getColorScaleFormatCondition(2);
echo " color scale: min={$g2->getMinimumColor()->getHexColor()} (type={$g2->getMinimumValueType()})"
. " mid={$g2->getMediumColor()->getHexColor()} (type={$g2->getMediumValueType()}, threshold={$g2->getMediumValue()})"
. " max={$g2->getMaximumColor()->getHexColor()} (type={$g2->getMaximumValueType()})\n";
$g3 = $fc->getDataBarFormatCondition(3);
echo " data bar: color={$g3->getDataBarColor()->getHexColor()}"
. " minType={$g3->getMinimumCondition()} maxType={$g3->getMaximumCondition()}\n";
$g4 = $fc->getTop10FormatCondition(4);
echo " top N: rank={$g4->getRank()} isTop=" . $bol($g4->getTopBottom()) . "\n";
$g5 = $fc->getAboveAverageCondition(5);
$xlAbove = \phposbxl\enums\XlAboveBelow::XlAboveAverage->value;
echo " above average: aboveBelow={$g5->getAboveBelow()} (XlAboveAverage={$xlAbove})\n";
// ---------- cells that met the conditions and got formatted (available after apply()) ----------
$byRow = fn($a, $b) => (int)substr($a, 1) - (int)substr($b, 1);
echo "[4] Cells that received formatting\n";
$appliedB = array\_keys($fc->getAppliedFontObject("B2:B11"));
usort($appliedB, $byRow);
echo " col B (>60) : " . implode(" ", $appliedB) . "\n";
$appliedE = array\_keys($fc->getAppliedFontObject("E2:E11"));
usort($appliedE, $byRow);
echo " col E (top/avg) : " . implode(" ", $appliedE) . "\n";
$appliedBar = $fc->getAppliedDataBar("D2:D11");
echo " col D data bars : " . count(array\_filter($appliedBar)) . " cells\n";
// ---------- removal ----------
$fc->removeCondition(4); // remove only the top-3 rule
echo "[5] rule count after removeCondition(4) = " . count($fc->getConditionTypes()) . "\n";
$fc->removeAllCondition(); // remove all rules
echo "[6] rule count after removeAllCondition() = " . count($fc->getConditionTypes()) . "\n";
// re-register the cell value rule and save (so the effect is visible when the file is opened)
$fc->setCellValueCondition(1, $cellVal);
$fc->apply();
$wb->save();
$wb->close();
echo "[done] saved condfmt.xlsx (cell-value rule re-registered)\n";
```

#### [WebAssembly]

```javascript
// conditional-format.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("condfmt-wasm.xlsx"); // output text says condfmt.xlsx, same as the Node.js version
const ws = wb.openWorkSheet("Sheet1");
// color helper: ColorObject takes a 6-digit hex without "#"
const color = (hex) => {
const c = new dto.ColorObject();
c.setHexColor(hex);
return c;
};
// ---------- data table ----------
// col B: cell value rule / col C: color scale / col D: data bar / col E: top 3 + above average
["Month", "Cell value rule", "Color scale", "Data bar", "Top/Average"].forEach((h, i) => {
const col = "ABCDE"[i];
ws.getRange(`${col}1:${col}1`).setValue(h);
});
const sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
sales.forEach((v, i) => {
ws.getRange(`A${i + 2}:A${i + 2}`).setNumberValue(i + 1);
for (const col of ["B", "C", "D", "E"]) {
ws.getRange(`${col}${i + 2}:${col}${i + 2}`).setNumberValue(v);
}
});
const fc = ws.getFormatConditions();
// ---------- (1) cell value rule: greater than 60 -> dark red bold + light red fill ----------
const cellVal = new dto.CellValueObject();
cellVal.setValueCondition(60, enums.XlFormatConditionOperator.FormatConditionGreater);
cellVal.setA1C1("B2:B11");
const redFont = new dto.FontObject();
redFont.setBold(true);
redFont.setColorObject(color("9C0006"));
cellVal.setFont(redFont);
const pinkFill = new dto.PatternFillObject();
pinkFill.setPattern(enums.XlPatternType.PatternSolid);
pinkFill.setCellColor(color("FFC7CE"));
cellVal.setPatternFill(pinkFill);
fc.setCellValueCondition(1, cellVal);
// ---------- (2) color scale: 3-color gradient on column C ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);
// ---------- (3) data bar: column D ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);
// ---------- (4) top 3 + above average: column E ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true); // true = top (false = bottom)
top3.setA1C1("E2:E11");
const blueFont = new dto.FontObject();
blueFont.setBold(true);
blueFont.setColorObject(color("0070C0"));
top3.setFont(blueFont);
fc.setTop10FormatCondition(4, top3);
const aboveAvg = new dto.AboveAverageObject();
aboveAvg.setAboveBelow(enums.XlAboveBelow.XlAboveAverage);
aboveAvg.setA1C1("E2:E11");
const greenFont = new dto.FontObject();
greenFont.setBold(true);
greenFont.setColorObject(color("006100"));
aboveAvg.setFont(greenFont);
fc.setAboveAverageCondition(5, aboveAvg);
fc.apply(); // apply() is required after setting all conditions
console.log("[1] Registered 5 rules and called apply()");
// ---------- read back the rule list ----------
const typeName = Object.fromEntries(
Object.entries(enums.XlFormatConditionType).map(([name, val]) => [val, name]));
console.log(`[2] rule count = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
for (const [range, type] of Object.entries(ranges)) {
console.log(` priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
}
}
// ---------- read back each rule's details ----------
console.log("[3] Reading back rule details");
const g1 = fc.getCellValueCondition(1);
console.log(` cell value : ${g1.getA1C1()} value=${g1.getConditionValue()}` +
` operator=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
` fontColor=${g1.getFont().getColorObject().getHexColor()} bold=${g1.getFont().isBold()}` +
` fill=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(` color scale: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, threshold=${g2.getMediumValue()})` +
` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(` data bar: color=${g3.getDataBarColor().getHexColor()}` +
` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(` top N: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(` above average: aboveBelow=${g5.getAboveBelow()}` +
` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);
// ---------- cells that met the conditions and got formatted (available after apply()) ----------
const byRow = (a, b) => Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] Cells that received formatting");
console.log(` col B (>60) : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(` col E (top/avg) : ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(` col D data bars : ${Object.values(appliedBar).filter(Boolean).length} cells`);
// ---------- removal ----------
fc.removeCondition(4); // remove only the top-3 rule
console.log(`[5] rule count after removeCondition(4) = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition(); // remove all rules
console.log(`[6] rule count after removeAllCondition() = ${Object.keys(fc.getConditionTypes()).length}`);
// re-register the cell value rule and save (so the effect is visible when the file is opened)
fc.setCellValueCondition(1, cellVal);
fc.apply();
wb.save();
wb.close();
console.log("[done] saved condfmt.xlsx (cell-value rule re-registered)");
```

```plaintext
[1] Registered 5 rules and called apply()
[2] rule count = 5
priority 1: B2:B11 → CellValue (type=1)
priority 2: C2:C11 → ColorScale (type=3)
priority 3: D2:D11 → Databar (type=4)
priority 4: E2:E11 → Top10 (type=5)
priority 5: E2:E11 → AboveAverageCondition (type=12)
[3] Reading back rule details
cell value : B2:B11 value=60 operator=5 (Greater=5) fontColor=9C0006 bold=true fill=FFC7CE
color scale: min=F8696B (type=5) mid=FFEB84 (type=4, threshold=50) max=63BE7B (type=6)
data bar: color=638EC6 minType=5 maxType=6
top N: rank=3 isTop=true
above average: aboveBelow=0 (XlAboveAverage=0)
[4] Cells that received formatting
col B (>60) : B3 B5 B8 B10
col E (top/avg) : E3 E5 E6 E8 E10
col D data bars : 10 cells
[5] rule count after removeCondition(4) = 4
[6] rule count after removeAllCondition() = 0
[done] saved condfmt.xlsx (cell-value rule re-registered)
```

`B3 B5 B8 B10` in column B are the cells whose values are 78, 91, 87 and 65 (= greater than 60); `E3 E5 E6 E8 E10` in column E are the union of "top 3" and "above the average (54.2)". Visual formatting cannot be verified programmatically, so the sample confirms the rules were set as intended by **reading back the rule count, types, thresholds and the cells that got formatted** and printing them.

## Notes

### priority (evaluation order)

* Rules are managed by `priority` (1-based evaluation order) and are **unique per sheet**. The keys of the map returned by `getConditionTypes()` are these priorities.
* **Registering at an existing priority shifts the existing rule back** (measured). For example, registering twice at priority 1 moves the first rule to priority 2.
* Removing with `removeCondition(p)` **bumps the later priorities up**.
* The read-back methods (`getCellValueCondition(p)` etc.) **throw when a different type of rule sits at that priority**. Check the types with `getConditionTypes()` first.

### What apply() does

* **Conditions are not evaluated until you call `apply()`.** Measured: `getAppliedFontObject()` returns an empty `{}` before `apply()`, and after `apply()` it returned the cells meeting the condition (B3 B5 B8 B10 in this sample).
* Call `apply()` every time you add or remove rules.

### Main enums

| enum | Common values |
| --- | --- |
| `XlFormatConditionOperator` | `FormatConditionGreater`(5) / `FormatConditionGreaterEqual`(7) / `FormatConditionLess`(6) / `FormatConditionLessEqual`(8) / `FormatConditionEqual`(3) / `FormatConditionBetween`(1) |
| `XlFormatConditionType` | `CellValue`(1) / `ColorScale`(3) / `Databar`(4) / `Top10`(5) / `XlIconSet`(6) / `UniqueValues`(8) / `AboveAverageCondition`(12) |
| `XlConditionValueType` | `ConditionValueTypeMinimum`(5) / `ConditionValueTypeMaximum`(6) / `ConditionValueTypePercentile`(4) / `ConditionValueTypeNumeric`(1) |
| `XlAboveBelow` | `XlAboveAverage`(0) / `XlBelowAverage`(2) / `XlAboveStdDev`(1) / `XlBelowStdDev`(3) |

### Persistence

Conditional formatting is saved into the .xlsx by `wb.save()`, and reopening the file with `app.openWorkBook()` restores `getConditionTypes()` and the individual read-back methods (verified).

## Caveats

* **Do not forget `apply()`.** Forgetting it does not throw — the `getApplied*` methods simply return empty results, which makes the mistake easy to miss (see [Notes](#what-apply-does)).
* **The operator getter is spelled `getConditionOperetor()`.** It is **`Operetor`**, not `Operator` (the library's own spelling), so be careful.
* **Data bar read-back method names differ from color scale.** On `DataBarObject`, `getMinimumCondition()` / `getMaximumCondition()` return the **`XlConditionValueType` (kind of threshold)** and the values come from `getMinimumValue()` / `getMaximumValue()`. On `ColorScaleObject`, the kind is `getMinimumValueType()` and the color is `getMinimumColor()`.
* **Colors are passed to `setHexColor()` as a 6-digit hex without `#`** (e.g. `"9C0006"`).
* **The second argument of `Top10Object.setRank(rank, isTop)` is a bool for top vs bottom** (`true` = top N, `false` = bottom N). Despite the name "Top10", N can be any number.
* **Every condition object needs its apply range set with `setA1C1()`.** The condition object owns the range; the `set…Condition()` calls take no range argument.
* \*\*The `getApplied*` methods return only "the cells that met the condition and actually got formatted"\*\* — not the whole apply range of the rule.

## See also

* [Index](index.md)
* Previous: [Tables and Comments](general-08-tables-comments.md) / Next: [Charts Basics](general-10-charts.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
