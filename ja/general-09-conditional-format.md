# 条件付き書式 — セル値ルール・カラースケール・データバー・上位 N 件

level: intermediate ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](/en/general-09-conditional-format.md)

## やりたいこと

「60 より大きいセルだけ赤くする」「値の大きさをグラデーションやバーで表す」といった **条件付き書式のルール**をプログラムから設定する。 あわせて、登録したルールの**個数・種類・しきい値・適用範囲の読み戻し**と**削除**を行う。

条件付き書式は `ws.getFormatConditions()` で取得する `FormatConditions` クラスで扱います。 流れは必ず次の 3 ステップです。

1. `dto` の条件オブジェクト（`CellValueObject` など）を作り、**適用範囲と書式をセット**する
2. `fc.set〜Condition(priority, obj)` で登録する（`priority` は適用順序。1 始まりでシート単位に一意）
3. **全条件を登録し終えたら `fc.apply()` を呼ぶ**（必須）

## 対象クラス

### `FormatConditions`（主要メソッド）

| 条件の種類 | 設定 | 読み戻し | 条件オブジェクト |
| --- | --- | --- | --- |
| セル値（大于・小于 など） | `setCellValueCondition(p, obj)` | `getCellValueCondition(p)` | `dto.CellValueObject` |
| セル値の範囲（between） | `setCellValueRangeCondition(p, obj)` | `getCellValueRangeCondition(p)` | `dto.CellValueRangeObject` |
| カラースケール | `setColorScaleFormatCondition(p, obj)` | `getColorScaleFormatCondition(p)` | `dto.ColorScaleObject` |
| データバー | `setDataBarFormatCondition(p, obj)` | `getDataBarFormatCondition(p)` | `dto.DataBarObject` |
| 上位 / 下位 N 件 | `setTop10FormatCondition(p, obj)` | `getTop10FormatCondition(p)` | `dto.Top10Object` |
| 平均比較 | `setAboveAverageCondition(p, obj)` | `getAboveAverageCondition(p)` | `dto.AboveAverageObject` |
| アイコンセット | `setIconsetFormatCondition(p, obj)` | `getIconsetFormatCondition(p)` | `dto.IconsetConditionObject` |
| ユニーク値 | `setUniqueValuesFormatCondition(p, obj)` | `getUniqueValuesFormatCondition(p)` | `dto.UniqueValuesObject` |

| その他のメソッド | 説明 |
| --- | --- |
| `apply()` | 条件の計算・反映。**全条件設定後に必ず呼ぶ**。 |
| `getConditionTypes()` | 全ルールの一覧。`{ priority: { セル範囲: 条件タイプ } }` の入れ子マップ。 |
| `removeCondition(p)` / `removeAllCondition()` | ルールの個別削除 / 全削除。 |
| `getAppliedFontObject(range)` など | 条件を満たして実際に書式が付いたセルの取得（`apply()` 後に使う）。 |

### 条件オブジェクトの共通メソッド

`setA1C1(適用範囲)` と `setFont(dto.FontObject)` / `setPatternFill(dto.PatternFillObject)` を持ちます。 色は `dto.ColorObject` に **`#` なしの 6 桁 hex**（例 `"9C0006"`）を `setHexColor()` で設定します。

## コード

実行結果:

#### [Node.js]

```javascript
// conditional-format.js — 条件付き書式ルールの追加・読み戻し・削除
const { nodeosbxl, enums, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("condfmt.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
const color = (hex) => {
const c = new dto.ColorObject();
c.setHexColor(hex);
return c;
};
// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"].forEach((h, i) => {
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
// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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
// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);
// ---------- (3) データバー: D列 ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);
// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true); // true = 上位（false なら下位）
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
fc.apply(); // 全条件の設定後に apply() が必須
console.log("[1] 5 件のルールを登録し apply() しました");
// ---------- ルール一覧の読み戻し ----------
const typeName = Object.fromEntries(
Object.entries(enums.XlFormatConditionType).map(([name, val]) => [val, name]));
console.log(`[2] ルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
for (const [range, type] of Object.entries(ranges)) {
console.log(` priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
}
}
// ---------- 各ルールの内容の読み戻し ----------
console.log("[3] ルール内容の読み戻し");
const g1 = fc.getCellValueCondition(1);
console.log(` セル値 : ${g1.getA1C1()} 値=${g1.getConditionValue()}` +
` 演算子=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
` 文字色=${g1.getFont().getColorObject().getHexColor()} 太字=${g1.getFont().isBold()}` +
` 背景=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(` カラースケール: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, 閾値=${g2.getMediumValue()})` +
` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(` データバー: 色=${g3.getDataBarColor().getHexColor()}` +
` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(` 上位 N 件: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(` 平均比較: aboveBelow=${g5.getAboveBelow()}` +
` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);
// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
const byRow = (a, b) => Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] 書式が適用されたセル");
console.log(` B列 (>60) : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(` E列 (上位/平均): ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(` D列 データバー : ${Object.values(appliedBar).filter(Boolean).length} セル`);
// ---------- 削除 ----------
fc.removeCondition(4); // 上位 3 件のルールだけ削除
console.log(`[5] removeCondition(4) 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition(); // 全ルールを削除
console.log(`[6] removeAllCondition() 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal);
fc.apply();
wb.save();
wb.close();
console.log("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）");
```

#### [Python]

```python
# conditional-format.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
e = pyosbxl.enums
def color(hex): # ColorObject は "#" なしの 6 桁 hex で作る
c = pyosbxl.dto.ColorObject()
c.setHexColor(hex)
return c
def num(v): # JS の数値と同じ表記にする（60.0 → "60"）
return str(int(v)) if float(v) == int(v) else str(v)
def bol(v): # JS の bool と同じ表記にする（True → "true"）
return str(bool(v)).lower()
app = pyosbxl.App()
wb = app.createWorkBook("condfmt-py.xlsx") # 出力文言は node 版に合わせ condfmt.xlsx と表示
ws = wb.openWorkSheet("Sheet1")
# ---------- データ表 ----------
# B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
for i, h in enumerate(["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"]):
col = "ABCDE"[i]
ws.getRange(f"{col}1:{col}1").setValue(h)
sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49]
for i, v in enumerate(sales):
ws.getRange(f"A{i + 2}:A{i + 2}").setNumberValue(i + 1)
for col in ["B", "C", "D", "E"]:
ws.getRange(f"{col}{i + 2}:{col}{i + 2}").setNumberValue(v)
fc = ws.getFormatConditions()
# ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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
# ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
colorScale = pyosbxl.dto.ColorScaleObject()
colorScale.setA1C1("C2:C11")
colorScale.setMinimumCondition(color("F8696B"), e.XlConditionValueType.ConditionValueTypeMinimum)
colorScale.setMediumCondition(color("FFEB84"), e.XlConditionValueType.ConditionValueTypePercentile, 50)
colorScale.setMaximumCondition(color("63BE7B"), e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setColorScaleFormatCondition(2, colorScale)
# ---------- (3) データバー: D列 ----------
dataBar = pyosbxl.dto.DataBarObject()
dataBar.setA1C1("D2:D11")
dataBar.setDataBarColor(color("638EC6"))
dataBar.setMinimumCondition(e.XlConditionValueType.ConditionValueTypeMinimum)
dataBar.setMaximumCondition(e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setDataBarFormatCondition(3, dataBar)
# ---------- (4) 上位 3 件 + 平均より上: E列 ----------
top3 = pyosbxl.dto.Top10Object()
top3.setRank(3, True) # True = 上位（False なら下位）
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
fc.apply() # 全条件の設定後に apply() が必須
print("[1] 5 件のルールを登録し apply() しました")
# ---------- ルール一覧の読み戻し ----------
types = fc.getConditionTypes()
print(f"[2] ルール数 = {len(types)}")
for prio in sorted(types):
for rng, t in types[prio].items():
print(f" priority {prio}: {rng} → {t.name} (type={int(t)})")
# ---------- 各ルールの内容の読み戻し ----------
print("[3] ルール内容の読み戻し")
g1 = fc.getCellValueCondition(1)
print(f" セル値 : {g1.getA1C1()} 値={num(g1.getConditionValue())}"
f" 演算子={int(g1.getConditionOperetor())}"
f" (Greater={int(e.XlFormatConditionOperator.FormatConditionGreater)})"
f" 文字色={g1.getFont().getColorObject().getHexColor()} 太字={bol(g1.getFont().isBold())}"
f" 背景={g1.getPatternFill().getCellColor().getHexColor()}")
g2 = fc.getColorScaleFormatCondition(2)
print(f" カラースケール: min={g2.getMinimumColor().getHexColor()} (type={int(g2.getMinimumValueType())})"
f" mid={g2.getMediumColor().getHexColor()} (type={int(g2.getMediumValueType())}, 閾値={num(g2.getMediumValue())})"
f" max={g2.getMaximumColor().getHexColor()} (type={int(g2.getMaximumValueType())})")
g3 = fc.getDataBarFormatCondition(3)
print(f" データバー: 色={g3.getDataBarColor().getHexColor()}"
f" minType={int(g3.getMinimumCondition())} maxType={int(g3.getMaximumCondition())}")
g4 = fc.getTop10FormatCondition(4)
print(f" 上位 N 件: rank={g4.getRank()} isTop={bol(g4.getTopBottom())}")
g5 = fc.getAboveAverageCondition(5)
print(f" 平均比較: aboveBelow={int(g5.getAboveBelow())}"
f" (XlAboveAverage={int(e.XlAboveBelow.XlAboveAverage)})")
# ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
byRow = lambda a: int(a[1:])
print("[4] 書式が適用されたセル")
appliedB = sorted(fc.getAppliedFontObject("B2:B11").keys(), key=byRow)
print(f" B列 (>60) : {' '.join(appliedB)}")
appliedE = sorted(fc.getAppliedFontObject("E2:E11").keys(), key=byRow)
print(f" E列 (上位/平均): {' '.join(appliedE)}")
appliedBar = fc.getAppliedDataBar("D2:D11")
print(f" D列 データバー : {sum(1 for v in appliedBar.values() if v)} セル")
# ---------- 削除 ----------
fc.removeCondition(4) # 上位 3 件のルールだけ削除
print(f"[5] removeCondition(4) 後のルール数 = {len(fc.getConditionTypes())}")
fc.removeAllCondition() # 全ルールを削除
print(f"[6] removeAllCondition() 後のルール数 = {len(fc.getConditionTypes())}")
# セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal)
fc.apply()
wb.save()
wb.close()
print("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）")
```

#### [Java]

```java
// conditional-format.java — node 版サンプルの javaosbxl 対応状況を検証する
//
// 【API 差】javaosbxl 1.5.0 の条件オブジェクト Wrapper（CellValueObjectWrapper など）には
// 基底クラス相当の setA1C1() / setFont() / setPatternFill() が公開されていません。
// 適用範囲(A1C1)を設定する手段が無いため setCellValueCondition() は
// "apply range(A1C1) must be set" 例外になり、条件付き書式ルールを登録できません。
// ここでは node 版と同等のルール登録を試みて例外になること、
// apply() / getConditionTypes() / removeCondition() / removeAllCondition() が
// 例外なく呼べることを確認し、その差分を明示します。
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.Map;
public class Main {
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("condfmt-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- データ表（node 版と同一） ----------
String[] headers = {"月", "セル値ルール", "カラースケール", "データバー", "上位/平均"};
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
// ---------- (1) セル値ルール: 60 より大きい ----------
// node 版では cellVal.setA1C1("B2:B11") / setFont(...) / setPatternFill(...) を行うが、
// javaosbxl の CellValueObjectWrapper にはこれらのメソッドが無い（API 差）
CellValueObjectWrapper cellVal = new CellValueObjectWrapper();
cellVal.setValueCondition(60, XlFormatConditionOperator.FormatConditionGreater);
try {
fc.setCellValueCondition(1, cellVal);
System.out.println("[1] セル値ルールを登録できました");
} catch (RuntimeException e) {
// A1C1 未設定のため登録は必ず例外になる（javaosbxl 1.5.0 の既知の API 差）
System.out.println("[1] セル値ルール登録は例外: " + e.getMessage());
}
fc.apply(); // ルール 0 件でも apply() は例外なく呼べる
// ---------- ルール一覧の読み戻し（登録できていないので 0 件） ----------
Map<Integer, Map<String, XlFormatConditionType>> types = fc.getConditionTypes();
System.out.println("[2] ルール数 = " + types.size());
// ---------- 削除系もルール 0 件で例外なく呼べる ----------
fc.removeCondition(4);
System.out.println("[5] removeCondition(4) 後のルール数 = " + fc.getConditionTypes().size());
fc.removeAllCondition();
System.out.println("[6] removeAllCondition() 後のルール数 = " + fc.getConditionTypes().size());
wb.save();
wb.close();
System.out.println("[done] condfmt-java.xlsx を保存しました（javaosbxl は条件付き書式ルールを登録できないため 0 件）");
}
}
```

#### [PHP]

```php
// conditional-format.md — node 版サンプルと同一検証 (phposbxl)。ボディのみ (<?php・autoload は harness 付与)
// 出力ヘルパー: node 版 stdout と同じ表記にする (bool → "true"/"false")
$bol = fn($v) => $v ? "true" : "false";
// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
$color = function (string $hex): \phposbxl\dto\ColorObject {
$c = new \phposbxl\dto\ColorObject();
$c->setHexColor($hex);
return $c;
};
$app = new \phposbxl\App();
$wb = $app->createWorkBook("condfmt-php.xlsx"); // 出力文言は node 版に合わせ condfmt.xlsx と表示
$ws = $wb->openWorkSheet("Sheet1");
// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
foreach (["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"] as $i => $h) {
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
// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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
// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
$colorScale = new \phposbxl\dto\ColorScaleObject();
$colorScale->setA1C1("C2:C11");
$colorScale->setMinimumCondition($color("F8696B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$colorScale->setMediumCondition($color("FFEB84"), \phposbxl\enums\XlConditionValueType::ConditionValueTypePercentile, 50);
$colorScale->setMaximumCondition($color("63BE7B"), \phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc->setColorScaleFormatCondition(2, $colorScale);
// ---------- (3) データバー: D列 ----------
$dataBar = new \phposbxl\dto\DataBarObject();
$dataBar->setA1C1("D2:D11");
$dataBar->setDataBarColor($color("638EC6"));
$dataBar->setMinimumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMinimum);
$dataBar->setMaximumCondition(\phposbxl\enums\XlConditionValueType::ConditionValueTypeMaximum);
$fc->setDataBarFormatCondition(3, $dataBar);
// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
$top3 = new \phposbxl\dto\Top10Object();
$top3->setRank(3, true); // true = 上位（false なら下位）
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
$fc->apply(); // 全条件の設定後に apply() が必須
echo "[1] 5 件のルールを登録し apply() しました\n";
// ---------- ルール一覧の読み戻し ----------
$types = $fc->getConditionTypes(); // array<int, array<string, int>>
ksort($types); // node の Object.entries と同じく priority 昇順で走査する
echo "[2] ルール数 = " . count($types) . "\n";
foreach ($types as $prio => $ranges) {
foreach ($ranges as $range => $type) {
$name = \phposbxl\enums\XlFormatConditionType::from($type)->name;
echo " priority {$prio}: {$range} → {$name} (type={$type})\n";
}
}
// ---------- 各ルールの内容の読み戻し ----------
echo "[3] ルール内容の読み戻し\n";
$g1 = $fc->getCellValueCondition(1);
$greater = \phposbxl\enums\XlFormatConditionOperator::FormatConditionGreater->value;
echo " セル値 : {$g1->getA1C1()} 値={$g1->getConditionValue()}"
. " 演算子={$g1->getConditionOperetor()} (Greater={$greater})"
. " 文字色={$g1->getFont()->getColorObject()->getHexColor()} 太字=" . $bol($g1->getFont()->isBold())
. " 背景={$g1->getPatternFill()->getCellColor()->getHexColor()}\n";
$g2 = $fc->getColorScaleFormatCondition(2);
echo " カラースケール: min={$g2->getMinimumColor()->getHexColor()} (type={$g2->getMinimumValueType()})"
. " mid={$g2->getMediumColor()->getHexColor()} (type={$g2->getMediumValueType()}, 閾値={$g2->getMediumValue()})"
. " max={$g2->getMaximumColor()->getHexColor()} (type={$g2->getMaximumValueType()})\n";
$g3 = $fc->getDataBarFormatCondition(3);
echo " データバー: 色={$g3->getDataBarColor()->getHexColor()}"
. " minType={$g3->getMinimumCondition()} maxType={$g3->getMaximumCondition()}\n";
$g4 = $fc->getTop10FormatCondition(4);
echo " 上位 N 件: rank={$g4->getRank()} isTop=" . $bol($g4->getTopBottom()) . "\n";
$g5 = $fc->getAboveAverageCondition(5);
$xlAbove = \phposbxl\enums\XlAboveBelow::XlAboveAverage->value;
echo " 平均比較: aboveBelow={$g5->getAboveBelow()} (XlAboveAverage={$xlAbove})\n";
// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
$byRow = fn($a, $b) => (int)substr($a, 1) - (int)substr($b, 1);
echo "[4] 書式が適用されたセル\n";
$appliedB = array\_keys($fc->getAppliedFontObject("B2:B11"));
usort($appliedB, $byRow);
echo " B列 (>60) : " . implode(" ", $appliedB) . "\n";
$appliedE = array\_keys($fc->getAppliedFontObject("E2:E11"));
usort($appliedE, $byRow);
echo " E列 (上位/平均): " . implode(" ", $appliedE) . "\n";
$appliedBar = $fc->getAppliedDataBar("D2:D11");
echo " D列 データバー : " . count(array\_filter($appliedBar)) . " セル\n";
// ---------- 削除 ----------
$fc->removeCondition(4); // 上位 3 件のルールだけ削除
echo "[5] removeCondition(4) 後のルール数 = " . count($fc->getConditionTypes()) . "\n";
$fc->removeAllCondition(); // 全ルールを削除
echo "[6] removeAllCondition() 後のルール数 = " . count($fc->getConditionTypes()) . "\n";
// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
$fc->setCellValueCondition(1, $cellVal);
$fc->apply();
$wb->save();
$wb->close();
echo "[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）\n";
```

#### [WebAssembly]

```javascript
// conditional-format.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("condfmt-wasm.xlsx"); // 出力文言は node 版と同一（condfmt.xlsx と表示）
const ws = wb.openWorkSheet("Sheet1");
// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
const color = (hex) => {
const c = new dto.ColorObject();
c.setHexColor(hex);
return c;
};
// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"].forEach((h, i) => {
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
// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
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
// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);
// ---------- (3) データバー: D列 ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);
// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true); // true = 上位（false なら下位）
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
fc.apply(); // 全条件の設定後に apply() が必須
console.log("[1] 5 件のルールを登録し apply() しました");
// ---------- ルール一覧の読み戻し ----------
const typeName = Object.fromEntries(
Object.entries(enums.XlFormatConditionType).map(([name, val]) => [val, name]));
console.log(`[2] ルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
for (const [range, type] of Object.entries(ranges)) {
console.log(` priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
}
}
// ---------- 各ルールの内容の読み戻し ----------
console.log("[3] ルール内容の読み戻し");
const g1 = fc.getCellValueCondition(1);
console.log(` セル値 : ${g1.getA1C1()} 値=${g1.getConditionValue()}` +
` 演算子=${g1.getConditionOperetor()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
` 文字色=${g1.getFont().getColorObject().getHexColor()} 太字=${g1.getFont().isBold()}` +
` 背景=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(` カラースケール: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, 閾値=${g2.getMediumValue()})` +
` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(` データバー: 色=${g3.getDataBarColor().getHexColor()}` +
` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(` 上位 N 件: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(` 平均比較: aboveBelow=${g5.getAboveBelow()}` +
` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);
// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
const byRow = (a, b) => Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] 書式が適用されたセル");
console.log(` B列 (>60) : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(` E列 (上位/平均): ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(` D列 データバー : ${Object.values(appliedBar).filter(Boolean).length} セル`);
// ---------- 削除 ----------
fc.removeCondition(4); // 上位 3 件のルールだけ削除
console.log(`[5] removeCondition(4) 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition(); // 全ルールを削除
console.log(`[6] removeAllCondition() 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal);
fc.apply();
wb.save();
wb.close();
console.log("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）");
```

```plaintext
[1] 5 件のルールを登録し apply() しました
[2] ルール数 = 5
priority 1: B2:B11 → CellValue (type=1)
priority 2: C2:C11 → ColorScale (type=3)
priority 3: D2:D11 → Databar (type=4)
priority 4: E2:E11 → Top10 (type=5)
priority 5: E2:E11 → AboveAverageCondition (type=12)
[3] ルール内容の読み戻し
セル値 : B2:B11 値=60 演算子=5 (Greater=5) 文字色=9C0006 太字=true 背景=FFC7CE
カラースケール: min=F8696B (type=5) mid=FFEB84 (type=4, 閾値=50) max=63BE7B (type=6)
データバー: 色=638EC6 minType=5 maxType=6
上位 N 件: rank=3 isTop=true
平均比較: aboveBelow=0 (XlAboveAverage=0)
[4] 書式が適用されたセル
B列 (>60) : B3 B5 B8 B10
E列 (上位/平均): E3 E5 E6 E8 E10
D列 データバー : 10 セル
[5] removeCondition(4) 後のルール数 = 4
[6] removeAllCondition() 後のルール数 = 0
[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）
```

B 列の `B3 B5 B8 B10` は値が 78・91・87・65（= 60 より大きい）のセル、 E 列の `E3 E5 E6 E8 E10` は「上位 3 件」と「平均（54.2）より上」の和集合です。 見た目の書式は検証できないため、このように**ルールの個数・種類・しきい値・適用されたセル**を 読み戻して出力することで、意図どおり設定できたことを確認しています。

## 解説

### priority（適用順序）

* ルールは `priority`（1 始まりの適用順序）で管理され、**シート単位で一意**です。  
   `getConditionTypes()` の返却値のキーがこの priority です。
* **既存の priority に登録すると、既存ルールが後ろへずれます**（実測）。  
   たとえば priority 1 に 2 回登録すると、先に登録したルールは priority 2 になります。
* `removeCondition(p)` で削除すると、**後ろの priority が繰り上がります**。
* 読み戻しメソッド（`getCellValueCondition(p)` など）は **priority の位置に別の種類のルールがあると例外**になります。  
   まず `getConditionTypes()` で種類を確認してから取得してください。

### apply() の役割

* **`apply()` を呼ぶまで条件は評価されません。** 実測では、`apply()` 前の `getAppliedFontObject()` は空 `{}` を返し、 `apply()` 後は条件を満たすセル（例では B3 B5 B8 B10）が返りました。
* ルールを追加・削除したら、その都度 `apply()` を呼んでください。

### 主な enum

| enum | よく使う値 |
| --- | --- |
| `XlFormatConditionOperator` | `FormatConditionGreater`(5) / `FormatConditionGreaterEqual`(7) / `FormatConditionLess`(6) / `FormatConditionLessEqual`(8) / `FormatConditionEqual`(3) / `FormatConditionBetween`(1) |
| `XlFormatConditionType` | `CellValue`(1) / `ColorScale`(3) / `Databar`(4) / `Top10`(5) / `XlIconSet`(6) / `UniqueValues`(8) / `AboveAverageCondition`(12) |
| `XlConditionValueType` | `ConditionValueTypeMinimum`(5) / `ConditionValueTypeMaximum`(6) / `ConditionValueTypePercentile`(4) / `ConditionValueTypeNumeric`(1) |
| `XlAboveBelow` | `XlAboveAverage`(0) / `XlBelowAverage`(2) / `XlAboveStdDev`(1) / `XlBelowStdDev`(3) |

### 永続化

条件付き書式は `wb.save()` で .xlsx に保存され、`app.openWorkBook()` で開き直しても `getConditionTypes()` や各読み戻しメソッドで復元できます（実測済み）。

## 注意点

* **`apply()` を忘れない。** 忘れても例外にはならず、`getApplied*` 系が空を返すだけなので 気づきにくいのが厄介です（[解説](#apply-の役割)参照）。
* **演算子の getter 名は `getConditionOperetor()` です。** `Operator` ではなく **`Operetor`**（ライブラリ側の綴りそのまま）なので注意してください。
* **データバーの読み戻しはメソッド名がカラースケールと違います。** `DataBarObject` では `getMinimumCondition()` / `getMaximumCondition()` が **`XlConditionValueType`（評価条件の種類）**を返し、しきい値は `getMinimumValue()` / `getMaximumValue()`。  
   一方 `ColorScaleObject` では種類が `getMinimumValueType()`、色が `getMinimumColor()` です。
* **色は `#` なしの 6 桁 hex**（例 `"9C0006"`）で `setHexColor()` に渡します。
* **`Top10Object.setRank(rank, isTop)` の第2引数は「上位か下位か」の bool** です （`true` = 上位 N 件、`false` = 下位 N 件）。  
  「Top10」という名前ですが N 件は任意です。
* **条件オブジェクトには `setA1C1()` で適用範囲を必ず設定します。** 適用範囲は条件オブジェクト側が持ち、`set〜Condition()` には範囲を渡しません。
* \*\*`getApplied*` 系が返すのは「条件を満たして書式が実際に付いたセル」だけ\*\*です。  
   ルールの適用範囲全体ではありません。

## 関連

* [目次](index.md)
* 前: [テーブルとコメント](general-08-tables-comments.md) / 次: [チャート](general-10-charts.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
