# 行・列・セル範囲の操作 — 挿入・削除・クリア・コピー

level: basic ／ verified: 2026-10-04 ／ 5言語タブ対応 ／ [English](../en/general-05-rows-cols.md)

## やりたいこと

ワークシート上の行・列・セル範囲を、挿入 / 削除 / クリア / コピーする。 あわせて行の高さ・列の幅・非表示などを `Row` / `Col` オブジェクトで設定する。

## 4 つの操作の違い

| 操作 | 挙動 | 行・列の場合 | セル範囲の場合 |
| --- | --- | --- | --- |
| **insert** | 空白を挿入し、既存データを**ずらす** | `insertRow(start, n)` / `insertCol(start, n)` | `insertCell(A1C1, direction)` |
| **delete** | 削除して、既存データを**ずらす** | `deleteRow(start, n)` / `deleteCol(start, n)` | `deleteCell(A1C1, direction)` |
| **clear** | 中身と書式を初期状態に戻す。**ずらさない** | `clearRow(start, n)` / `clearCol(start, n)` | `clearCell(A1C1)` |
| **copy** | 別シート（または自シート）から**上書きコピー** | `copyRow(fromSheet, fromRow, n, toRow, copyType?)` | `copyCell(fromSheet, fromA1C1, toA1, copyType?)` |

`insertCell` / `deleteCell` の `direction` は `enums.XlInsertDirection` を使います。

| 値 | 数値 | 用途 |
| --- | --- | --- |
| `InsertDirectionRight` | 1 | セル挿入時に右へずらす |
| `InsertDirectionDown` | 2 | セル挿入時に下へずらす |
| `DeleteDirectionLeft` | 3 | セル削除時に左へずらす |
| `DeleteDirectionUp` | 4 | セル削除時に上へずらす |

## コード 1: 行・列・セル範囲の操作

実行結果:

#### [Node.js]

```javascript
// rows-cols.js — 行・列・セル範囲の挿入/削除/クリア/コピー
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("rows-cols.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
const dump = (label, range) => {
const v = ws.getRange(range).getValue();
const keys = Object.keys(v).map((k) => /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =>
Number(a[2]) - Number(b[2]) ||
app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
console.log(`${label} ${keys.map((m) => `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// A1:C5 を r{行}c{列} で埋める
const fill = () => {
for (let r = 1; r <= 5; r++) for (let c = 1; c <= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};
fill();
dump("[初期] A列:", "A1:A5");
// ---------- 行 ----------
ws.insertRow(3, 2); // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow] A列:", "A1:A7");
ws.deleteRow(3, 2); // 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow] A列:", "A1:A5");
ws.clearRow(2, 1); // 2 行目をクリア（位置はずらさない）
dump("[clearRow] A列:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4); // 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow] A列:", "A1:A5");
fill();
// ---------- 列 ----------
ws.insertCol(2, 1); // B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1");
ws.deleteCol(2, 1); // 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1");
ws.clearCol(3, 1); // C 列をクリア
dump("[clearCol] 1行目:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5); // A,B 列を E,F 列へコピー
dump("[copyCol] 1行目:", "A1:F1");
// ---------- セル範囲 ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] B列:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] B列:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell] F8:G9:", "F8:G9");
wb.save();
wb.close();
```

#### [Python]

```python
# rows-cols.py — node 版サンプルと同一検証 (pyosbxl)
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
dump("[初期] A列:", "A1:A5")
# ---------- 行 ----------
ws.insertRow(3, 2) # 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow] A列:", "A1:A7")
ws.deleteRow(3, 2) # 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow] A列:", "A1:A5")
ws.clearRow(2, 1) # 2 行目をクリア（位置はずらさない）
dump("[clearRow] A列:", "A1:A5")
ws.copyRow("Sheet1", 1, 1, 4) # 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow] A列:", "A1:A5")
fill()
# ---------- 列 ----------
ws.insertCol(2, 1) # B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1")
ws.deleteCol(2, 1) # 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1")
ws.clearCol(3, 1) # C 列をクリア
dump("[clearCol] 1行目:", "A1:C1")
fill()
ws.copyCol("Sheet1", 1, 2, 5) # A,B 列を E,F 列へコピー
dump("[copyCol] 1行目:", "A1:F1")
# ---------- セル範囲 ----------
ws.insertCell("B2:B3", pyosbxl.enums.XlInsertDirection.InsertDirectionDown)
dump("[insertCell] B列:", "B1:B5")
ws.deleteCell("B2:B3", pyosbxl.enums.XlInsertDirection.DeleteDirectionUp)
dump("[deleteCell] B列:", "B1:B5")
ws.clearCell("A1:B2")
dump("[clearCell] A1:B2:", "A1:B2")
ws.copyCell("Sheet1", "A3:B4", "F8")
dump("[copyCell] F8:G9:", "F8:G9")
wb.save()
wb.close()
```

#### [Java]

```java
// rows-cols.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.enums.\*;
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
Map<String, String> v = ws.getRange(range).getValue();
List<String> keys = new ArrayList<>(v.keySet());
keys.sort((a, b) -> {
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
if (sb.length() > 0) sb.append(" ");
String val = v.get(k);
sb.append(k).append("=").append(val == null || val.isEmpty() ? "·" : val);
}
System.out.println(label + " " + sb);
}
// A1:C5 を r{行}c{列} で埋める
static void fill() {
for (int r = 1; r <= 5; r++)
for (int c = 1; c <= 3; c++)
ws.getCells(r, c).setValue("r" + r + "c" + c);
}
public static void main(String[] args) {
app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("rows-cols-java.xlsx");
ws = wb.openWorkSheet("Sheet1");
fill();
dump("[初期] A列:", "A1:A5");
// ---------- 行 ----------
ws.insertRow(3, 2); // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow] A列:", "A1:A7");
ws.deleteRow(3, 2); // 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow] A列:", "A1:A5");
ws.clearRow(2, 1); // 2 行目をクリア（位置はずらさない）
dump("[clearRow] A列:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4); // 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow] A列:", "A1:A5");
fill();
// ---------- 列 ----------
ws.insertCol(2, 1); // B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1");
ws.deleteCol(2, 1); // 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1");
ws.clearCol(3, 1); // C 列をクリア
dump("[clearCol] 1行目:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5); // A,B 列を E,F 列へコピー
dump("[copyCol] 1行目:", "A1:F1");
// ---------- セル範囲 ----------
ws.insertCell("B2:B3", XlInsertDirection.InsertDirectionDown);
dump("[insertCell] B列:", "B1:B5");
ws.deleteCell("B2:B3", XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] B列:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell] F8:G9:", "F8:G9");
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// rows-cols.md — 行・列・セル範囲の挿入/削除/クリア/コピー (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("rows-cols-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
$dump = function (string $label, string $range) use ($ws, $app): void {
$v = $ws->getRange($range)->getValue();
$keys = array\_keys($v);
usort($keys, function (string $a, string $b) use ($app): int {
preg\_match('/^([A-Z]+)(\d+)$/', $a, $ma);
preg\_match('/^([A-Z]+)(\d+)$/', $b, $mb);
return ((int)$ma[2] <=> (int)$mb[2])
?: ($app->convertToColumnNumber($ma[1]) <=> $app->convertToColumnNumber($mb[1]));
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
for ($r = 1; $r <= 5; $r++) {
for ($c = 1; $c <= 3; $c++) {
$ws->getCells($r, $c)->setValue("r{$r}c{$c}");
}
}
};
$fill();
$dump("[初期] A列:", "A1:A5");
// ---------- 行 ----------
$ws->insertRow(3, 2); // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
$dump("[insertRow] A列:", "A1:A7");
$ws->deleteRow(3, 2); // 3 行目から 2 行削除 → 下の行が上がる
$dump("[deleteRow] A列:", "A1:A5");
$ws->clearRow(2, 1); // 2 行目をクリア（位置はずらさない）
$dump("[clearRow] A列:", "A1:A5");
$ws->copyRow("Sheet1", 1, 1, 4); // 自シートの 1 行目を 4 行目へ上書きコピー
$dump("[copyRow] A列:", "A1:A5");
$fill();
// ---------- 列 ----------
$ws->insertCol(2, 1); // B 列の前に 1 列挿入
$dump("[insertCol] 1行目:", "A1:D1");
$ws->deleteCol(2, 1); // 挿入した B 列を削除
$dump("[deleteCol] 1行目:", "A1:C1");
$ws->clearCol(3, 1); // C 列をクリア
$dump("[clearCol] 1行目:", "A1:C1");
$fill();
$ws->copyCol("Sheet1", 1, 2, 5); // A,B 列を E,F 列へコピー
$dump("[copyCol] 1行目:", "A1:F1");
// ---------- セル範囲 ----------
$ws->insertCell("B2:B3", \phposbxl\enums\XlInsertDirection::InsertDirectionDown);
$dump("[insertCell] B列:", "B1:B5");
$ws->deleteCell("B2:B3", \phposbxl\enums\XlInsertDirection::DeleteDirectionUp);
$dump("[deleteCell] B列:", "B1:B5");
$ws->clearCell("A1:B2");
$dump("[clearCell] A1:B2:", "A1:B2");
$ws->copyCell("Sheet1", "A3:B4", "F8");
$dump("[copyCell] F8:G9:", "F8:G9");
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// rows-cols.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("rows-cols-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// 行 → 列 の順に並べて 1 行で表示する（· は空セル）
const dump = (label, range) => {
const v = ws.getRange(range).getValue();
const keys = Object.keys(v).map((k) => /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =>
Number(a[2]) - Number(b[2]) ||
app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
console.log(`${label} ${keys.map((m) => `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// A1:C5 を r{行}c{列} で埋める
const fill = () => {
for (let r = 1; r <= 5; r++) for (let c = 1; c <= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};
fill();
dump("[初期] A列:", "A1:A5");
// ---------- 行 ----------
ws.insertRow(3, 2); // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow] A列:", "A1:A7");
ws.deleteRow(3, 2); // 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow] A列:", "A1:A5");
ws.clearRow(2, 1); // 2 行目をクリア（位置はずらさない）
dump("[clearRow] A列:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4); // 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow] A列:", "A1:A5");
fill();
// ---------- 列 ----------
ws.insertCol(2, 1); // B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1");
ws.deleteCol(2, 1); // 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1");
ws.clearCol(3, 1); // C 列をクリア
dump("[clearCol] 1行目:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5); // A,B 列を E,F 列へコピー
dump("[copyCol] 1行目:", "A1:F1");
// ---------- セル範囲 ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] B列:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] B列:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell] F8:G9:", "F8:G9");
wb.save();
wb.close();
```

```plaintext
[初期] A列: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[insertRow] A列: A1=r1c1 A2=r2c1 A3=· A4=· A5=r3c1 A6=r4c1 A7=r5c1
[deleteRow] A列: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[clearRow] A列: A1=r1c1 A2=· A3=r3c1 A4=r4c1 A5=r5c1
[copyRow] A列: A1=r1c1 A2=· A3=r3c1 A4=r1c1 A5=r5c1
[insertCol] 1行目: A1=r1c1 B1=· C1=r1c2 D1=r1c3
[deleteCol] 1行目: A1=r1c1 B1=r1c2 C1=r1c3
[clearCol] 1行目: A1=r1c1 B1=r1c2 C1=·
[copyCol] 1行目: A1=r1c1 B1=r1c2 C1=r1c3 D1=· E1=r1c1 F1=r1c2
[insertCell] B列: B1=r1c2 B2=· B3=· B4=r2c2 B5=r3c2
[deleteCell] B列: B1=r1c2 B2=r2c2 B3=r3c2 B4=r4c2 B5=r5c2
[clearCell] A1:B2: A1=· B1=· A2=· B2=·
[copyCell] F8:G9: F8=r3c1 G8=r3c2 F9=r4c1 G9=r4c2
```

`·` は空セルです。`insertRow` で 3・4 行目が空白になり既存データが下へずれたこと、 `clearRow` では位置がずれていないこと、`copyRow` で 1 行目の内容が 4 行目に上書きされたことが読み取れます。

## コード 2: `Row` / `Col` オブジェクト

実行結果:

#### [Node.js]

```javascript
// row-col-props.js — 行の高さ・列の幅・非表示・自動調整
const { nodeosbxl } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("row-col-props.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- Row ----------
const row = ws.getRow(1); // 1 行目
console.log("既定の高さ =", row.getHeight());
row.setHeight(30);
console.log("setHeight(30)後 =", row.getHeight());
console.log("getHidden 既定 =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true); // 2 行目を非表示
console.log("2行目 hidden =", ws.getRow(2).getHidden());
// ---------- Col ----------
const col = ws.getCol(1); // 1 列目（A 列）。列番号は数値で渡す
console.log("既定の列幅 = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -> columnWidth=" + col.getColumnWidth(),
"/ width=" + col.getWidth());
ws.getCol(2).setHidden(true); // B 列を非表示
console.log("B列 hidden =", ws.getCol(2).getHidden());
ws.getRange("C1:C1").setValue("自動調整のサンプル文字列");
console.log("C列 bestFit 既定 =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true); // 内容に合わせて列幅を自動調整
console.log("setBestFit(true) -> bestFit =", ws.getCol(3).getBestFit(),
"/ columnWidth =", ws.getCol(3).getColumnWidth());
// アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1); // 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1); // B〜C 列をレベル 1 でグループ化
console.log("4行目 outlineLevel =", ws.getRow(4).getOutlineLevel(),
"/ B列 outlineLevel =", ws.getCol(2).getOutlineLevel());
wb.save();
wb.close();
```

#### [Python]

```python
# row-col-props.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("row-col-props-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
def num(x): # 整数値の float は int 表記にする（JS の出力と揃える）
return int(x) if isinstance(x, float) and x.is\_integer() else x
def boolstr(x): # Python の True/False を JS 風の true/false で表示する
return str(x).lower()
# ---------- Row ----------
row = ws.getRow(1) # 1 行目
print("既定の高さ =", num(row.getHeight()))
row.setHeight(30)
print("setHeight(30)後 =", num(row.getHeight()))
print("getHidden 既定 =", boolstr(row.getHidden()), "/ getOutlineLevel =", row.getOutlineLevel())
ws.getRow(2).setHidden(True) # 2 行目を非表示
print("2行目 hidden =", boolstr(ws.getRow(2).getHidden()))
# ---------- Col ----------
col = ws.getCol(1) # 1 列目（A 列）。列番号は数値で渡す
print("既定の列幅 = columnWidth:", col.getColumnWidth(), "/ width:", num(col.getWidth()))
col.setColumnWidth(20)
print("setColumnWidth(20) -> columnWidth=" + str(num(col.getColumnWidth())),
"/ width=" + str(num(col.getWidth())))
ws.getCol(2).setHidden(True) # B 列を非表示
print("B列 hidden =", boolstr(ws.getCol(2).getHidden()))
ws.getRange("C1:C1").setValue("自動調整のサンプル文字列")
print("C列 bestFit 既定 =", boolstr(ws.getCol(3).getBestFit()))
ws.getCol(3).setBestFit(True) # 内容に合わせて列幅を自動調整
print("setBestFit(true) -> bestFit =", boolstr(ws.getCol(3).getBestFit()),
"/ columnWidth =", ws.getCol(3).getColumnWidth())
# アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1) # 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1) # B〜C 列をレベル 1 でグループ化
print("4行目 outlineLevel =", ws.getRow(4).getOutlineLevel(),
"/ B列 outlineLevel =", ws.getCol(2).getOutlineLevel())
wb.save()
wb.close()
```

#### [Java]

```java
// row-col-props.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.\*;
public class Main {
// 整数値の double は int 表記にする（JS の出力と揃える）
static String num(double d) {
if (d == Math.rint(d) && !Double.isInfinite(d)) return String.valueOf((long) d);
return String.valueOf(d);
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("row-col-props-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- Row ----------
RowWrapper row = ws.getRow(1); // 1 行目
System.out.println("既定の高さ = " + num(row.getHeight()));
row.setHeight(30);
System.out.println("setHeight(30)後 = " + num(row.getHeight()));
System.out.println("getHidden 既定 = " + row.getHidden()
+ " / getOutlineLevel = " + row.getOutlineLevel());
ws.getRow(2).setHidden(true); // 2 行目を非表示
System.out.println("2行目 hidden = " + ws.getRow(2).getHidden());
// ---------- Col ----------
ColWrapper col = ws.getCol(1); // 1 列目（A 列）。列番号は数値で渡す
System.out.println("既定の列幅 = columnWidth: " + num(col.getColumnWidth())
+ " / width: " + num(col.getWidth()));
col.setColumnWidth(20);
System.out.println("setColumnWidth(20) -> columnWidth=" + num(col.getColumnWidth())
+ " / width=" + num(col.getWidth()));
ws.getCol(2).setHidden(true); // B 列を非表示
System.out.println("B列 hidden = " + ws.getCol(2).getHidden());
ws.getRange("C1:C1").setValue("自動調整のサンプル文字列");
System.out.println("C列 bestFit 既定 = " + ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true); // 内容に合わせて列幅を自動調整
System.out.println("setBestFit(true) -> bestFit = " + ws.getCol(3).getBestFit()
+ " / columnWidth = " + num(ws.getCol(3).getColumnWidth()));
// アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1); // 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1); // B〜C 列をレベル 1 でグループ化
System.out.println("4行目 outlineLevel = " + ws.getRow(4).getOutlineLevel()
+ " / B列 outlineLevel = " + ws.getCol(2).getOutlineLevel());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// row-col-props.md — 行の高さ・列の幅・非表示・自動調整 (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("row-col-props-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// PHP の bool は echo すると 1/空文字になるため JS 風の true/false に変換する
$boolstr = fn(bool $b): string => $b ? "true" : "false";
// ---------- Row ----------
$row = $ws->getRow(1); // 1 行目
echo "既定の高さ = ", $row->getHeight(), "\n";
$row->setHeight(30);
echo "setHeight(30)後 = ", $row->getHeight(), "\n";
echo "getHidden 既定 = ", $boolstr($row->getHidden()), " / getOutlineLevel = ", $row->getOutlineLevel(), "\n";
$ws->getRow(2)->setHidden(true); // 2 行目を非表示
echo "2行目 hidden = ", $boolstr($ws->getRow(2)->getHidden()), "\n";
// ---------- Col ----------
$col = $ws->getCol(1); // 1 列目（A 列）。列番号は数値で渡す
echo "既定の列幅 = columnWidth: ", $col->getColumnWidth(), " / width: ", $col->getWidth(), "\n";
$col->setColumnWidth(20);
echo "setColumnWidth(20) -> columnWidth=", $col->getColumnWidth(),
" / width=", $col->getWidth(), "\n";
$ws->getCol(2)->setHidden(true); // B 列を非表示
echo "B列 hidden = ", $boolstr($ws->getCol(2)->getHidden()), "\n";
$ws->getRange("C1:C1")->setValue("自動調整のサンプル文字列");
echo "C列 bestFit 既定 = ", $boolstr($ws->getCol(3)->getBestFit()), "\n";
$ws->getCol(3)->setBestFit(true); // 内容に合わせて列幅を自動調整
echo "setBestFit(true) -> bestFit = ", $boolstr($ws->getCol(3)->getBestFit()),
" / columnWidth = ", $ws->getCol(3)->getColumnWidth(), "\n";
// アウトライン（グループ化）は WorkSheet 側で設定する
$ws->setRowOutline(4, 6, 1); // 4〜6 行目をレベル 1 でグループ化
$ws->setColOutline(2, 3, 1); // B〜C 列をレベル 1 でグループ化
echo "4行目 outlineLevel = ", $ws->getRow(4)->getOutlineLevel(),
" / B列 outlineLevel = ", $ws->getCol(2)->getOutlineLevel(), "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// row-col-props.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("row-col-props-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- Row ----------
const row = ws.getRow(1); // 1 行目
console.log("既定の高さ =", row.getHeight());
row.setHeight(30);
console.log("setHeight(30)後 =", row.getHeight());
console.log("getHidden 既定 =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true); // 2 行目を非表示
console.log("2行目 hidden =", ws.getRow(2).getHidden());
// ---------- Col ----------
const col = ws.getCol(1); // 1 列目（A 列）。列番号は数値で渡す
console.log("既定の列幅 = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -> columnWidth=" + col.getColumnWidth(),
"/ width=" + col.getWidth());
ws.getCol(2).setHidden(true); // B 列を非表示
console.log("B列 hidden =", ws.getCol(2).getHidden());
ws.getRange("C1:C1").setValue("自動調整のサンプル文字列");
console.log("C列 bestFit 既定 =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true); // 内容に合わせて列幅を自動調整
console.log("setBestFit(true) -> bestFit =", ws.getCol(3).getBestFit(),
"/ columnWidth =", ws.getCol(3).getColumnWidth());
// アウトライン（グループ化）は WorkSheet 側で設定する
ws.setRowOutline(4, 6, 1); // 4〜6 行目をレベル 1 でグループ化
ws.setColOutline(2, 3, 1); // B〜C 列をレベル 1 でグループ化
console.log("4行目 outlineLevel =", ws.getRow(4).getOutlineLevel(),
"/ B列 outlineLevel =", ws.getCol(2).getOutlineLevel());
wb.save();
wb.close();
```

```plaintext
既定の高さ = 18.75
setHeight(30)後 = 30
getHidden 既定 = false / getOutlineLevel = 0
2行目 hidden = true
既定の列幅 = columnWidth: 8.38 / width: 72
setColumnWidth(20) -> columnWidth=20 / width=165
B列 hidden = true
C列 bestFit 既定 = false
setBestFit(true) -> bestFit = true / columnWidth = 25.38
4行目 outlineLevel = 1 / B列 outlineLevel = 1
```

`setBestFit(true)` による列幅は、セルの内容とブックの既定フォントから算出されるため 環境によって値が変わります（上記は 25.38）。

## 解説

* **行番号・列番号はすべて 1 始まりの数値。** 列は `"A"` ではなく `1` で渡します。  
   文字列から変換する場合は `app.convertToColumnNumber("B")` → `2` を使います。
* **`insert` / `delete` はデータをずらし、`clear` はずらしません。** 行を空にしたいだけなら `clearRow()`（既存データの位置は保たれる）、 行そのものを無くしたいなら `deleteRow()` を使います。
* **`copy` 系はコピー元を「シート名」で指定します。** 自シート内コピーもシート名を渡します。  
   第5引数 `copyType` に `enums.XlCopyContentType` を渡すとコピー対象を絞れます。  
   | 値 | 数値 | コピー対象 | |---|---|---| | `CopyContentTypeAll` | 0 | 値、セル書式、関数セルのセル書式（既定） | | `CopyContentTypeValueOnly` | 1 | 値のみ | | `CopyContentTypeValueAndFormat` | 2 | 値とセル書式 | | `CopyContentTypeFormatOnly` | 3 | セル書式のみ |
* **`Col.getWidth()` と `Col.getColumnWidth()` は単位が違います。** 実測で、既定状態の A 列は `getWidth()` = 72、`getColumnWidth()` = 8.38 でした。  
   `setColumnWidth(20)` の後は `getWidth()` = 165、`getColumnWidth()` = 20 です。  
   **設定できるのは `setColumnWidth()`（文字幅）だけ**で、`getWidth()` は読み取り専用です。
* **`setBestFit(true)` は設定時のセル内容に合わせて列幅を再計算します。** 後に内容を書き換えても自動では追従しないので、データを入れ終えてから呼んでください。
* **アウトライン（グループ化）は `WorkSheet` のメソッド。** `setRowOutline(startRow, endRow, level)` / `setColOutline(startCol, endCol, level)` で設定し、 `clearRowOutline(level)` / `clearColOutline(level)` で解除します。  
   結果は `Row.getOutlineLevel()` / `Col.getOutlineLevel()` で確認できます。

## 注意点

* **`insertCell` / `deleteCell` の `direction` は操作に合わせて使い分けます。** 挿入は `InsertDirectionRight`(1) / `InsertDirectionDown`(2)、 削除は `DeleteDirectionLeft`(3) / `DeleteDirectionUp`(4) が本来の指定です。
* **`copy` 系でコピーされないものがあります。** コピー対象を `CopyContentTypeAll` にしても、 **図形（グラフを含む）・ピボットテーブル・queryTable 形式のテーブル・データテーブルは コピーされません。**
* **`copyCell` の第3引数は「コピー先の左上セル」1 つだけ。** `copyCell("Sheet1", "A3:B4", "F8")` は F8:G9 に展開されます（範囲の大きさはコピー元から決まる）。
* **`copyRow` / `copyCol` のコピー元はシート名、位置は行番号・列番号（数値）。** セル範囲の文字列（A1 形式）は使いません。  
  混同しやすいので注意してください。

## 関連

* [目次](index.md)
* 前: [セル値の型と表示書式](general-04-value-types.md) / 次: [大量データの一括投入](general-06-bulk-input.md)
* API リファレンス

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
