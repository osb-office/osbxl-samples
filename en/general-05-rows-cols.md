# Rows, Columns and Ranges — insert, delete, copy

level: basic ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](/ja/general-05-rows-cols.md)

## What you want to do

Insert / delete / clear / copy rows, columns and cell ranges on a worksheet. Also set row height, column width, hidden state and more with the `Row` / `Col` objects.

## The four operations compared

| Operation | Behavior | For rows / columns | For cell ranges |
| --- | --- | --- | --- |
| **insert** | Inserts blanks and **shifts** existing data | `insertRow(start, n)` / `insertCol(start, n)` | `insertCell(A1C1, direction)` |
| **delete** | Deletes and **shifts** existing data | `deleteRow(start, n)` / `deleteCol(start, n)` | `deleteCell(A1C1, direction)` |
| **clear** | Resets contents and formatting to the initial state. **Does not shift** | `clearRow(start, n)` / `clearCol(start, n)` | `clearCell(A1C1)` |
| **copy** | **Overwrite-copies** from another sheet (or the same sheet) | `copyRow(fromSheet, fromRow, n, toRow, copyType?)` | `copyCell(fromSheet, fromA1C1, toA1, copyType?)` |

The `direction` of `insertCell` / `deleteCell` uses `enums.XlInsertDirection`.

| Value | Number | Purpose |
| --- | --- | --- |
| `InsertDirectionRight` | 1 | Shift cells right when inserting |
| `InsertDirectionDown` | 2 | Shift cells down when inserting |
| `DeleteDirectionLeft` | 3 | Shift cells left when deleting |
| `DeleteDirectionUp` | 4 | Shift cells up when deleting |

## Code 1: Rows, columns and cell ranges

Output:

#### [Node.js]

```javascript
// rows-cols.js — insert/delete/clear/copy for rows, columns and cell ranges
const { nodeosbxl, enums } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("rows-cols.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// Print a range on one line, ordered by row → column (· marks an empty cell)
const dump = (label, range) => {
const v = ws.getRange(range).getValue();
const keys = Object.keys(v).map((k) => /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =>
Number(a[2]) - Number(b[2]) ||
app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
console.log(`${label} ${keys.map((m) => `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// Fill A1:C5 with r{row}c{col}
const fill = () => {
for (let r = 1; r <= 5; r++) for (let c = 1; c <= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};
fill();
dump("[initial] col A:", "A1:A5");
// ---------- rows ----------
ws.insertRow(3, 2); // insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow] col A:", "A1:A7");
ws.deleteRow(3, 2); // delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow] col A:", "A1:A5");
ws.clearRow(2, 1); // clear row 2 (positions do not shift)
dump("[clearRow] col A:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4); // overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow] col A:", "A1:A5");
fill();
// ---------- columns ----------
ws.insertCol(2, 1); // insert 1 column before column B
dump("[insertCol] row 1:", "A1:D1");
ws.deleteCol(2, 1); // delete the column B just inserted
dump("[deleteCol] row 1:", "A1:C1");
ws.clearCol(3, 1); // clear column C
dump("[clearCol] row 1:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5); // copy columns A,B onto columns E,F
dump("[copyCol] row 1:", "A1:F1");
// ---------- cell ranges ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] col B:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] col B:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell] F8:G9:", "F8:G9");
wb.save();
wb.close();
```

#### [Python]

```python
# rows-cols.py — same verification as the Node.js sample (pyosbxl)
import re
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("rows-cols-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# Print a range on one line, ordered by row → column (· marks an empty cell)
def dump(label, rng):
v = ws.getRange(rng).getValue()
def keyf(k):
m = re.match(r"^([A-Z]+)(\d+)$", k)
return (int(m.group(2)), app.convertToColumnNumber(m.group(1)))
keys = sorted(v.keys(), key=keyf)
print(label, " ".join(f"{k}={v[k] or '·'}" for k in keys))
# Fill A1:C5 with r{row}c{col}
def fill():
for r in range(1, 6):
for c in range(1, 4):
ws.getCells(r, c).setValue(f"r{r}c{c}")
fill()
dump("[initial] col A:", "A1:A5")
# ---------- rows ----------
ws.insertRow(3, 2) # insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow] col A:", "A1:A7")
ws.deleteRow(3, 2) # delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow] col A:", "A1:A5")
ws.clearRow(2, 1) # clear row 2 (positions do not shift)
dump("[clearRow] col A:", "A1:A5")
ws.copyRow("Sheet1", 1, 1, 4) # overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow] col A:", "A1:A5")
fill()
# ---------- columns ----------
ws.insertCol(2, 1) # insert 1 column before column B
dump("[insertCol] row 1:", "A1:D1")
ws.deleteCol(2, 1) # delete the column B just inserted
dump("[deleteCol] row 1:", "A1:C1")
ws.clearCol(3, 1) # clear column C
dump("[clearCol] row 1:", "A1:C1")
fill()
ws.copyCol("Sheet1", 1, 2, 5) # copy columns A,B onto columns E,F
dump("[copyCol] row 1:", "A1:F1")
# ---------- cell ranges ----------
ws.insertCell("B2:B3", pyosbxl.enums.XlInsertDirection.InsertDirectionDown)
dump("[insertCell] col B:", "B1:B5")
ws.deleteCell("B2:B3", pyosbxl.enums.XlInsertDirection.DeleteDirectionUp)
dump("[deleteCell] col B:", "B1:B5")
ws.clearCell("A1:B2")
dump("[clearCell] A1:B2:", "A1:B2")
ws.copyCell("Sheet1", "A3:B4", "F8")
dump("[copyCell] F8:G9:", "F8:G9")
wb.save()
wb.close()
```

#### [Java]

```java
// rows-cols.java — same verification as the Node.js sample (javaosbxl)
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
// Print a range on one line, ordered by row → column (· marks an empty cell)
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
// Fill A1:C5 with r{row}c{col}
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
dump("[initial] col A:", "A1:A5");
// ---------- rows ----------
ws.insertRow(3, 2); // insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow] col A:", "A1:A7");
ws.deleteRow(3, 2); // delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow] col A:", "A1:A5");
ws.clearRow(2, 1); // clear row 2 (positions do not shift)
dump("[clearRow] col A:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4); // overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow] col A:", "A1:A5");
fill();
// ---------- columns ----------
ws.insertCol(2, 1); // insert 1 column before column B
dump("[insertCol] row 1:", "A1:D1");
ws.deleteCol(2, 1); // delete the column B just inserted
dump("[deleteCol] row 1:", "A1:C1");
ws.clearCol(3, 1); // clear column C
dump("[clearCol] row 1:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5); // copy columns A,B onto columns E,F
dump("[copyCol] row 1:", "A1:F1");
// ---------- cell ranges ----------
ws.insertCell("B2:B3", XlInsertDirection.InsertDirectionDown);
dump("[insertCell] col B:", "B1:B5");
ws.deleteCell("B2:B3", XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] col B:", "B1:B5");
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
// rows-cols.md — insert/delete/clear/copy for rows, columns and cell ranges (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("rows-cols-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// Print a range on one line, ordered by row → column (· marks an empty cell)
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
// Fill A1:C5 with r{row}c{col}
$fill = function () use ($ws): void {
for ($r = 1; $r <= 5; $r++) {
for ($c = 1; $c <= 3; $c++) {
$ws->getCells($r, $c)->setValue("r{$r}c{$c}");
}
}
};
$fill();
$dump("[initial] col A:", "A1:A5");
// ---------- rows ----------
$ws->insertRow(3, 2); // insert 2 rows BEFORE row 3 → existing data shifts down
$dump("[insertRow] col A:", "A1:A7");
$ws->deleteRow(3, 2); // delete 2 rows starting at row 3 → rows below move up
$dump("[deleteRow] col A:", "A1:A5");
$ws->clearRow(2, 1); // clear row 2 (positions do not shift)
$dump("[clearRow] col A:", "A1:A5");
$ws->copyRow("Sheet1", 1, 1, 4); // overwrite-copy row 1 of the same sheet onto row 4
$dump("[copyRow] col A:", "A1:A5");
$fill();
// ---------- columns ----------
$ws->insertCol(2, 1); // insert 1 column before column B
$dump("[insertCol] row 1:", "A1:D1");
$ws->deleteCol(2, 1); // delete the column B just inserted
$dump("[deleteCol] row 1:", "A1:C1");
$ws->clearCol(3, 1); // clear column C
$dump("[clearCol] row 1:", "A1:C1");
$fill();
$ws->copyCol("Sheet1", 1, 2, 5); // copy columns A,B onto columns E,F
$dump("[copyCol] row 1:", "A1:F1");
// ---------- cell ranges ----------
$ws->insertCell("B2:B3", \phposbxl\enums\XlInsertDirection::InsertDirectionDown);
$dump("[insertCell] col B:", "B1:B5");
$ws->deleteCell("B2:B3", \phposbxl\enums\XlInsertDirection::DeleteDirectionUp);
$dump("[deleteCell] col B:", "B1:B5");
$ws->clearCell("A1:B2");
$dump("[clearCell] A1:B2:", "A1:B2");
$ws->copyCell("Sheet1", "A3:B4", "F8");
$dump("[copyCell] F8:G9:", "F8:G9");
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// rows-cols.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("rows-cols-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// Print a range on one line, ordered by row → column (· marks an empty cell)
const dump = (label, range) => {
const v = ws.getRange(range).getValue();
const keys = Object.keys(v).map((k) => /^([A-Z]+)(\d+)$/.exec(k)).sort((a, b) =>
Number(a[2]) - Number(b[2]) ||
app.convertToColumnNumber(a[1]) - app.convertToColumnNumber(b[1]));
console.log(`${label} ${keys.map((m) => `${m[1]}${m[2]}=${v[m[1] + m[2]] || "·"}`).join(" ")}`);
};
// Fill A1:C5 with r{row}c{col}
const fill = () => {
for (let r = 1; r <= 5; r++) for (let c = 1; c <= 3; c++) ws.getCells(r, c).setValue(`r${r}c${c}`);
};
fill();
dump("[initial] col A:", "A1:A5");
// ---------- rows ----------
ws.insertRow(3, 2); // insert 2 rows BEFORE row 3 → existing data shifts down
dump("[insertRow] col A:", "A1:A7");
ws.deleteRow(3, 2); // delete 2 rows starting at row 3 → rows below move up
dump("[deleteRow] col A:", "A1:A5");
ws.clearRow(2, 1); // clear row 2 (positions do not shift)
dump("[clearRow] col A:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4); // overwrite-copy row 1 of the same sheet onto row 4
dump("[copyRow] col A:", "A1:A5");
fill();
// ---------- columns ----------
ws.insertCol(2, 1); // insert 1 column before column B
dump("[insertCol] row 1:", "A1:D1");
ws.deleteCol(2, 1); // delete the column B just inserted
dump("[deleteCol] row 1:", "A1:C1");
ws.clearCol(3, 1); // clear column C
dump("[clearCol] row 1:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5); // copy columns A,B onto columns E,F
dump("[copyCol] row 1:", "A1:F1");
// ---------- cell ranges ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] col B:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] col B:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell] F8:G9:", "F8:G9");
wb.save();
wb.close();
```

```plaintext
[initial] col A: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[insertRow] col A: A1=r1c1 A2=r2c1 A3=· A4=· A5=r3c1 A6=r4c1 A7=r5c1
[deleteRow] col A: A1=r1c1 A2=r2c1 A3=r3c1 A4=r4c1 A5=r5c1
[clearRow] col A: A1=r1c1 A2=· A3=r3c1 A4=r4c1 A5=r5c1
[copyRow] col A: A1=r1c1 A2=· A3=r3c1 A4=r1c1 A5=r5c1
[insertCol] row 1: A1=r1c1 B1=· C1=r1c2 D1=r1c3
[deleteCol] row 1: A1=r1c1 B1=r1c2 C1=r1c3
[clearCol] row 1: A1=r1c1 B1=r1c2 C1=·
[copyCol] row 1: A1=r1c1 B1=r1c2 C1=r1c3 D1=· E1=r1c1 F1=r1c2
[insertCell] col B: B1=r1c2 B2=· B3=· B4=r2c2 B5=r3c2
[deleteCell] col B: B1=r1c2 B2=r2c2 B3=r3c2 B4=r4c2 B5=r5c2
[clearCell] A1:B2: A1=· B1=· A2=· B2=·
[copyCell] F8:G9: F8=r3c1 G8=r3c2 F9=r4c1 G9=r4c2
```

`·` marks an empty cell. You can see that `insertRow` left rows 3 and 4 blank and shifted the existing data down, that `clearRow` did not shift positions, and that `copyRow` overwrote row 4 with the contents of row 1.

## Code 2: `Row` / `Col` objects

Output:

#### [Node.js]

```javascript
// row-col-props.js — row height, column width, hidden state, autofit
const { nodeosbxl } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("row-col-props.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- Row ----------
const row = ws.getRow(1); // row 1
console.log("default height =", row.getHeight());
row.setHeight(30);
console.log("after setHeight(30) =", row.getHeight());
console.log("getHidden default =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true); // hide row 2
console.log("row 2 hidden =", ws.getRow(2).getHidden());
// ---------- Col ----------
const col = ws.getCol(1); // column 1 (column A). Column numbers are passed as numbers
console.log("default col width = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -> columnWidth=" + col.getColumnWidth(),
"/ width=" + col.getWidth());
ws.getCol(2).setHidden(true); // hide column B
console.log("col B hidden =", ws.getCol(2).getHidden());
ws.getRange("C1:C1").setValue("Sample string for autofit");
console.log("col C bestFit default =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true); // autofit the column width to the contents
console.log("setBestFit(true) -> bestFit =", ws.getCol(3).getBestFit(),
"/ columnWidth =", ws.getCol(3).getColumnWidth());
// Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1); // group rows 4-6 at level 1
ws.setColOutline(2, 3, 1); // group columns B-C at level 1
console.log("row 4 outlineLevel =", ws.getRow(4).getOutlineLevel(),
"/ col B outlineLevel =", ws.getCol(2).getOutlineLevel());
wb.save();
wb.close();
```

#### [Python]

```python
# row-col-props.py — same verification as the Node.js sample (pyosbxl)
import pyosbxl
app = pyosbxl.App()
wb = app.createWorkBook("row-col-props-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
def num(x): # render integer-valued floats as ints (to match the JS output)
return int(x) if isinstance(x, float) and x.is\_integer() else x
def boolstr(x): # render Python's True/False as JS-style true/false
return str(x).lower()
# ---------- Row ----------
row = ws.getRow(1) # row 1
print("default height =", num(row.getHeight()))
row.setHeight(30)
print("after setHeight(30) =", num(row.getHeight()))
print("getHidden default =", boolstr(row.getHidden()), "/ getOutlineLevel =", row.getOutlineLevel())
ws.getRow(2).setHidden(True) # hide row 2
print("row 2 hidden =", boolstr(ws.getRow(2).getHidden()))
# ---------- Col ----------
col = ws.getCol(1) # column 1 (column A). Column numbers are passed as numbers
print("default col width = columnWidth:", col.getColumnWidth(), "/ width:", num(col.getWidth()))
col.setColumnWidth(20)
print("setColumnWidth(20) -> columnWidth=" + str(num(col.getColumnWidth())),
"/ width=" + str(num(col.getWidth())))
ws.getCol(2).setHidden(True) # hide column B
print("col B hidden =", boolstr(ws.getCol(2).getHidden()))
ws.getRange("C1:C1").setValue("Sample string for autofit")
print("col C bestFit default =", boolstr(ws.getCol(3).getBestFit()))
ws.getCol(3).setBestFit(True) # autofit the column width to the contents
print("setBestFit(true) -> bestFit =", boolstr(ws.getCol(3).getBestFit()),
"/ columnWidth =", ws.getCol(3).getColumnWidth())
# Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1) # group rows 4-6 at level 1
ws.setColOutline(2, 3, 1) # group columns B-C at level 1
print("row 4 outlineLevel =", ws.getRow(4).getOutlineLevel(),
"/ col B outlineLevel =", ws.getCol(2).getOutlineLevel())
wb.save()
wb.close()
```

#### [Java]

```java
// row-col-props.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
public class Main {
// Render integer-valued doubles as ints (to match the JS output)
static String num(double d) {
if (d == Math.rint(d) && !Double.isInfinite(d)) return String.valueOf((long) d);
return String.valueOf(d);
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("row-col-props-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- Row ----------
RowWrapper row = ws.getRow(1); // row 1
System.out.println("default height = " + num(row.getHeight()));
row.setHeight(30);
System.out.println("after setHeight(30) = " + num(row.getHeight()));
System.out.println("getHidden default = " + row.getHidden()
+ " / getOutlineLevel = " + row.getOutlineLevel());
ws.getRow(2).setHidden(true); // hide row 2
System.out.println("row 2 hidden = " + ws.getRow(2).getHidden());
// ---------- Col ----------
ColWrapper col = ws.getCol(1); // column 1 (column A). Column numbers are passed as numbers
System.out.println("default col width = columnWidth: " + num(col.getColumnWidth())
+ " / width: " + num(col.getWidth()));
col.setColumnWidth(20);
System.out.println("setColumnWidth(20) -> columnWidth=" + num(col.getColumnWidth())
+ " / width=" + num(col.getWidth()));
ws.getCol(2).setHidden(true); // hide column B
System.out.println("col B hidden = " + ws.getCol(2).getHidden());
ws.getRange("C1:C1").setValue("Sample string for autofit");
System.out.println("col C bestFit default = " + ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true); // autofit the column width to the contents
System.out.println("setBestFit(true) -> bestFit = " + ws.getCol(3).getBestFit()
+ " / columnWidth = " + num(ws.getCol(3).getColumnWidth()));
// Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1); // group rows 4-6 at level 1
ws.setColOutline(2, 3, 1); // group columns B-C at level 1
System.out.println("row 4 outlineLevel = " + ws.getRow(4).getOutlineLevel()
+ " / col B outlineLevel = " + ws.getCol(2).getOutlineLevel());
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// row-col-props.md — row height, column width, hidden state, autofit (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("row-col-props-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// echo of a PHP bool prints 1/empty string, so render JS-style true/false
$boolstr = fn(bool $b): string => $b ? "true" : "false";
// ---------- Row ----------
$row = $ws->getRow(1); // row 1
echo "default height = ", $row->getHeight(), "\n";
$row->setHeight(30);
echo "after setHeight(30) = ", $row->getHeight(), "\n";
echo "getHidden default = ", $boolstr($row->getHidden()), " / getOutlineLevel = ", $row->getOutlineLevel(), "\n";
$ws->getRow(2)->setHidden(true); // hide row 2
echo "row 2 hidden = ", $boolstr($ws->getRow(2)->getHidden()), "\n";
// ---------- Col ----------
$col = $ws->getCol(1); // column 1 (column A). Column numbers are passed as numbers
echo "default col width = columnWidth: ", $col->getColumnWidth(), " / width: ", $col->getWidth(), "\n";
$col->setColumnWidth(20);
echo "setColumnWidth(20) -> columnWidth=", $col->getColumnWidth(),
" / width=", $col->getWidth(), "\n";
$ws->getCol(2)->setHidden(true); // hide column B
echo "col B hidden = ", $boolstr($ws->getCol(2)->getHidden()), "\n";
$ws->getRange("C1:C1")->setValue("Sample string for autofit");
echo "col C bestFit default = ", $boolstr($ws->getCol(3)->getBestFit()), "\n";
$ws->getCol(3)->setBestFit(true); // autofit the column width to the contents
echo "setBestFit(true) -> bestFit = ", $boolstr($ws->getCol(3)->getBestFit()),
" / columnWidth = ", $ws->getCol(3)->getColumnWidth(), "\n";
// Outlines (grouping) are set on the WorkSheet side
$ws->setRowOutline(4, 6, 1); // group rows 4-6 at level 1
$ws->setColOutline(2, 3, 1); // group columns B-C at level 1
echo "row 4 outlineLevel = ", $ws->getRow(4)->getOutlineLevel(),
" / col B outlineLevel = ", $ws->getCol(2)->getOutlineLevel(), "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// row-col-props.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("row-col-props-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- Row ----------
const row = ws.getRow(1); // row 1
console.log("default height =", row.getHeight());
row.setHeight(30);
console.log("after setHeight(30) =", row.getHeight());
console.log("getHidden default =", row.getHidden(), "/ getOutlineLevel =", row.getOutlineLevel());
ws.getRow(2).setHidden(true); // hide row 2
console.log("row 2 hidden =", ws.getRow(2).getHidden());
// ---------- Col ----------
const col = ws.getCol(1); // column 1 (column A). Column numbers are passed as numbers
console.log("default col width = columnWidth:", col.getColumnWidth(), "/ width:", col.getWidth());
col.setColumnWidth(20);
console.log("setColumnWidth(20) -> columnWidth=" + col.getColumnWidth(),
"/ width=" + col.getWidth());
ws.getCol(2).setHidden(true); // hide column B
console.log("col B hidden =", ws.getCol(2).getHidden());
ws.getRange("C1:C1").setValue("Sample string for autofit");
console.log("col C bestFit default =", ws.getCol(3).getBestFit());
ws.getCol(3).setBestFit(true); // autofit the column width to the contents
console.log("setBestFit(true) -> bestFit =", ws.getCol(3).getBestFit(),
"/ columnWidth =", ws.getCol(3).getColumnWidth());
// Outlines (grouping) are set on the WorkSheet side
ws.setRowOutline(4, 6, 1); // group rows 4-6 at level 1
ws.setColOutline(2, 3, 1); // group columns B-C at level 1
console.log("row 4 outlineLevel =", ws.getRow(4).getOutlineLevel(),
"/ col B outlineLevel =", ws.getCol(2).getOutlineLevel());
wb.save();
wb.close();
```

```plaintext
default height = 18.75
after setHeight(30) = 30
getHidden default = false / getOutlineLevel = 0
row 2 hidden = true
default col width = columnWidth: 8.38 / width: 72
setColumnWidth(20) -> columnWidth=20 / width=165
col B hidden = true
col C bestFit default = false
setBestFit(true) -> bestFit = true / columnWidth = 26.38
row 4 outlineLevel = 1 / col B outlineLevel = 1
```

The column width produced by `setBestFit(true)` is computed from the cell contents and the book's default font, so the value varies by environment (above it is 26.38).

## Notes

* **Row and column numbers are all 1-based numeric values.** Columns are passed as `1`, not `"A"`. To convert from a string, use `app.convertToColumnNumber("B")` → `2`.
* **`insert` / `delete` shift data; `clear` does not.** To just empty a row use `clearRow()` (existing data keeps its position); to remove the row itself use `deleteRow()`.
* **The `copy` family takes the copy source as a "sheet name".** Same-sheet copies also pass the sheet name.  
  Passing `enums.XlCopyContentType` as the fifth argument `copyType` narrows what gets copied. | Value | Number | What gets copied | |---|---|---| | `CopyContentTypeAll` | 0 | Values, cell formats, and formats of formula cells (default) | | `CopyContentTypeValueOnly` | 1 | Values only | | `CopyContentTypeValueAndFormat` | 2 | Values and cell formats | | `CopyContentTypeFormatOnly` | 3 | Cell formats only |
* **`Col.getWidth()` and `Col.getColumnWidth()` use different units.** Measured: for a default column A, `getWidth()` = 72 and `getColumnWidth()` = 8.38. After `setColumnWidth(20)`, `getWidth()` = 165 and `getColumnWidth()` = 20. **Only `setColumnWidth()` (character width) can be set**; `getWidth()` is read-only.
* **`setBestFit(true)` recomputes the column width from the cell contents at the time of the call.** It does not follow later content changes automatically, so call it after all data is in place.
* **Outlines (grouping) are `WorkSheet` methods.** Set them with `setRowOutline(startRow, endRow, level)` / `setColOutline(startCol, endCol, level)` and release them with `clearRowOutline(level)` / `clearColOutline(level)`. Check the result with `Row.getOutlineLevel()` / `Col.getOutlineLevel()`.

## Caveats

* **Use the `direction` of `insertCell` / `deleteCell` according to the operation.** For insertion the intended values are `InsertDirectionRight`(1) / `InsertDirectionDown`(2); for deletion, `DeleteDirectionLeft`(3) / `DeleteDirectionUp`(4).
* **Some things are not copied by the `copy` family.** Even with `CopyContentTypeAll`, **shapes (including charts), pivot tables, queryTable-style tables and data tables are not copied.**
* **The third argument of `copyCell` is a single "top-left cell of the destination".** `copyCell("Sheet1", "A3:B4", "F8")` expands to F8:G9 (the range size is determined by the source).
* **For `copyRow` / `copyCol`, the source is a sheet name and positions are row/column numbers (numeric).** Do not use A1-style range strings — they are easy to confuse, so be careful.

## See also

* [Index](index.md)
* Previous: [Cell value types and number formats](general-04-value-types.md) / Next: [Bulk input of large data](general-06-bulk-input.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
