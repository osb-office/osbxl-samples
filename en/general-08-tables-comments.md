# Tables and Comments — create, totals, structured refs

level: intermediate ／ verified: 2026-10-04 ／ tabs: Node.js・Python・Java・PHP・WebAssembly ／ [日本語](/ja/general-08-tables-comments.md)

## What you want to do

Turn a data range into an Excel "table" (`ListObject`) and work with its name, column names, style and totals row. Use "structured references" to refer to table columns from formulas, and add, read and remove cell comments (modern threaded comments) and memos (legacy notes).

## Classes used

Get tables from `ws.getListObjects()` and comments from `ws.getComments()`.

### `ListObjects` (collection of tables)

| Method | Description |
| --- | --- |
| `addList(name, A1C1, useFirstRowAsHeader, insertTotals)` | Creates a table from an existing cell range and returns the `Table`. |
| `addListFromRange(name, topA1, copyFromA1C1, useFirstRowAsHeader, insertTotals)` | Creates a table by **copying** the data to `topA1`. `copyFromA1C1` may point to another sheet, like `Sheet2!A1:C1`. |
| `getList(name)` | Gets a table. Throws a `table is not found` exception if it does not exist. |
| `removeList(name, deleteData?)` | Removes a table. `deleteData` defaults to `true` (**the data is deleted too**). Pass `false` to keep it. |

### `Table` (a single table)

| Group | Methods |
| --- | --- |
| Ranges | `getAllRange()` / `getHeaderRowRange()` / `getDataBodyRange()` / `getTotalRowRange()` (all return a `Range`; `getTotalRowRange()` throws `totalRow is not set` when there is no totals row) |
| Name | `getName()` / `changeTableName(name)` |
| Column names | `getColumnName()` / `setColumnName([...])` (also changes what the header cells display) |
| Style | `getTableStyleName()` / `setBuiltinStyleName(enums.XlDefaultTableStyle.Xxx, clearFormat)` / `setCustomTableStyleName(name, clearFormat)` |
| Totals row | `isShowTotals()` / `setShowTotals(bool)` / `getTotalRowFunction(col)` / `setTotalRowFunction(col, enums.XlTotalsCalculation.Xxx)` / `setCustomRowFunction(col, formula, isArray?)` / `getTotalRowLabel(col)` / `setTotalRowLabel(col, text)` |
| Filter / display | `getAutoFilter()` / `removeAutoFilter()` / `isShowAutoFilter()` / `isShowHeaders()` / `setShowHeaders(bool)` / `isShowTableStyleRowStripes()` and more |

### `Comments` (cell comments)

Excel has two comment systems: **modern** comments (threads, replies, done flag) and **legacy memos** (callout notes).

| Method | System | Description |
| --- | --- | --- |
| `setComment(A1, commentObject)` | modern | Adds one comment. On a cell that already has comments this **appends** (automatically as a reply). |
| `setCommentThread(A1, [commentObject...])` | modern | Sets a whole thread (**overwrites**). The first entry is the body, the rest are replies. |
| `getComment(A1)` | modern | Returns `Array<dto.CommentObject>`. Throws `comment not found` if there is none. |
| `removeComment(A1)` | modern | Removes modern comments. |
| `setMemo(A1, author, text, fontObject, visible?, rowColumnsNum?, colColumnsNum?)` | legacy | Sets a callout memo. `fontObject` can be `new dto.FontObject()`. |
| `getMemoText(A1)` / `getMemoAuthor(A1)` | legacy | Gets the memo text / author. Throws `memo not found` if there is none. |
| `removeMemo(A1)` | legacy | Removes a memo. |

`dto.CommentObject` has `setAuthor` / `setContent` / `setParentId` (parent ID of a reply) / `setDone` (done flag) plus the corresponding getters. Once set, an ID is automatically assigned as a UUID in `{...}` form.

### Totals functions `enums.XlTotalsCalculation`

| Member | Value | Meaning |
| --- | --- | --- |
| `TotalsCalculationNone` | 0 | No totals |
| `TotalsCalculationSum` | 1 | Sum |
| `TotalsCalculationAverage` | 2 | Average |
| `TotalsCalculationCount` | 3 | Count |
| `TotalsCalculationCountNums` | 4 | Count of numbers |
| `TotalsCalculationMin` / `TotalsCalculationMax` | 5 / 6 | Min / Max |
| `TotalsCalculationStdDev` / `TotalsCalculationVar` | 7 / 8 | StdDev / Var |
| `TotalsCalculationCustom` | 9 | Custom formula (set with `setCustomRowFunction`) |

## Code

Output:

#### [Node.js]

```javascript
// tables-comments.js — table (ListObject) creation, totals row, structured references, cell comments
const { nodeosbxl, enums, dto } = require("nodeosbxl");
const app = new nodeosbxl.App();
const wb = app.createWorkBook("tables-comments.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- fill in the data (1 header row + 3 data rows) ----------
const rows = [
["Product", "Unit Price", "Qty"],
["Apple", 120, 3],
["Orange", 80, 5],
["Grape", 300, 2],
];
const values = [];
rows.forEach((row, ri) => row.forEach((cell, ci) => {
const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
const o = new dto.InputValueObject();
if (typeof cell === "number") o.setNumberValue(a1, cell);
else o.setStringValue(a1, cell);
values.push(o);
}));
ws.setValueArray(values);
// ---------- 1. create the table ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] Table created");
console.log(` name = ${table.getName()}`);
console.log(` columns = ${JSON.stringify(table.getColumnName())}`);
console.log(` style = ${table.getTableStyleName()}`);
console.log(` full range = ${table.getAllRange().getAddress()}`);
console.log(` header = ${table.getHeaderRowRange().getAddress()}`);
console.log(` data body = ${table.getDataBodyRange().getAddress()}`);
console.log(` totals row = ${table.getTotalRowRange().getAddress()}`);
// ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] Totals row");
console.log(` label (col 1) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(` formula = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(` value = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);
// ---------- 3. change the table style ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] style after change = ${table.getTableStyleName()}`);
// ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])"); // no leading "="
console.log("[4] Structured reference");
console.log(` E1 formula = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(` E1 value = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);
// ---------- 5. comments (modern) and memos (legacy) ----------
const cm = ws.getComments();
const mkComment = (author, content) => {
const c = new dto.CommentObject();
c.setAuthor(author);
c.setContent(content);
return c;
};
// modern comment (thread): body + reply = 2 entries -> B2 (kept in the file after save)
cm.setCommentThread("B2", [mkComment("alice", "Please check this number"), mkComment("bob", "Checked")]);
// single modern comment -> D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -> A2
cm.setMemo("A2", "Carol", "Top-priority product", new dto.FontObject());
const thread = cm.getComment("B2");
console.log("[5] Comments");
console.log(` B2 thread count = ${thread.length}`);
thread.forEach((c, i) => console.log(` [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (reply)" : ""}`));
console.log(` D2 comment count = ${cm.getComment("D2").length}`);
console.log(` A2 memo = ${JSON.stringify(cm.getMemoText("A2"))} / author ${JSON.stringify(cm.getMemoAuthor("A2"))}`);
// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) => {
try { fn(); console.log(` after removal ${label}: no exception`); }
catch (e) { console.log(` after removal ${label}: exception ${e.message}`); }
};
tryRead("D2 comment", () => cm.getComment("D2"));
tryRead("A2 memo", () => cm.getMemoText("A2"));
console.log(` B2 is kept, so ${cm.getComment("B2").length} comments remain after save`);
const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[Check] table=${table.getName()} total(Unit Price)=${total} structuredRef(E1)=${e1}`);
wb.save();
wb.close();
```

#### [Python]

```python
# tables-comments.py — same verification as the Node.js sample (pyosbxl)
import json
import pyosbxl
def jdump(obj): # JSON.stringify-equivalent output (no separator spaces, non-ASCII kept as-is)
return json.dumps(obj, ensure\_ascii=False, separators=(",", ":"))
app = pyosbxl.App()
wb = app.createWorkBook("tables-comments-py.xlsx")
ws = wb.openWorkSheet("Sheet1")
# ---------- fill in the data (1 header row + 3 data rows) ----------
rows = [
["Product", "Unit Price", "Qty"],
["Apple", 120, 3],
["Orange", 80, 5],
["Grape", 300, 2],
]
values = []
for ri, row in enumerate(rows):
for ci, cell in enumerate(row):
a1 = app.convertFromRowColNumber(ri + 1, ci + 1)
o = pyosbxl.dto.InputValueObject()
if isinstance(cell, (int, float)):
o.setNumberValue(a1, cell)
else:
o.setStringValue(a1, cell)
values.append(o)
ws.setValueArray(values)
# ---------- 1. create the table ----------
lo = ws.getListObjects()
table = lo.addList("Sales", "A1:C4", True, True) # useFirstRowAsHeader, insertTotals
print("[1] Table created")
print(f" name = {table.getName()}")
print(f" columns = {jdump(table.getColumnName())}")
print(f" style = {table.getTableStyleName()}")
print(f" full range = {table.getAllRange().getAddress()}")
print(f" header = {table.getHeaderRowRange().getAddress()}")
print(f" data body = {table.getDataBodyRange().getAddress()}")
print(f" totals row = {table.getTotalRowRange().getAddress()}")
# ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total")
table.setTotalRowFunction(2, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
table.setTotalRowFunction(3, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
print("[2] Totals row")
print(f" label (col 1) = {jdump(table.getTotalRowLabel(1))}")
print(f" formula = {jdump(ws.getRange('B5:C5').getFormula())}")
print(f" value = {jdump(ws.getRange('B5:C5').getValue(True))}")
# ---------- 3. change the table style ----------
table.setBuiltinStyleName(pyosbxl.enums.XlDefaultTableStyle.TableStyleMedium9, True)
print(f"[3] style after change = {table.getTableStyleName()}")
# ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])") # no leading "="
print("[4] Structured reference")
print(f" E1 formula = {jdump(ws.getRange('E1:E1').getFormula())}")
print(f" E1 value = {jdump(ws.getRange('E1:E1').getValue(True))}")
# ---------- 5. comments (modern) and memos (legacy) ----------
cm = ws.getComments()
def mkComment(author, content):
c = pyosbxl.dto.CommentObject()
c.setAuthor(author)
c.setContent(content)
return c
# modern comment (thread): body + reply = 2 entries -> B2 (kept in the file after save)
cm.setCommentThread("B2", [mkComment("alice", "Please check this number"), mkComment("bob", "Checked")])
# single modern comment -> D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"))
# legacy memo (callout) -> A2
cm.setMemo("A2", "Carol", "Top-priority product", pyosbxl.dto.FontObject())
thread = cm.getComment("B2")
print("[5] Comments")
print(f" B2 thread count = {len(thread)}")
for i, c in enumerate(thread):
print(f" [{i}] {c.getAuthor()}: {c.getContent()}{' (reply)' if c.getParentId() else ''}")
print(f" D2 comment count = {len(cm.getComment('D2'))}")
print(f" A2 memo = {jdump(cm.getMemoText('A2'))} / author {jdump(cm.getMemoAuthor('A2'))}")
# removal: removeComment for modern, removeMemo for legacy
cm.removeComment("D2")
cm.removeMemo("A2")
# reading after removal throws "comment not found" / "memo not found" just like the Node.js version
def tryRead(label, fn):
try:
fn()
print(f" after removal {label}: no exception")
except Exception as e:
print(f" after removal {label}: exception {e}")
tryRead("D2 comment", lambda: cm.getComment("D2"))
tryRead("A2 memo", lambda: cm.getMemoText("A2"))
print(f" B2 is kept, so {len(cm.getComment('B2'))} comments remain after save")
total = ws.getRange("B5:B5").getValue(True)["B5"]
e1 = ws.getRange("E1:E1").getValue(True)["E1"]
print(f"[Check] table={table.getName()} total(Unit Price)={total} structuredRef(E1)={e1}")
wb.save()
wb.close()
```

#### [Java]

```java
// tables-comments.java — same verification as the Node.js sample (javaosbxl)
import com.osboffice.osbxl.\*;
import com.osboffice.osbxl.dto.\*;
import com.osboffice.osbxl.enums.\*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
public class Main {
// JSON.stringify-equivalent output helpers (Maps keep insertion order)
static String jsonMap(Map<String, String> m) {
StringBuilder sb = new StringBuilder("{");
boolean first = true;
for (Map.Entry<String, String> e : m.entrySet()) {
if (!first) sb.append(",");
first = false;
sb.append("\"").append(e.getKey()).append("\":\"").append(e.getValue()).append("\"");
}
return sb.append("}").toString();
}
static String jsonList(List<String> l) {
StringBuilder sb = new StringBuilder("[");
for (int i = 0; i < l.size(); i++) {
if (i > 0) sb.append(",");
sb.append("\"").append(l.get(i)).append("\"");
}
return sb.append("]").toString();
}
static CommentObjectWrapper mkComment(String author, String content) {
CommentObjectWrapper c = new CommentObjectWrapper();
c.setAuthor(author);
c.setContent(content);
return c;
}
public static void main(String[] args) {
AppWrapper app = new AppWrapper();
WorkBookWrapper wb = app.createWorkBook("tables-comments-java.xlsx");
WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");
// ---------- fill in the data (1 header row + 3 data rows) ----------
Object[][] rows = {
{"Product", "Unit Price", "Qty"},
{"Apple", 120, 3},
{"Orange", 80, 5},
{"Grape", 300, 2},
};
List<InputValueObjectWrapper> values = new ArrayList<>();
for (int ri = 0; ri < rows.length; ri++) {
for (int ci = 0; ci < rows[ri].length; ci++) {
String a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
InputValueObjectWrapper o = new InputValueObjectWrapper();
if (rows[ri][ci] instanceof Number) {
o.setNumberValue(a1, ((Number) rows[ri][ci]).doubleValue());
} else {
o.setStringValue(a1, (String) rows[ri][ci]);
}
values.add(o);
}
}
ws.setValueArray(values);
// ---------- 1. create the table ----------
ListObjectsWrapper lo = ws.getListObjects();
TableWrapper table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
System.out.println("[1] Table created");
System.out.println(" name = " + table.getName());
System.out.println(" columns = " + jsonList(table.getColumnName()));
System.out.println(" style = " + table.getTableStyleName());
System.out.println(" full range = " + table.getAllRange().getAddress());
System.out.println(" header = " + table.getHeaderRowRange().getAddress());
System.out.println(" data body = " + table.getDataBodyRange().getAddress());
System.out.println(" totals row = " + table.getTotalRowRange().getAddress());
// ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total");
table.setTotalRowFunction(2, XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, XlTotalsCalculation.TotalsCalculationSum);
System.out.println("[2] Totals row");
System.out.println(" label (col 1) = \"" + table.getTotalRowLabel(1) + "\"");
System.out.println(" formula = " + jsonMap(ws.getRange("B5:C5").getFormula()));
System.out.println(" value = " + jsonMap(ws.getRange("B5:C5").getValue(true)));
// ---------- 3. change the table style ----------
table.setBuiltinStyleName(XlDefaultTableStyle.TableStyleMedium9, true);
System.out.println("[3] style after change = " + table.getTableStyleName());
// ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])"); // no leading "="
System.out.println("[4] Structured reference");
System.out.println(" E1 formula = " + jsonMap(ws.getRange("E1:E1").getFormula()));
System.out.println(" E1 value = " + jsonMap(ws.getRange("E1:E1").getValue(true)));
// ---------- 5. comments (modern) and memos (legacy) ----------
CommentsWrapper cm = ws.getComments();
// modern comment (thread): body + reply = 2 entries -> B2 (kept in the file after save)
cm.setCommentThread("B2", Arrays.asList(
mkComment("alice", "Please check this number"), mkComment("bob", "Checked")));
// single modern comment -> D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -> A2
cm.setMemo("A2", "Carol", "Top-priority product", new FontObjectWrapper());
List<CommentObjectWrapper> thread = cm.getComment("B2");
System.out.println("[5] Comments");
System.out.println(" B2 thread count = " + thread.size());
for (int i = 0; i < thread.size(); i++) {
CommentObjectWrapper c = thread.get(i);
String pid = c.getParentId();
boolean isReply = pid != null && !pid.isEmpty();
System.out.println(" [" + i + "] " + c.getAuthor() + ": " + c.getContent()
+ (isReply ? " (reply)" : ""));
}
System.out.println(" D2 comment count = " + cm.getComment("D2").size());
System.out.println(" A2 memo = \"" + cm.getMemoText("A2") + "\" / author \"" + cm.getMemoAuthor("A2") + "\"");
// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
cm.removeComment("D2");
cm.removeMemo("A2");
try {
cm.getComment("D2");
System.out.println(" after removal D2 comment: no exception");
} catch (RuntimeException e) {
System.out.println(" after removal D2 comment: exception " + e.getMessage());
}
try {
cm.getMemoText("A2");
System.out.println(" after removal A2 memo: no exception");
} catch (RuntimeException e) {
System.out.println(" after removal A2 memo: exception " + e.getMessage());
}
System.out.println(" B2 is kept, so " + cm.getComment("B2").size() + " comments remain after save");
String total = ws.getRange("B5:B5").getValue(true).get("B5");
String e1 = ws.getRange("E1:E1").getValue(true).get("E1");
System.out.println("[Check] table=" + table.getName() + " total(Unit Price)=" + total + " structuredRef(E1)=" + e1);
wb.save();
wb.close();
}
}
```

#### [PHP]

```php
// tables-comments.md — table (ListObject) creation, totals row, structured references, cell comments (phposbxl)
$app = new \phposbxl\App();
$wb = $app->createWorkBook("tables-comments-php.xlsx");
$ws = $wb->openWorkSheet("Sheet1");
// JSON.stringify-compatible output (non-ASCII kept as-is)
$jq = fn($x): string => json\_encode($x, JSON\_UNESCAPED\_UNICODE | JSON\_UNESCAPED\_SLASHES);
// ---------- fill in the data (1 header row + 3 data rows) ----------
$rows = [
["Product", "Unit Price", "Qty"],
["Apple", 120, 3],
["Orange", 80, 5],
["Grape", 300, 2],
];
$values = [];
foreach ($rows as $ri => $row) {
foreach ($row as $ci => $cell) {
$a1 = $app->convertFromRowColNumber($ri + 1, $ci + 1);
$o = new \phposbxl\dto\InputValueObject();
if (is\_int($cell) || is\_float($cell)) {
$o->setNumberValue($a1, $cell);
} else {
$o->setStringValue($a1, $cell);
}
$values[] = $o;
}
}
$ws->setValueArray($values);
// ---------- 1. create the table ----------
$lo = $ws->getListObjects();
$table = $lo->addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
echo "[1] Table created\n";
echo " name = ", $table->getName(), "\n";
echo " columns = ", $jq($table->getColumnName()), "\n";
echo " style = ", $table->getTableStyleName(), "\n";
echo " full range = ", $table->getAllRange()->getAddress(), "\n";
echo " header = ", $table->getHeaderRowRange()->getAddress(), "\n";
echo " data body = ", $table->getDataBodyRange()->getAddress(), "\n";
echo " totals row = ", $table->getTotalRowRange()->getAddress(), "\n";
// ---------- 2. totals row ----------
$table->setTotalRowLabel(1, "Total");
$table->setTotalRowFunction(2, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
$table->setTotalRowFunction(3, \phposbxl\enums\XlTotalsCalculation::TotalsCalculationSum);
echo "[2] Totals row\n";
echo " label (col 1) = ", $jq($table->getTotalRowLabel(1)), "\n";
echo " formula = ", $jq($ws->getRange("B5:C5")->getFormula()), "\n";
echo " value = ", $jq($ws->getRange("B5:C5")->getValue(true)), "\n";
// ---------- 3. change the table style ----------
$table->setBuiltinStyleName(\phposbxl\enums\XlDefaultTableStyle::TableStyleMedium9, true);
echo "[3] style after change = ", $table->getTableStyleName(), "\n";
// ---------- 4. structured reference (formula referring to a table column) ----------
$ws->getRange("E1:E1")->setFormula("SUM(Sales[Unit Price])"); // no leading "="
echo "[4] Structured reference\n";
echo " E1 formula = ", $jq($ws->getRange("E1:E1")->getFormula()), "\n";
echo " E1 value = ", $jq($ws->getRange("E1:E1")->getValue(true)), "\n";
// ---------- 5. comments (modern) and memos (legacy) ----------
$cm = $ws->getComments();
$mkComment = function (string $author, string $content): \phposbxl\dto\CommentObject {
$c = new \phposbxl\dto\CommentObject();
$c->setAuthor($author);
$c->setContent($content);
return $c;
};
// modern comment (thread): body + reply = 2 entries -> B2 (kept in the file after save)
$cm->setCommentThread("B2", [$mkComment("alice", "Please check this number"), $mkComment("bob", "Checked")]);
// single modern comment -> D2 (used to demonstrate removal)
$cm->setComment("D2", $mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -> A2
$cm->setMemo("A2", "Carol", "Top-priority product", new \phposbxl\dto\FontObject());
$thread = $cm->getComment("B2");
echo "[5] Comments\n";
echo " B2 thread count = ", count($thread), "\n";
foreach ($thread as $i => $c) {
echo " [{$i}] ", $c->getAuthor(), ": ", $c->getContent(), $c->getParentId() ? " (reply)" : "", "\n";
}
echo " D2 comment count = ", count($cm->getComment("D2")), "\n";
echo " A2 memo = ", $jq($cm->getMemoText("A2")), " / author ", $jq($cm->getMemoAuthor("A2")), "\n";
// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
$cm->removeComment("D2");
$cm->removeMemo("A2");
$tryRead = function (string $label, callable $fn): void {
try {
$fn();
echo " after removal {$label}: no exception\n";
} catch (\Exception $e) {
echo " after removal {$label}: exception ", $e->getMessage(), "\n";
}
};
$tryRead("D2 comment", fn() => $cm->getComment("D2"));
$tryRead("A2 memo", fn() => $cm->getMemoText("A2"));
echo " B2 is kept, so ", count($cm->getComment("B2")), " comments remain after save\n";
$total = $ws->getRange("B5:B5")->getValue(true)["B5"];
$e1 = $ws->getRange("E1:E1")->getValue(true)["E1"];
echo "[Check] table=", $table->getName(), " total(Unit Price)=", $total, " structuredRef(E1)=", $e1, "\n";
$wb->save();
$wb->close();
```

#### [WebAssembly]

```javascript
// tables-comments.wasm.js — same verification as the Node.js sample (wasmosbxl / embind)
// Note (WebAssembly build): load the module first with createOsbxl(createModule).
// osbxl / dto / chart / enums used below are the four namespaces it returns
// (the equivalent of require("nodeosbxl") in the Node.js build).
const app = new osbxl.App();
const wb = app.createWorkBook("tables-comments-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");
// ---------- fill in the data (1 header row + 3 data rows) ----------
const rows = [
["Product", "Unit Price", "Qty"],
["Apple", 120, 3],
["Orange", 80, 5],
["Grape", 300, 2],
];
const values = [];
rows.forEach((row, ri) => row.forEach((cell, ci) => {
const a1 = app.convertFromRowColNumber(ri + 1, ci + 1);
const o = new dto.InputValueObject();
if (typeof cell === "number") o.setNumberValue(a1, cell);
else o.setStringValue(a1, cell);
values.push(o);
}));
ws.setValueArray(values);
// ---------- 1. create the table ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] Table created");
console.log(` name = ${table.getName()}`);
console.log(` columns = ${JSON.stringify(table.getColumnName())}`);
console.log(` style = ${table.getTableStyleName()}`);
console.log(` full range = ${table.getAllRange().getAddress()}`);
console.log(` header = ${table.getHeaderRowRange().getAddress()}`);
console.log(` data body = ${table.getDataBodyRange().getAddress()}`);
console.log(` totals row = ${table.getTotalRowRange().getAddress()}`);
// ---------- 2. totals row ----------
table.setTotalRowLabel(1, "Total");
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] Totals row");
console.log(` label (col 1) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(` formula = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(` value = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);
// ---------- 3. change the table style ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] style after change = ${table.getTableStyleName()}`);
// ---------- 4. structured reference (formula referring to a table column) ----------
ws.getRange("E1:E1").setFormula("SUM(Sales[Unit Price])"); // no leading "="
console.log("[4] Structured reference");
console.log(` E1 formula = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(` E1 value = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);
// ---------- 5. comments (modern) and memos (legacy) ----------
const cm = ws.getComments();
const mkComment = (author, content) => {
const c = new dto.CommentObject();
c.setAuthor(author);
c.setContent(content);
return c;
};
// modern comment (thread): body + reply = 2 entries -> B2 (kept in the file after save)
cm.setCommentThread("B2", [mkComment("alice", "Please check this number"), mkComment("bob", "Checked")]);
// single modern comment -> D2 (used to demonstrate removal)
cm.setComment("D2", mkComment("dave", "Comment to delete later"));
// legacy memo (callout) -> A2
cm.setMemo("A2", "Carol", "Top-priority product", new dto.FontObject());
const thread = cm.getComment("B2");
console.log("[5] Comments");
console.log(` B2 thread count = ${thread.length}`);
thread.forEach((c, i) => console.log(` [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (reply)" : ""}`));
console.log(` D2 comment count = ${cm.getComment("D2").length}`);
console.log(` A2 memo = ${JSON.stringify(cm.getMemoText("A2"))} / author ${JSON.stringify(cm.getMemoAuthor("A2"))}`);
// removal: removeComment for modern, removeMemo for legacy (reading after removal throws "not found")
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) => {
try { fn(); console.log(` after removal ${label}: no exception`); }
catch (e) {
// embind exception messages carry a prefix like "Comments.getComment: ";
// strip it so the output matches the Node.js version exactly
const msg = String(e.message || e).replace(/^[A-Za-z\_$][\w$]\*\.[\w$]+: /, "");
console.log(` after removal ${label}: exception ${msg}`);
}
};
tryRead("D2 comment", () => cm.getComment("D2"));
tryRead("A2 memo", () => cm.getMemoText("A2"));
console.log(` B2 is kept, so ${cm.getComment("B2").length} comments remain after save`);
const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[Check] table=${table.getName()} total(Unit Price)=${total} structuredRef(E1)=${e1}`);
wb.save();
wb.close();
```

```plaintext
[1] Table created
name = Sales
columns = ["Product","Unit Price","Qty"]
style = TableStyleMedium2
full range = A1:C5
header = A1:C1
data body = A2:C4
totals row = A5:C5
[2] Totals row
label (col 1) = "Total"
formula = {"B5":"SUBTOTAL(109,Sales[Unit Price])","C5":"SUBTOTAL(109,Sales[Qty])"}
value = {"B5":"500","C5":"10"}
[3] style after change = TableStyleMedium9
[4] Structured reference
E1 formula = {"E1":"SUM(Sales[Unit Price])"}
E1 value = {"E1":"500"}
[5] Comments
B2 thread count = 2
[0] alice: Please check this number
[1] bob: Checked (reply)
D2 comment count = 1
A2 memo = "Top-priority product" / author "Carol"
after removal D2 comment: exception comment not found
after removal A2 memo: exception memo not found
B2 is kept, so 2 comments remain after save
[Check] table=Sales total(Unit Price)=500 structuredRef(E1)=500
```

## Notes

* **Tables are handled by name, not by range.** From the return value of `addList()` (a `Table`) or from `getList("Sales")` you can pull out the header, data body and totals row ranges as `Range` objects.  
  Because a totals row was inserted, `getAllRange()` is `A1:C5` (data `A1:C4` + totals row 5).
* **The totals row holds `SUBTOTAL` formulas.** Setting `setTotalRowFunction(2, ...Sum)` auto-generates `SUBTOTAL(109,Sales[Unit Price])` in B5 (109 = sum, nothing ignored). Average maps to `SUBTOTAL(101,...)` and Count to `SUBTOTAL(103,...)` (measured). To put an arbitrary formula there, use `setCustomRowFunction(col, "MAX(C2:C4)")`; in that case `getTotalRowFunction()` returns `TotalsCalculationCustom`(9).
* **Structured references are written as "TableName[ColumnName]".** `SUM(Sales[Unit Price])` evaluates to 500. A bare column name (`SUM(Unit Price)`) gives a `#NAME?` error.  
  Formulas do not take a leading `=`.
* **Renaming a column also updates structured references inside the table.** When you change column names with `setColumnName([...])`, `Sales[OldName]` in the totals row and in existing formulas is automatically replaced with `Sales[NewName]` (measured).
* **There are two comment systems.** The **modern** one with threads/replies/done flag (`setComment`, `setCommentThread`, `getComment`) and the legacy Excel callout **memo** (`setMemo`, `getMemoText`). Pick whichever fits the use case.  
  Both are persisted to the file by `save()`.

## Caveats

* **Table names that look like cell references are not allowed.** `addList("A1", ...)` or `addList("S2", ...)` throws a `table name is invalid format` exception.  
  Re-creating a table with an existing name throws `table name is already exists`.
* **With `addList(insertTotals=true)` the default totals are "None" for every column (Excel-compatible, fixed 2026-10-01).** Only the hardcoded label `集計` ("Total" in Japanese) is placed automatically in the first column — set your own with `setTotalRowLabel()`. No totals function is set for the data columns (older versions defaulted the last column to `Count` (`SUBTOTAL(103)`)). Set the columns you want totaled explicitly with `setTotalRowFunction()`.
* **`removeList()` deletes the data by default.** `removeList("Sales")` behaves as `deleteData=true` and wipes the cell values too.  
  To remove only the table outline and keep the values, use `removeList("Sales", false)`.
* **Modern comments require a `dto.CommentObject` instance.** A **plain literal like `setComment("A1", { text: "..." })` from the docs' `@example` throws `Invalid argument`** (measured). Build the object with `new dto.CommentObject()` → `setAuthor()` / `setContent()` and pass that.
* **`setComment` appends, `setCommentThread` overwrites.** Calling `setComment` on a cell that already has comments increases the count, and the new one automatically becomes a reply (with `parentId`) in the existing thread.  
  Use `setCommentThread` when you want to replace the whole thread (measured).
* **Reading a nonexistent comment/memo throws.** `getComment()` throws `comment not found` and `getMemoText()` / `getMemoAuthor()` throw `memo not found` (fixed 2026-10-01). Verify state after removal by catching the exception in a try/catch.
* **Averages etc. in the totals row come back as raw floating point.** `getValue(true)` returns pre-rounding values as strings, so an average `SUBTOTAL(101,...)` can be something like `3.3333333333333335` (round it via the display format).

## See also

* [Index](index.md)
* Previous: [AutoFilter and Sorting](general-07-autofilter-sort.md) / Next: [Conditional Formatting](general-09-conditional-format.md)
* API reference

  #### [Node.js]

  #### [Python]

  #### [Java]

  #### [PHP]

  #### [WebAssembly]
