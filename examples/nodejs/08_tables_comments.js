// tables-comments.js — テーブル(ListObject)作成・集計行・構造化参照・セルコメント
const { nodeosbxl, enums, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("tables-comments.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- データ投入（ヘッダ1行 + データ3行）----------
const rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", "120", "3"],
    ["ミカン", "80", "5"],
    ["ブドウ", "300", "2"],
];
ws.setRangeData("A1", rows);

// ---------- 1. テーブルを作成 ----------
const lo = ws.getListObjects();
const table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
console.log("[1] テーブル作成");
console.log(`    名前      = ${table.getName()}`);
console.log(`    列名      = ${JSON.stringify(table.getColumnName())}`);
console.log(`    スタイル  = ${table.getTableStyleName()}`);
console.log(`    全範囲    = ${table.getAllRange().getAddress()}`);
console.log(`    ヘッダ    = ${table.getHeaderRowRange().getAddress()}`);
console.log(`    データ部  = ${table.getDataBodyRange().getAddress()}`);
console.log(`    集計行    = ${table.getTotalRowRange().getAddress()}`);

// ---------- 2. 集計行 ----------
table.setTotalRowLabel(0, "合計");
table.setTotalRowFunction(1, enums.XlTotalsCalculation.TotalsCalculationSum);
table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum);
console.log("[2] 集計行");
console.log(`    ラベル(1列目 (0-based: 0)) = ${JSON.stringify(table.getTotalRowLabel(1))}`);
console.log(`    数式          = ${JSON.stringify(ws.getRange("B5:C5").getFormula())}`);
console.log(`    値            = ${JSON.stringify(ws.getRange("B5:C5").getValue(true))}`);

// ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(enums.XlDefaultTableStyle.TableStyleMedium9, true);
console.log(`[3] スタイル変更後 = ${table.getTableStyleName()}`);

// ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
console.log("[4] 構造化参照");
console.log(`    E1 数式 = ${JSON.stringify(ws.getRange("E1:E1").getFormula())}`);
console.log(`    E1 値   = ${JSON.stringify(ws.getRange("E1:E1").getValue(true))}`);

// ---------- 5. コメント（新式）とメモ（旧式）----------
const cm = ws.getComments();
const mkComment = (author, content) => {
    const c = new dto.CommentObject();
    c.setAuthor(author);
    c.setContent(content);
    return c;
};

// 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")]);
// 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
// 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", new dto.FontObject());

const thread = cm.getComment("B2");
console.log("[5] コメント");
console.log(`    B2 スレッド件数 = ${thread.length}`);
thread.forEach((c, i) => console.log(`      [${i}] ${c.getAuthor()}: ${c.getContent()}${c.getParentId() ? " (返信)" : ""}`));
console.log(`    D2 コメント件数 = ${cm.getComment("D2").length}`);
console.log(`    A2 メモ = ${JSON.stringify(cm.getMemoText("A2"))} / 作者 ${JSON.stringify(cm.getMemoAuthor("A2"))}`);

// 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
cm.removeComment("D2");
cm.removeMemo("A2");
const tryRead = (label, fn) => {
    try { fn(); console.log(`    削除後 ${label}: 例外なし`); }
    catch (e) { console.log(`    削除後 ${label}: 例外 ${e.message}`); }
};
tryRead("D2コメント", () => cm.getComment("D2"));
tryRead("A2メモ", () => cm.getMemoText("A2"));
console.log(`    B2 は残すので save 後も ${cm.getComment("B2").length}件`);

const total = ws.getRange("B5:B5").getValue(true)["B5"];
const e1 = ws.getRange("E1:E1").getValue(true)["E1"];
console.log(`[確認] テーブル=${table.getName()} 集計(単価)=${total} 構造化参照(E1)=${e1}`);

wb.save();
wb.close();