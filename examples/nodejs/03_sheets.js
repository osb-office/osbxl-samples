// sheets.js — シートの追加・移動・リネーム・削除と、インデックスアクセス
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sheets.xlsx");

// インデックスは 0 始まり (zeroBase)。0..getSheetCount()-1 を順に開いて名前を並べる
const show = (label) => {
    const names = [];
    for (let i = 0; i < wb.getSheetCount(); i++) {
        names.push(wb.openWorkSheet(i).getName());
    }
    console.log(`${label} count=${wb.getSheetCount()} [${names.join(", ")}]`);
};

show("[1] 新規作成直後        :");

wb.addWorkSheet("集計");                       // position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :");

wb.addWorkSheet("受注", 0);                    // position=0 → 先頭に挿入 (0-based)
show("[3] addWorkSheet(受注,0):");

wb.moveSheet("受注", 2);                       // 3 番目へ移動 (0-based: index 2)
show("[4] moveSheet(受注,2)   :");

const ws = wb.openWorkSheet("受注");
console.log("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"));
ws.getRange("A1:A1").setValue("受注データ");

wb.openWorkSheet("Sheet1").setName("明細");     // リネーム
show("[6] setName(明細)       :");

// --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
console.log("[7] isActive: 集計 =", wb.openWorkSheet("集計").isActive(),
            "/ 明細 =", wb.openWorkSheet("明細").isActive());
wb.openWorkSheet("集計").activate(true);   // 他のシートのアクティブは自動的に解除されます
console.log("    activate後: 集計 =", wb.openWorkSheet("集計").isActive(),
            "/ 明細 =", wb.openWorkSheet("明細").isActive());

// --- 削除 ---
wb.addWorkSheet("一時");
wb.deleteWorkSheet("一時");
show("[8] deleteWorkSheet(一時):");

wb.save();
wb.close();

// --- 開き直して確認 ---
const wb2 = app.openWorkBook("sheets.xlsx");
const names = [];
for (let i = 0; i < wb2.getSheetCount(); i++) names.push(wb2.openWorkSheet(i).getName());
console.log("[終] シート構成:", names.join(", "));
console.log("     受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"]);
wb2.close();