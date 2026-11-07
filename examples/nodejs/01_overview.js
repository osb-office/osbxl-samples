// basic-flow.js — 新規作成 → 編集 → 保存 → 開き直し → 別名保存 の一連の流れ
const { nodeosbxl } = require("nodeosbxl");

const app = new nodeosbxl.App();                  // [1] App を生成

// ---------- 新規ブックを作って編集し、保存する ----------
let wb = app.createWorkBook("flow.xlsx");         // [2] 新規作成
console.log("[2] createWorkBook -> getBookPath =", wb.getBookPath());

let ws = wb.openWorkSheet("Sheet1");              // [3] ワークシートを開く
console.log("[3] openWorkSheet  -> シート名 =", ws.getName());

// [4] Range を取得して [5] 値を設定
//     数値(行, 列)で取る ws.getCells(0, 0) も getRange("A1:A1") と同じ Range を返します (0-based)
ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100*3");         // [5] 数式を設定（先頭の "=" は付けない）
console.log("[6] getValue(true) -> B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);

wb.save();                                        // [7] 保存
console.log("[7] save() 完了");
wb.close();                                       // [8] 閉じる（WorkSheet の close は不要）
console.log("[8] close() 完了");

// ---------- 既存ブックを開いて、別名で保存する ----------
wb = app.openWorkBook("flow.xlsx");               // [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1");                  // [3']
const read = ws.getRange("A1:B1").getValue(true);
console.log("[2'] openWorkBook -> A1 =", read["A1"], "/ B1 =", read["B1"]);

ws.getRange("A2:A2").setValue("開き直して追記");
wb.saveAs("flow-copy.xlsx");                      // [7'] 別名で保存
console.log("[7'] saveAs() 完了 -> getBookPath =", wb.getBookPath());
wb.close();                                       // [8']
console.log("[8] close() 完了");