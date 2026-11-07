// values-formulas.js — ブックを作って、値と数式を入れて保存する
const { nodeosbxl, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("sales.xlsx");   // 新規ブック作成
const ws = wb.openWorkSheet("Sheet1");         // 既定のシートを開く

// --- 見出し (1行目) ---
// app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
["商品", "単価", "個数", "金額"].forEach((text, i) => {
    const cell = app.convertFromRowColNumber(1, i + 1);
    ws.getRange(`${cell}:${cell}`).setValue(text);
});

// --- 明細 (2〜4行目) ---
const data = [["りんご", 100, 3], ["みかん", 50, 6], ["ぶどう", 150, 2]];
data.forEach(([name, price, qty], i) => {
    const row = i + 2;
    ws.getRange(`A${row}:A${row}`).setValue(name);
    ws.getRange(`B${row}:B${row}`).setNumberValue(price);
    ws.getRange(`C${row}:C${row}`).setNumberValue(qty);
});

// --- 金額列 (D2:D4) ---
// 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
// 数式文字列に先頭の "=" は付けない。
const formulas = data.map((_, i) => [`B${i + 2}*C${i + 2}`]);
ws.setRangeFormulas("D2", formulas);

// --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計");
ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

// --- 読み戻し ---
// getValue(true) は「生値」= 数式の計算結果を返す
const values = ws.getRange("A1:D5").getValue(true);
console.log("金額 D2:D4 =", [values["D2"], values["D3"], values["D4"]].join(", "));
console.log("合計:", values["D5"]);

// --- 保存して閉じる ---
wb.save();
wb.close();