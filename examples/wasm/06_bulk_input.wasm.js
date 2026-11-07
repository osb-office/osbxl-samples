// 06_bulk_input.wasm.js — WebAssembly (wasmosbxl) sample
const fs = require("fs");
const path = require("path");

function getModulePath(name) {
    const candidates = [
        path.join(__dirname, "node_modules", "wasmosbxl", name),
        path.join(__dirname, "..", "..", "node_modules", "wasmosbxl", name),
        path.join(process.cwd(), "node_modules", "wasmosbxl", name),
    ];
    for (const c of candidates) {
        if (fs.existsSync(c)) return c;
    }
    return name;
}

const createModule = require(getModulePath("osbxl_wasm.js"));
const { createOsbxl } = require(getModulePath("osbxl_em.js"));

createOsbxl(createModule).then(async ({ Module, osbxl, dto, chart, enums }) => {
    // Setup license if present
    const lic = process.env.OSB_LICENSE_PATH_WASM || "/Users/user/lic_wasm.osb";
    if (fs.existsSync(lic)) {
        try {
            Module.FS.mkdirTree(lic.substring(0, lic.lastIndexOf("/")));
            Module.FS.writeFile(lic, fs.readFileSync(lic));
        } catch(e) {}
    }

// bulk-input.wasm.js — 2次元配列による値と数式の一括高速投入 (wasmosbxl / embind)
const app = new osbxl.App();
const wb = app.createWorkBook("bulk-input.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
const data = [
    ["商品名", "単価", "数量", "金額"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
];
ws.setRangeData("A1", data);

console.log("[1] setRangeData 投入後:");
for (const a1 of ["A1", "B2", "C2", "D1"]) {
    console.log(`    ${a1} = ${JSON.stringify(ws.getRange(a1).getValue()[a1])}`);
}

// ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
const formulas = [
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
];
ws.setRangeFormulas("D2", formulas);

console.log("[2] setRangeFormulas 投入後:");
console.log("    数式 D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getFormula()));
console.log("    計算値 D2:D4 =", JSON.stringify(ws.getRange("D2:D4").getValue(true)));

// ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
const ROWS = 2000, COLS = 5;

// セルごとに getCells() して投入
const wsCell = wb.addWorkSheet("CellByCell");
let t = Date.now();
for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
}
const msCell = Date.now() - t;

// 2次元配列で setRangeData() 投入 (最高速)
const wsBulk = wb.addWorkSheet("Bulk");
t = Date.now();
const matrix = [];
for (let r = 1; r <= ROWS; r++) {
    const row = [];
    for (let c = 1; c <= COLS; c++) row.push(String(r * 100 + c));
    matrix.push(row);
}
const msBuild = Date.now() - t;
t = Date.now();
wsBulk.setRangeData("A1", matrix);
const msApply = Date.now() - t;

console.log(`[3] ${ROWS * COLS} セルの投入速度比較:`);
console.log(`    セルごと getCells().setNumberValue() : ${msCell} ms`);
console.log(`    setRangeData (構築${msBuild} + 投入${msApply}) : ${msBuild + msApply} ms (最速)`);

const a = wsCell.getRange("E2000:E2000").getValue(true)["E2000"];
const b = wsBulk.getRange("E2000:E2000").getValue(true)["E2000"];
console.log(`[検証] 両方式とも E2000 = ${a === Number(b) ? b : `不一致 ${a} / ${b}`}`);

wb.save();
wb.close();
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
