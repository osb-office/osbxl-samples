// 05_rows_cols.wasm.js — WebAssembly (wasmosbxl) sample
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
dump("[初期]       A列:", "A1:A5");

// ---------- 行 ----------
ws.insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow]  A列:", "A1:A7");
ws.deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow]  A列:", "A1:A5");
ws.clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
dump("[clearRow]   A列:", "A1:A5");
ws.copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow]    A列:", "A1:A5");

fill();
// ---------- 列 ----------
ws.insertCol(2, 1);                    // B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1");
ws.deleteCol(2, 1);                    // 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1");
ws.clearCol(3, 1);                     // C 列をクリア
dump("[clearCol]  1行目:", "A1:C1");
fill();
ws.copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
dump("[copyCol]   1行目:", "A1:F1");

// ---------- セル範囲 ----------
ws.insertCell("B2:B3", enums.XlInsertDirection.InsertDirectionDown);
dump("[insertCell] B列:", "B1:B5");
ws.deleteCell("B2:B3", enums.XlInsertDirection.DeleteDirectionUp);
dump("[deleteCell] B列:", "B1:B5");
ws.clearCell("A1:B2");
dump("[clearCell] A1:B2:", "A1:B2");
ws.copyCell("Sheet1", "A3:B4", "F8");
dump("[copyCell]  F8:G9:", "F8:G9");

wb.save();
wb.close();
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
