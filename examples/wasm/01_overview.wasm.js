// 01_overview.wasm.js — WebAssembly (wasmosbxl) sample
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

// basic-flow.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();                        // [1] App を生成

const wb = app.createWorkBook("flow-wasm.xlsx");    // [2] 新規作成
const ws = wb.openWorkSheet("Sheet1");              // [3] ワークシートを開く

ws.getRange("A1:A1").setValue("新規作成");
ws.getRange("B1:B1").setFormula("100*3");           // [5] 数式（先頭の "=" は付けない）
console.log("[6] getValue(true) -> B1 =", ws.getRange("B1:B1").getValue(true)["B1"]);

wb.save();                                          // [7] 保存
wb.close();                                         // [8] 閉じる
console.log("[8] close() 完了");

const wb2 = app.openWorkBook("flow-wasm.xlsx");     // [2'] 既存のブックを開く
const ws2 = wb2.openWorkSheet("Sheet1");
console.log("[2'] openWorkBook -> A1 =", ws2.getRange("A1:A1").getValue(true)["A1"]);
wb2.close();
console.log("[8] close() 完了");
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
