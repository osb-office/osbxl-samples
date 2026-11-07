// 03_sheets.wasm.js — WebAssembly (wasmosbxl) sample
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

// sheets.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("sheets-wasm.xlsx");

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
const wb2 = app.openWorkBook("sheets-wasm.xlsx");
const names = [];
for (let i = 0; i < wb2.getSheetCount(); i++) names.push(wb2.openWorkSheet(i).getName());
console.log("[終] シート構成:", names.join(", "));
console.log("     受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"]);
wb2.close();
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
