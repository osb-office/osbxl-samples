// 04_value_types.wasm.js — WebAssembly (wasmosbxl) sample
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

// value-types.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("value-types-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

const show = (a1) => {
    const r = ws.getRange(`${a1}:${a1}`);
    console.log(`  ${a1}: 表示値=${JSON.stringify(r.getValue()[a1])}` +
                ` 生値=${JSON.stringify(r.getValue(true)[a1])}` +
                ` 書式=${JSON.stringify(r.getNumberFormat()[a1])}`);
};

// --- 日付・時刻 ---
// DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
const date = new dto.DateTimeObject();
date.setYMD(2024, 1, 15);

const datetime = new dto.DateTimeObject();
datetime.setYMD(2024, 1, 15);
datetime.setHMS(13, 45, 30);

const time = new dto.DateTimeObject();
time.setHMS(9, 5, 0);

console.log("[日付・時刻]");
ws.getRange("A1:A1").setDateValue(date);                      // 日付のみ
ws.getRange("A2:A2").setDateValue(datetime);                  // 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time);                      // 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15");        // 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");      // 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");  // 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語も使えます
["A1", "A2", "A3", "A4", "A5", "A6", "A7"].forEach(show);

// --- 数値・文字列 ---
console.log("[数値・文字列]");
ws.getRange("B1:B1").setNumberValue(1234567.891);
ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
ws.getRange("B4:B4").setValue("00123");      // 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15");  // 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, false, '"円"#,##0');   // 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン → [Red]
["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"].forEach(show);

// --- 真偽値 ---
console.log("[真偽値]");
ws.getRange("C1:C1").setBooleanValue(true);
ws.getRange("C2:C2").setBooleanValue(false);
ws.getRange("C3:C3").setFormula("TRUE");     // 数式でも真偽値になる
["C1", "C2", "C3"].forEach(show);

// --- 日付のシリアル値 ---
console.log("[シリアル値]");
console.log("  2024/1/15 のシリアル値 =", app.getNumericValue(date));
console.log("  1904 年形式の場合      =", app.getNumericValue(date, true));
console.log("  wb.isDate1904()        =", wb.isDate1904());

wb.save();
wb.close();
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
