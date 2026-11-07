// 11_pivottables.wasm.js — WebAssembly (wasmosbxl) sample
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

// pivot-sample.wasm.js — ピボットテーブル: node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("pivot-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. ソースデータ ----------
// 地域 / 商品 / 数量 の 10 行
const headers = ["地域", "商品", "数量"];
const srcData = [
  ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
  ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
  ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
  ["北", "りんご", 20],
];
headers.forEach((h, i) => {
  const col = String.fromCharCode(65 + i);
  ws.getRange(`${col}1:${col}1`).setValue(h);
});
srcData.forEach((row, ri) => {
  row.forEach((v, ci) => {
    const col = String.fromCharCode(65 + ci);
    const r = ri + 2;
    if (typeof v === "number") ws.getRange(`${col}${r}:${col}${r}`).setNumberValue(v);
    else ws.getRange(`${col}${r}:${col}${r}`).setValue(v);
  });
});
const srcTotal = srcData.reduce((s, r) => s + r[2], 0);
console.log(`[1] ソースデータ: ${srcData.length} 行, 数量の合計 = ${srcTotal}`);

// ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
const pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");

const rf = pt.getField("地域");
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow);
const df = pt.getField("数量");
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData);
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);

pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");

// ---------- 3. 集計結果の読み戻しと検算 ----------
// ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
const pivotRows = {};
for (let r = 2; r <= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined && label !== "") {
    pivotRows[label] = Number(val);
    console.log(`    ${label} = ${val}`);
  }
}
const grandTotal = pivotRows["総計"];
console.log(`[3] 検算: 総計 ${grandTotal} === ソース合計 ${srcTotal} → ${grandTotal === srcTotal ? "OK" : "NG"}`);

// ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df]);
pt.refreshData();
pt.refresh();
console.log("[4] 集計方法を合計→平均に変更");
for (let r = 2; r <= 10; r++) {
  const label = ws.getRange(`E${r}:E${r}`).getValue()[`E${r}`];
  const val   = ws.getRange(`F${r}:F${r}`).getValue()[`F${r}`];
  if (label !== undefined && label !== "") {
    console.log(`    ${label} = ${val}`);
  }
}

// ---------- 5. フィールド情報の取得 ----------
const fields = pt.getFields();
console.log(`[5] getFields: ${fields.length} 個 → ${fields.map(f => f.getCaption()).join(", ")}`);
console.log(`    getLayoutType = ${pt.getLayoutType()} (TabularForm=${enums.XlPivotTableLayoutType.TabularForm})`);

// ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計");
try {
  ws.getPivotTables().getPivotTable("売上集計");
  console.log("[6] 削除失敗");
} catch (e) {
  console.log("[6] removePivotTable → 削除確認OK");
}

wb.save();
wb.close();
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
