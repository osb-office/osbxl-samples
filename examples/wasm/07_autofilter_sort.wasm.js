// 07_autofilter_sort.wasm.js — WebAssembly (wasmosbxl) sample
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

// af-sort.wasm.js — node 版サンプルと同一検証 (wasmosbxl / embind)
// 注意（WebAssembly 版）: モジュールは createOsbxl(createModule) で先に読み込みます。
// 下記で使う osbxl / dto / chart / enums はその時点で得られる4つの namespace です
// （Node.js 版の require("nodeosbxl") に相当）。
const app = new osbxl.App();
const wb = app.createWorkBook("af-sort-wasm.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// ---------- 1. データ準備（ヘッダ + 10 行）----------
const DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
];
["商品名", "地域", "数量", "単価"].forEach((h, i) =>
    ws.getRange(`${"ABCD"[i]}1:${"ABCD"[i]}1`).setValue(h));
DATA.forEach((row, i) => {
    const r = i + 2;
    ws.getRange(`A${r}:A${r}`).setValue(row[0]);
    ws.getRange(`B${r}:B${r}`).setValue(row[1]);
    ws.getRange(`C${r}:C${r}`).setNumberValue(row[2]);
    ws.getRange(`D${r}:D${r}`).setNumberValue(row[3]);
});
console.log("[1] データ準備: A1:D11（ヘッダ + 10 行）");

// 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
const text = (r, col) => ws.getRange(`${col}${r}:${col}${r}`).getValue()[`${col}${r}`];
const label = (r, col) => `${text(r, "A")}(${text(r, col)})`;
const visible = (col) => {
    const out = [];
    for (let r = 2; r <= 11; r++) if (!ws.getRow(r).getHidden()) out.push(label(r, col));
    return out.join(" ");
};
const hiddenList = () => {
    const out = [];
    for (let r = 2; r <= 11; r++) if (ws.getRow(r).getHidden()) out.push(r);
    return out.length ? out.join(",") : "なし";
};

// ---------- 2. オートフィルター設定 ----------
const af = ws.getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
console.log(`[2] getAutoFilter → 範囲 = ${af.getAddress()} / 非表示行 = ${hiddenList()}`);

// ---------- 3. 条件フィルター: 数量 > 100 ----------
// colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(enums.XlAutoFilterOperator.GreaterThan, "100", 3);
console.log(`[3] 数量>100 の可視行: ${visible("C")}`);
console.log(`    非表示行 = ${hiddenList()} / getValue("A2:A11") の件数 = ${Object.keys(ws.getRange("A2:A11").getValue()).length}`);
af.resetAllFilter();                            // 条件だけ解除（フィルター自体は残る）

// ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3);                   // 数量の上位 3 件
console.log(`[4] 数量上位3件の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesNumberFilter([120, 40], 3);   // 数量が 120 または 40
console.log(`[5] 数量∈{120,40} の可視行: ${visible("C")}`);
af.resetAllFilter();
af.setUniqueValuesStringFilter(["東京", "福岡"], 2);   // 地域が 東京 または 福岡（文字列リスト）
console.log(`[6] 地域∈{東京,福岡} の可視行: ${visible("B")}`);
af.resetAllFilter();
af.setCustomFilter(enums.XlAutoFilterOperator.Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
console.log(`[7] 地域="大阪" の可視行: ${visible("B")}`);
af.resetAllFilter();

// ---------- 5. 並べ替え（単一キー）----------
// ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
ws.getSort("A2:D11").executeSortAscending(2);   // 数量（範囲内3列目 (0-based: 2)）で昇順
console.log(`[8] 数量の昇順: ${visible("C")}`);

// オートフィルター経由の Sort はヘッダ行が自動的に保護される
const sort = af.getSort();
sort.executeSortDescending(3);                  // 単価（4列目 (0-based: 3)）で降順
console.log(`[9] 単価の降順: ${visible("D")}`);

// ---------- 6. 並び替え条件の読み戻し ----------
const cond = sort.getSortConditions()[0];
console.log(`[10] 条件の読み戻し: 対象列 = ${cond.getTarget()} / 昇順 = ${cond.getSortAscending()}`);

// ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter();
console.log(`[11] removeFilter 後: 非表示行 = ${hiddenList()}`);

wb.save();
wb.close();
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
