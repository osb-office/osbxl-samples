// 12_pdf_export.wasm.js — WebAssembly (wasmosbxl) sample
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

// 注意: WebAssembly (wasmosbxl) ビルドでは現在 PDF 出力機能 (OSBXL_ENABLE_PDF) は未サポートです。
// PageSetupObject によるページ余白・用紙サイズ・印刷の向きの設定やブックの保存は WebAssembly 環境でも動作します。
// PDF エクスポートを実行する場合は Node.js / Python / Java / PHP の各版をご利用ください。
}).catch(err => {
    console.error("Execution error:", err);
    process.exit(1);
});
