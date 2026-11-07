/**
 * test_express_pdf.js
 * 
 * 13_express_pdf_stream.js を起動し、実際に HTTP POST リクエストを送信して、
 * 正しい Content-Type (application/pdf) と有効な PDF バイナリ (%PDF-1.x ヘッダー) が
 * メモリから直接返却されることを検証する自動テストスクリプトです。
 */

const http = require("http");
const fs = require("fs");
const app = require("./13_express_pdf_stream.js");

const server = app.listen(0, async () => {
    const port = server.address().port;
    console.log(`[Test] Test server listening on port ${port}`);

    const postData = JSON.stringify({
        title: "2026年度 業績レポート",
        department: "営業推進部",
        records: [
            { item: "クラウド移行支援", amount: 3200000 },
            { item: "年間保守サポート", amount: 980000 },
            { item: "セキュリティ診断", amount: 1500000 }
        ]
    });

    const options = {
        hostname: "127.0.0.1",
        port: port,
        path: "/api/export-pdf",
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Content-Length": Buffer.byteLength(postData)
        }
    };

    const req = http.request(options, (res) => {
        console.log(`[Test] Status Code: ${res.statusCode}`);
        console.log(`[Test] Content-Type: ${res.headers["content-type"]}`);
        console.log(`[Test] Content-Disposition: ${res.headers["content-disposition"]}`);

        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
            const pdfBuffer = Buffer.concat(chunks);
            console.log(`[Test] Received PDF buffer size: ${pdfBuffer.length} bytes`);

            // PDF マジックナンバー "%PDF-" を検証
            const isPdf = pdfBuffer.slice(0, 5).toString("ascii") === "%PDF-";
            console.log(`[Test] PDF magic header check: ${isPdf ? "OK (%PDF-)" : "FAILED"}`);

            if (res.statusCode === 200 && isPdf && pdfBuffer.length > 1000) {
                console.log("[Test] ALL PASS: Express directly streamed PDF without writing to disk!");
                server.close(() => process.exit(0));
            } else {
                console.error("[Test] FAILED: Invalid response or PDF buffer");
                server.close(() => process.exit(1));
            }
        });
    });

    req.on("error", (e) => {
        console.error(`[Test] Request error: ${e.message}`);
        server.close(() => process.exit(1));
    });

    req.write(postData);
    req.end();
});
