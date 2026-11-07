<?php
// test_slim_pdf.php — Slim 4 PDF 直接ストリーミングのエンドツーエンド検証
$port = 8083;
echo "[1/3] PHP Slim サーバー (ポート {$port}) を起動中...\n";
$cmd = sprintf("php -S 127.0.0.1:%d %s", $port, escapeshellarg(__DIR__ . '/13_slim_pdf_stream.php'));
$descriptors = [
    0 => ["pipe", "r"],
    1 => ["pipe", "w"],
    2 => ["pipe", "w"]
];
$process = proc_open($cmd, $descriptors, $pipes);
usleep(1500000); // 1.5 秒待機

try {
    echo "[2/3] HTTP POST http://127.0.0.1:{$port}/api/export-pdf を送信...\n";
    $ctx = stream_context_create([
        'http' => [
            'method' => 'POST',
            'ignore_errors' => true
        ]
    ]);
    $res = file_get_contents("http://127.0.0.1:{$port}/api/export-pdf", false, $ctx);
    $headers = $http_response_header;

    $statusLine = $headers[0] ?? '';
    $contentType = '';
    foreach ($headers as $h) {
        if (stripos($h, 'Content-Type:') === 0) {
            $contentType = trim(substr($h, 13));
        }
    }

    echo "    ステータス: {$statusLine}\n";
    echo "    Content-Type: {$contentType}\n";
    echo "    取得バイト長: " . strlen($res) . " bytes\n";
    echo "    ヘッダー先頭: " . substr($res, 0, 5) . "\n";

    if (!str_contains($statusLine, '200')) {
        throw new Exception("HTTP status is not 200: {$statusLine}");
    }
    if (!str_contains($contentType, 'application/pdf')) {
        throw new Exception("Content-Type is not application/pdf: {$contentType}");
    }
    if (!str_starts_with($res, "%PDF-")) {
        throw new Exception("Data does not start with %PDF-");
    }

    echo "[3/3] 検証成功: PHP (Slim 4) によるインメモリ PDF 直接ストリーミングが完全に動作しました！\n";
} finally {
    proc_terminate($process);
    proc_close($process);
}
