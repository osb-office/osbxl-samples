"""test_fastapi_pdf.py — FastAPI PDF 直接ストリーミングのエンドツーエンド検証
"""
import sys
import subprocess
import time
import urllib.request

def run_test():
    print("[1/3] FastAPI サーバー (ポート 8081) を起動中...")
    server = subprocess.Popen(
        [sys.executable, "-c", """
import uvicorn
import importlib.util
spec = importlib.util.spec_from_file_location("stream_app", "13_fastapi_pdf_stream.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
uvicorn.run(mod.app, host='127.0.0.1', port=8081, log_level='error')
"""],
        cwd="/Users/user/phposbxl-test/examples/python"
    )

    try:
        time.sleep(1.5)
        print("[2/3] HTTP POST http://127.0.0.1:8081/api/export-pdf を送信...")
        req = urllib.request.Request("http://127.0.0.1:8081/api/export-pdf", method="POST")
        with urllib.request.urlopen(req) as res:
            status = res.status
            content_type = res.headers.get("Content-Type")
            data = res.read()

        print(f"    ステータス: {status} (期待値: 200)")
        print(f"    Content-Type: {content_type} (期待値: application/pdf)")
        print(f"    取得バイト長: {len(data)} bytes")
        print(f"    ヘッダー先頭: {data[:5]}")

        assert status == 200, f"Expected 200, got {status}"
        assert content_type == "application/pdf", f"Expected application/pdf, got {content_type}"
        assert data.startswith(b"%PDF-"), "Header must start with %PDF-"

        print("[3/3] 検証成功: FastAPI によるインメモリ PDF 直接ストリーミングが完全に動作しました！")
    finally:
        server.terminate()
        server.wait()

if __name__ == "__main__":
    run_test()
