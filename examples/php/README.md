# PHP (Slim 4) PDF ストリーミング直接出力サンプル (phposbxl v2.0.0)

このサンプルは、**phposbxl 2.0.0** の新機能である `wb.saveAsPdfToBuffer()` を活用し、**Slim 4** マイクロフレームワーク上で一時ファイルをディスクに書き出すことなく、HTTP POST リクエストに対してインメモリで PDF を動的生成・直接ストリーミング返却する完全な実用例です。

## 特徴
- **ディスク I/O ゼロ**: `wb.saveAsPdfToBuffer()` により、一時ファイル生成やディスクアクセスのオーバーヘッドなしで高速に PDF バイナリ (`string`) を生成します。
- **データテーブル & 集合縦棒グラフの同時描画**: 構造化テーブル（ListObjects）と `TotalsCalculationSum` 集計行、およびモダンなカラーパレットの集合縦棒グラフを同一ページ内に美しくレイアウト。
- **堅牢な PDF セキュリティ (AES-256)**: `dto.PdfExportObject` による閲覧パスワード保護（`customer2026`）、印刷・コピー制限、カラーモードおよび列自動収縮フィットを設定。
- **サーバーレス・コンテナ親和性**: 一時ファイル削除処理が不要なため、AWS Lambda (Bref), Google Cloud Run, Docker コンテナ環境などで安全・スケーラブルに動作します。
- **PageSetupObject による精密な印刷設定**: A4・横向き (Landscape)・余白 (Margins)・水平/垂直センタリング・ヘッダー/フッターを自在に制御できます。
- **zeroBase 準拠**: osbxl v2.0.0 の既定である 0 始まりの行・列・シートインデックスに対応しています。

---

## 📷 出力結果プレビュー

Slim 4 サーバーから返却された PDF（開封パスワード: `customer2026`）の実際のレンダリング画像です：

![売上実績レポート PDF 出力プレビュー](../../assets/images/sales_report_pdf_preview.png)

---

## ファイル構成
- **[`13_slim_pdf_stream.php`](13_slim_pdf_stream.php)**: Slim 4 サーバー実装本体。POST `/api/export-pdf` で PDF を直接返却します。
- **[`test_slim_pdf.php`](test_slim_pdf.php)**: 実際にサーバーを起動して HTTP POST を送信し、HTTP 200、`Content-Type: application/pdf`、および有効な PDF バイナリ（`%PDF-` ヘッダー）が返却されることを検証する自動テストスクリプト。

---

## 実行方法

### 1. 依存ライブラリのインストール
```bash
composer require slim/slim slim/psr7
```

### 2. サーバーの起動 (PHP ビルトインサーバー)
```bash
php -S localhost:8080 13_slim_pdf_stream.php
```
サーバーが `http://localhost:8080` で起動します。

### 3. curl コマンドによる PDF 出力テスト
```bash
curl -X POST http://localhost:8080/api/export-pdf --output report_slim.pdf
```
実行すると、サーバーから直接返却された `report_slim.pdf` が保存され、A4 横向きで美しくレイアウトされた帳票 PDF が確認できます。

### 4. 自動テストの実行
```bash
php test_slim_pdf.php
```

