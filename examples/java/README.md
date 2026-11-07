# Spring Boot PDF ストリーミング直接出力サンプル (osbxl Java v2.0.0)

このサンプルは、**osbxl (Java) 2.0.0** の新機能である `wb.saveAsPdfToBuffer()` を活用し、**Spring Boot** Web サーバー上で一時ファイルをディスクに書き出すことなく、HTTP POST リクエストに対してインメモリで PDF を動的生成・直接ストリーミング返却する完全な実用例です。

## 特徴
- **ディスク I/O ゼロ**: `wb.saveAsPdfToBuffer()` により、一時ファイル生成やディスクアクセスのオーバーヘッドなしで高速に PDF バイナリ (`byte[]`) を生成します。
- **データテーブル & 集合縦棒グラフの同時描画**: 構造化テーブル（ListObjects）と `TotalsCalculationSum` 集計行、およびモダンなカラーパレットの集合縦棒グラフを同一ページ内に美しくレイアウト。
- **堅牢な PDF セキュリティ (AES-256)**: `PdfExportObjectWrapper` による閲覧パスワード保護（`customer2026`）、印刷・コピー制限、カラーモードおよび列自動収縮フィットを設定。
- **サーバーレス・コンテナ親和性**: 一時ファイル削除処理が不要なため、AWS ECS/Fargate, Google Cloud Run, Kubernetes, Docker コンテナ環境などで安全・スケーラブルに動作します。
- **PageSetupObject による精密な印刷設定**: A4・横向き (Landscape)・余白 (Margins)・水平/垂直センタリング・ヘッダー/フッターを自在に制御できます。
- **zeroBase 準拠**: osbxl v2.0.0 の既定である 0 始まりの行・列・シートインデックスに対応しています。

---

## 📷 出力結果プレビュー

Spring Boot サーバーから返却された PDF（開封パスワード: `customer2026`）の実際のレンダリング画像です：

![売上実績レポート PDF 出力プレビュー](../../../assets/images/sales_report_pdf_preview.png)

---

## プロジェクト構成
```text
spring-boot-pdf/
├── pom.xml
└── src/main/java/com/example/demo/
    └── Application.java  # Spring Boot アプリケーション & @RestController
```

---

## 実行方法

### 1. ビルド
```bash
mvn clean package -DskipTests
```

### 2. サーバーの起動
```bash
mvn spring-boot:run
```
サーバーが `http://localhost:8080` で起動します。

### 3. curl コマンドによる PDF 出力テスト
```bash
curl -X POST http://localhost:8080/api/export-pdf --output report_spring.pdf
```
実行すると、サーバーから直接返却された `report_spring.pdf` が保存され、A4 横向きで美しくレイアウトされた帳票 PDF が確認できます。

