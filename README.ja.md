# osbxl サンプル集 (osbxl-samples)

[English README is available here](README.md)

**[osbxl 公式サイト (https://osb-office.com)](https://osb-office.com/)** の公式サンプルコード・パターン集です。  
osbxl は、AI によるコード自動生成に対応した、C++ 製コアによる超高速マルチプラットフォーム・スプレッドシート計算＆Excel 生成エンジンです。

[![公式サイト](https://img.shields.io/badge/公式サイト-osb--office.com-blue?style=flat-square)](https://osb-office.com/)
[![ライブラリ比較](https://img.shields.io/badge/性能比較-vs%20POI%20%7C%20openpyxl%20%7C%20PhpSpreadsheet-red?style=flat-square)](https://osb-office.com/pages/comparison.php)
[![AI対応](https://img.shields.io/badge/AI最適化-Cursor%20%7C%20Copilot%20%7C%20Gemini-orange?style=flat-square)](https://osb-office.com/pages/ai-skill.php)
[![対応言語](https://img.shields.io/badge/対応言語-Python%20%7C%20Node.js%20%7C%20Java%20%7C%20PHP%20%7C%20Wasm-green?style=flat-square)](https://osb-office.com/pages/features.php)

---

## 📊 主要Excelライブラリとの違い・性能比較

> **詳細な比較表・移行ガイドは [公式サイトのライブラリ比較ページ](https://osb-office.com/pages/comparison.php) をご覧ください。**

従来の高級言語ライブラリ（Apache POI、openpyxl、PhpSpreadsheet 等）で大量データを扱う際の**「メモリ不足（Memory Exhausted）」「GC による速度低下」「計算エンジンの不在」**を、C++ ネイティブコアにより根本から解決します。

| 項目 | **osbxl** | **Apache POI (Java)** | **openpyxl (Python)** | **PhpSpreadsheet (PHP)** | **SheetJS / ExcelJS (JS)** |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **コア実装** | **C++ ネイティブ** | Java | Python | PHP | JavaScript |
| **大量データ処理** | **◎ 圧倒的に高速** | ◯ (SXSSFストリーム) | △ (行数増加で遅延) | ✕ (数万行で極めて遅い) | ◯ / △ |
| **メモリ消費** | **◎ 最小限（最大90%削減）** | △ (ヒープ大量消費) | △ (オブジェクト過多) | ✕ (メモリ制限で500エラー) | △ (V8ヒープ制限) |
| **Excel計算式評価** | **約500関数（XLOOKUP等）** | 300〜400関数 | ✕ (数式評価不可) | △ (一部のみ・重い) | 限定的 / 有償版 |
| **AIコード生成** | **◎ 公式AI Skill同梱** | なし (ハルシネーション頻発)| なし | なし | なし |
| **WebAssembly** | **◎ 公式対応 (Wasm 3.0)** | ✕ | △ (巨大ランタイム要) | ✕ | ◯ |

---

## 🌟 主な特徴

- **🤖 AIコーディング完全対応（ハルシネーション・嘘のコードゼロ）**:  
  **Cursor / GitHub Copilot / Gemini / Claude Code** 向けの公式 AI Skill パッケージを同梱。AI が厳格な API 仕様や落とし穴（Pitfalls）を自動認識し、1発で正常動作する Excel 操作コードを自動生成します。
- **⚡ C++ コアによる圧倒的な計算速度**:  
  ネイティブ C++ エンジンにより、大規模ワークブックや複雑な数式計算もミリ秒単位で超高速処理。メモリ消費も最小限に抑えます。
- **🌐 5つのプラットフォームにネイティブ対応**:  
  **Python (pyosbxl)** / **Node.js (nodeosbxl)** / **Java (osbxl)** / **PHP (phposbxl)** / **WebAssembly (wasmosbxl)** で共通の直感的な API 設計。
- **📊 充実した Excel 互換機能**:  
  Excel 互換の 500 に近い関数、グラフ・チャート（縦棒・折れ線等）、オートフィルター、テーブル（ListObjects・集計行）、条件付き書式、ピボットテーブルをサポート。
- **🔒 セキュリティ**:  
  パスワード付き・暗号化された `.xlsx` ファイルの読み込みおよび生成に対応。

---

## 📁 ディレクトリ構成

```text
osbxl-samples/
├── README.md               # 英語版 README (English)
├── README.ja.md            # 日本語版 README (本ファイル)
├── LICENSE                 # ライセンス条項 (英語版)
├── LICENSE.ja.md           # ライセンス条項 (日本語版)
├── examples/               # 実行可能コードサンプル
│   ├── python/             # Python サンプル
│   │   ├── quickstart.py             # 基本操作 & 数式
│   │   ├── chart_sample.py           # 縦棒 & 折れ線グラフ作成
│   │   ├── table_comments_sample.py  # テーブル (ListObjects) & コメント
│   │   └── conditional_format_sample.py # 条件付き書式 & データバー
│   ├── nodejs/             # Node.js サンプル
│   │   ├── quickstart.js             # 基本操作 & 数式
│   │   ├── chart_sample.js           # 縦棒 & 折れ線グラフ作成
│   │   └── table_comments_sample.js  # テーブル (ListObjects) & コメント
│   ├── java/               # Java サンプル (QuickStart.java)
│   ├── php/                # PHP サンプル (quickstart.php)
│   └── wasm/               # WebAssembly サンプル (quickstart.wat, quickstart.wasm.js)
├── ja/                     # 日本語版 Web 形式リファレンス (general-01 ~ 11)
└── en/                     # 英語版 Web 形式リファレンス (general-01 ~ 11)
```

---

## 📊 ライブラリ別 総合機能比較 ＆ 実測パフォーマンス

実務のExcel自動生成で重要となる高度な機能（数式評価、ピボットテーブル、チャート、パスワード暗号化など）について、オープンソース主要ライブラリと **osbxl** の対応状況および実測性能です。

### 1. 機能対応マトリクス

| 機能項目 | osbxl (5言語共通) | Apache POI (Java) | openpyxl (Python) | PhpSpreadsheet (PHP) | ExcelJS / SheetJS (JS) |
|---|:---:|:---:|:---:|:---:|:---:|
| **.xlsx 読み書き** | **◎ 高速・安定** | **◎** | **◎** | **◎** | **◎** |
| **ヘッドレス数式評価**<br>(Excel非起動下での値計算) | **◎ 約500関数**<br>(XLOOKUP等含む) | **◯ 300〜400関数**<br>(Evaluator) | **✕ 非対応**<br>(Excel起動までNone) | **△ 一部のみ**<br>(重厚・関数限定) | **✕ 非対応**<br>(SheetJS無償版不可) |
| **テーブル / 構造化参照**<br>(`SUM(Table[列名])` 等) | **◎ 完全対応・ネイティブ評価** | **△ 計算不可**<br>(保存のみ) | **✕ 評価不可** | **✕ 評価不可** | **✕ 評価不可** |
| **オートフィルター** | **◎ 高度フィルター対応** | **◯** | **◯** | **◯** | **◯** |
| **ピボットテーブル作成** | **◎ 標準API対応** | **◯ 低レイヤXML** | **△ 読み取り中心** | **✕ 非対応** | **✕ 非対応** |
| **チャート・グラフ生成** | **◎ 2D / 3D 多数** | **◯ XDDF経由** | **◯ 2D/3D対応** | **△ 描画制約あり** | **✕〜△ 限定的** |
| **パスワード暗号化 / 復号** | **◎ ビルトイン対応** | **◯ 標準対応** | **✕ 外部ツール要** | **✕ 拡張モジュール要** | **✕ 非対応** |
| **コメント (新旧スレッド/メモ)** | **◎ スレッド・メモ両対応** | **◯ メモ対応** | **◯ コメント対応** | **◯ コメント対応** | **◯ コメント対応** |
| **AIコード生成最適化** | **◎ 専用AI Skill提供**<br>(捏造・ハルシネーション0) | **✕ 非公式**<br>(架空API誤生成多発) | **✕ 非公式** | **✕ 非公式** | **✕ 非公式** |

### 2. 実務の標準規模「200行（800セル）」実測ベンチマーク

業務システムで最も頻出する帳票サイズ（**200行 × 4列 = 計800セル + SUM総合計数式**）を開発環境（Linux x86_64）で実測比較しました。

| 言語 | ライブラリ | 処理時間 (秒) | メモリ消費 (MB) | 特徴・評価 |
|---|---|:---:|:---:|---|
| **Java** (JDK 1.8) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | 0.622 秒<br>**0.139 秒** | 15.19 MB<br>**16.87 MB** | **osbxl が約 4.5 倍高速**。<br>POIは巨大スキーマの初期化が重く、200行では osbxl が圧倒的に軽快。 |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.059 秒<br>**0.035 秒** | 109.07 MB<br>**125.61 MB** | **nodeosbxl がExcelJSを上回る0.03秒台（約1.7倍高速）**。<br>同等の書き出し条件でExcelJSを圧倒し、V8ヒープ消費もわずか18MB。 |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | **0.018 秒**<br>0.032 秒 | 53.73 MB<br>70.84 MB | openpyxlと同等の書き出し専用モードで**0.03秒台の超高速生成**。<br>Pythonヒープ消費はわずか12MB。 |
| **PHP** (8.5) | PhpSpreadsheet 5.10.0<br>**phposbxl** | 0.117 秒<br>**0.122 秒** | 26.00 MB<br>**2.00 MB** | 処理時間は0.1秒台で互角だが、**メモリは phposbxl が 13 倍省メモリ（2MB）**。<br>Webサーバーでの並行ダウンロード時に致命的な差となります。 |

### 3. 大規模データ限界負荷テスト「50,000行（200,000セル）」実測

データ規模を **50,000行 × 4列 = 計200,000セル＋SUM集計** まで拡大した限界テスト結果です。既存ライブラリがメモリ枯渇（OOM）でクラッシュする中、osbxl は全言語で完走します。

| 言語 / 環境 | ライブラリ | 処理時間 (秒) | メモリ消費 (MB) | 実測結果・限界検証 |
|---|---|:---:|:---:|---|
| **PHP** (8.5 / 128M) | PhpSpreadsheet 5.10.0<br>**phposbxl** | ✕ クラッシュ<br>**5.941 秒** | ✕ 128MB上限突破 (OOM)<br>**32.00 MB** | **PhpSpreadsheet はメモリ上限超過で即死**（`Fatal error: Allowed memory size exhausted`）。<br>**phposbxl はわずか 32MB で安全完走**（ファイルサイズ 1.93MB）。 |
| **Java** (JDK 1.8 / 256M) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | ✕ クラッシュ<br>**12.813 秒** | ✕ 256MBヒープ枯渇 (OOM)<br>**42.13 MB** | **POI はヒープ枯渇で即死**（`OutOfMemoryError: Java heap space`）。<br>**osbxl は 42MB の軽量ヒープで完走**（SXSSFのような機能制限書き換え不要）。 |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 1.165 秒<br>**2.329 秒** | 766.54 MB<br>**1,799 MB**<br>*(V8ヒープ: 20.9MB)* | 同等の書き出し条件（setOnlyMode）で2.3秒台を達成。<br>ExcelJSがV8ヒープを大量消費（420MB超）するのに対し、**nodeosbxlのV8ヒープ消費はわずか20.9MBで安全**。 |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 2.470 秒<br>**2.433 秒** | 378.28 MB<br>**1,439 MB**<br>*(Pythonヒープ: 28MB)* | 同等の書き出し条件（setOnlyMode）で**openpyxlを上回る2.4秒台（互角以上）を達成**。<br>Pythonヒープ消費はわずか28MB。 |

> 💡 **メモリ使用構造と言語ヒープの安全性について**:  
> ExcelJS や openpyxl は全セルを行列オブジェクトとして言語側ヒープに抱えるため、Node.js や Python のヒープ上限を急速に圧迫します。<br>
> 一方、osbxl はデータを2次元配列から即座に C++ エンジンへ引き渡すため、**Node.js 側の V8 ヒープ消費は 50,000 行でもわずか 20.9 MB、Python 側も 28 MB** に抑えられ、言語ランタイムの OOM（ヒープ枯渇）クラッシュを起こしません。プロセス全体の Peak RSS は、20万セルの XML 並列生成と ZIP ストリーム圧縮バッファを C++ コアで高速処理するために一時的に確保されるネイティブ作業領域です。  
> 📖 **詳細な機能マトリクスと考察**: [Excelライブラリ 総合機能比較 ＆ ベンチマークレポート (ja/benchmark-comparison.md)](ja/benchmark-comparison.md) をご覧ください。

---

## 🚀 クイックコード例

### Python (`pyosbxl`)

```python
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sales_report.xlsx")
ws = wb.openWorkSheet("Sheet1")

ws.getRange("A1").setValue("売上レポート")
ws.getRange("A2").setValue(150000)
ws.getRange("A3").setValue(230000)
ws.getRange("A4").setFormula("SUM(A2:A3)")

wb.save()
wb.close()
```

### Node.js (`nodeosbxl`)

```javascript
const osbxl = require("nodeosbxl");

const app = new osbxl.App();
const wb = app.createWorkBook("report.xlsx");
const ws = wb.addWorkSheet("Sheet1");

ws.getRange("A1").setNumberValue(100);
ws.getRange("A2").setNumberValue(200);
ws.getRange("A3").setFormula("SUM(A1:A2)");

wb.save();
wb.close();
```

---

## 🔗 公式リンク・リソース

- **公式サイト**: [https://osb-office.com/](https://osb-office.com/)
- **Excelライブラリ 性能・機能比較**: [https://osb-office.com/pages/comparison.php](https://osb-office.com/pages/comparison.php)
- **osbxl対著名ライブラリ詳細対比（200行 & 5万行）**: [https://osb-office.com/pages/members/benchmark_report.php](https://osb-office.com/pages/members/benchmark_report.php)
- **負荷テスト実測レポート**: [https://osb-office.com/pages/members/tips.php](https://osb-office.com/pages/members/tips.php)
- **公式 AI Skill ガイド & ダウンロード**: [https://osb-office.com/pages/ai-skill.php](https://osb-office.com/pages/ai-skill.php)
- **Web サンプル集（動作確認済み）**: [https://osb-office.com/sample/](https://osb-office.com/sample/)
- **製品特徴**: [https://osb-office.com/pages/features.php](https://osb-office.com/pages/features.php)
- **無料デモ版ダウンロード**: [https://osb-office.com/pages/download.php](https://osb-office.com/pages/download.php)

---

## 📄 ライセンス

- **サンプルコード & AI スキル**: [MIT ライセンス (LICENSE)](LICENSE) — ご自身のプロジェクトで商用・非商用問わず自由にご利用・改変いただけます。
- **OSBXL 本体（各言語ライブラリ）**: 株式会社OSBオフィス 商用ソフトウェア。利用条件については [LICENSE.ja.md](LICENSE.ja.md) または [LICENSE](LICENSE) をご確認ください。


