# osbxl サンプル集 (osbxl-samples)

[English README is available here](README.md)

**[osbxl 公式サイト (https://osb-office.com)](https://osb-office.com/)** の公式サンプルコード・パターン集です。  
osbxl は、AI によるコード自動生成に対応した、C++ 製コアによる超高速マルチプラットフォーム・スプレッドシート計算＆Excel 生成エンジンです。

[![公式サイト](https://img.shields.io/badge/公式サイト-osb--office.com-blue?style=flat-square)](https://osb-office.com/)
[![AI対応](https://img.shields.io/badge/AI最適化-Cursor%20%7C%20Copilot%20%7C%20Gemini-orange?style=flat-square)](https://osb-office.com/pages/ai-skill.php)
[![対応言語](https://img.shields.io/badge/対応言語-Python%20%7C%20Node.js%20%7C%20Java%20%7C%20PHP%20%7C%20Wasm-green?style=flat-square)](https://osb-office.com/pages/features.php)

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
├── examples/               # 実行可能コードサンプル
│   ├── python/             # Python サンプル (quickstart.py)
│   ├── nodejs/             # Node.js サンプル (quickstart.js)
│   ├── java/               # Java サンプル (QuickStart.java)
│   ├── php/                # PHP サンプル (quickstart.php)
│   └── wasm/               # WebAssembly サンプル (quickstart.wat, quickstart.wasm.js)
├── ja/                     # 日本語版 詳細ドキュメント・Webサンプル集
│   ├── index.php           # サンプル一覧・やりたいこと別インデックス
│   ├── general-01-overview.php          # 概要・オブジェクトモデル・基本の流れ
│   ├── general-02-values-formulas.php   # 値と数式・相対参照
│   ├── general-03-sheets.php            # ワークシート操作（追加・削除・リネーム）
│   ├── general-04-value-types.php       # セル値の型（文字列・数値・日付・時刻）
│   ├── general-05-rows-cols.php         # 行・列・セル範囲（挿入・クリア・コピー）
│   ├── general-06-bulk-input.php        # 大量データの一括高速投入
│   ├── general-07-autofilter-sort.php   # オートフィルター・並べ替え
│   ├── general-08-tables-comments.php   # テーブル（ListObjects）・集計行・コメント
│   ├── general-09-conditional-format.php# 条件付き書式（カラースケール・データバー）
│   ├── general-10-charts.php            # グラフ（縦棒・折れ線）作成
│   └── general-11-pivottables.php       # ピボットテーブル集計
└── en/                     # 英語版 詳細ドキュメント・Webサンプル集
```

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
- **公式 AI Skill ガイド & ダウンロード**: [https://osb-office.com/pages/ai-skill.php](https://osb-office.com/pages/ai-skill.php)
- **Web サンプル集（動作確認済み）**: [https://osb-office.com/sample/](https://osb-office.com/sample/)
- **製品特徴**: [https://osb-office.com/pages/features.php](https://osb-office.com/pages/features.php)
- **無料デモ版ダウンロード**: [https://osb-office.com/pages/download.php](https://osb-office.com/pages/download.php)

---

## 📄 ライセンス

- **サンプルコード & AI スキル**: [MIT ライセンス (LICENSE)](LICENSE) — ご自身のプロジェクトで商用・非商用問わず自由にご利用・改変いただけます。
- **OSBXL 本体（各言語ライブラリ）**: 株式会社OSBオフィス 商用ソフトウェア。利用条件については [LICENSE.ja.md](LICENSE.ja.md) または [LICENSE](LICENSE) をご確認ください。


