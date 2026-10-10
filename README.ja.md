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


