# osbxl-samples

> **[🇯🇵 日本語版の README はこちら (README.ja.md)](README.ja.md)**

Official verified code samples and patterns for **[osbxl (https://osb-office.com)](https://osb-office.com/)** — an AI-ready, ultra-fast C++ core multi-platform spreadsheet calculation and Excel generation engine.

[![Official Website](https://img.shields.io/badge/Official_Website-osb--office.com-blue?style=flat-square)](https://osb-office.com/)
[![AI-Ready](https://img.shields.io/badge/AI--Ready-Cursor%20%7C%20Copilot%20%7C%20Gemini-orange?style=flat-square)](https://osb-office.com/pages/ai-skill.php)
[![Languages](https://img.shields.io/badge/Languages-Python%20%7C%20Node.js%20%7C%20Java%20%7C%20PHP%20%7C%20Wasm-green?style=flat-square)](https://osb-office.com/pages/features.php)

---

## 🌟 Key Features

- **🤖 AI-Ready with Zero Hallucinations**: Bundled with official AI Skill packages tailored for **Cursor, GitHub Copilot, Gemini, and Claude Code**. AI generates verified, runnable Excel manipulation code on the first try.
- **⚡ Overwhelming C++ Speed**: Native C++ core engine processes large workbooks and complex formula calculations in milliseconds with minimal memory footprint.
- **🌐 5 Supported Platforms**: Consistent, clean APIs across **Python (pyosbxl)**, **Node.js (nodeosbxl)**, **Java (osbxl)**, **PHP (phposbxl)**, and **WebAssembly (wasmosbxl)**.
- **📊 Rich Spreadsheet Capabilities**: Nearly 500 Excel-compatible functions, charts (bar, line, etc.), AutoFilters, Tables (ListObjects), Conditional Formatting, and PivotTables.
- **🔒 Security**: Read and write password-protected, encrypted `.xlsx` workbooks.

---

## 📁 Repository Structure

```text
osbxl-samples/
├── README.md               # Overview and quick reference
├── ja/                     # 日本語版サンプル集 (HTML / Web形式)
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
└── en/                     # English Samples Collection
    ├── index.php           # Sample index and task-based guide
    ├── general-01-overview.php          # Overview, object model & basic lifecycle
    ├── general-02-values-formulas.php   # Values, formulas & relative references
    ├── general-03-sheets.php            # Worksheets operations
    ├── general-04-value-types.php       # Cell value types & formatting
    ├── general-05-rows-cols.php         # Rows, columns & ranges
    ├── general-06-bulk-input.php        # High-speed bulk input
    ├── general-07-autofilter-sort.php   # AutoFilter & sorting
    ├── general-08-tables-comments.php   # Tables (ListObjects) & comments
    ├── general-09-conditional-format.php# Conditional formatting
    ├── general-10-charts.php            # Charts creation
    └── general-11-pivottables.php       # PivotTables
```

---

## 🚀 Quick Code Examples

### Python (`pyosbxl`)

```python
import pyosbxl

app = pyosbxl.App()
book = app.createWorkBook("sales_report.xlsx")
sheet = book.addWorkSheet("Sales")

sheet.getRange("A1").setValue("Monthly Sales")
sheet.getRange("A2").setValue(150000)
sheet.getRange("A3").setValue(230000)
sheet.getRange("A4").setFormula("SUM(A2:A3)")

book.save()
book.close()
```

### Node.js (`nodeosbxl`)

```javascript
const osbxl = require("nodeosbxl");

const app = new osbxl.App();
const book = app.createWorkBook("report.xlsx");
const sheet = book.addWorkSheet("Sheet1");

sheet.getRange("A1").setNumberValue(100);
sheet.getRange("A2").setNumberValue(200);
sheet.getRange("A3").setFormula("SUM(A1:A2)");

book.save();
book.close();
```

---

## 🔗 Official Links & Documentation

- **Official Website**: [https://osb-office.com/](https://osb-office.com/)
- **AI Skill Guide & Download**: [https://osb-office.com/pages/ai-skill.php](https://osb-office.com/pages/ai-skill.php)
- **Interactive Web Samples**: [https://osb-office.com/sample/](https://osb-office.com/sample/)
- **Product Features**: [https://osb-office.com/pages/features.php](https://osb-office.com/pages/features.php)
- **Free Demo Download**: [https://osb-office.com/pages/download.php](https://osb-office.com/pages/download.php)

---

## 📄 License

- **Sample Code & AI Skills**: [MIT License](LICENSE) — Feel free to use, copy, and modify in your projects.
- **OSBXL Engine (Libraries)**: Commercial / Proprietary by osb-office.inc. See [LICENSE](LICENSE) or [LICENSE.ja.md](LICENSE.ja.md) for software license agreement details.

