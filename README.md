# osbxl-samples

> **[🇯🇵 日本語版の README はこちら (README.ja.md)](README.ja.md)**

Official verified code samples and patterns for **[osbxl (https://osb-office.com)](https://osb-office.com/)** — an AI-ready, ultra-fast C++ core multi-platform spreadsheet calculation and Excel generation engine.

[![Official Website](https://img.shields.io/badge/Official_Website-osb--office.com-blue?style=flat-square)](https://osb-office.com/)
[![Library Comparison](https://img.shields.io/badge/Benchmark-vs%20POI%20%7C%20openpyxl%20%7C%20PhpSpreadsheet-red?style=flat-square)](https://osb-office.com/en/pages/comparison.php)
[![AI-Ready](https://img.shields.io/badge/AI--Ready-Cursor%20%7C%20Copilot%20%7C%20Gemini-orange?style=flat-square)](https://osb-office.com/pages/ai-skill.php)
[![Languages](https://img.shields.io/badge/Languages-Python%20%7C%20Node.js%20%7C%20Java%20%7C%20PHP%20%7C%20Wasm-green?style=flat-square)](https://osb-office.com/pages/features.php)

---

## 📊 Comparison with Popular Excel Libraries

> **For in-depth analysis and migration tips, check out our [Online Library Comparison Guide](https://osb-office.com/en/pages/comparison.php).**

Pure high-level libraries (Apache POI, openpyxl, PhpSpreadsheet) often suffer from **Out of Memory errors, GC pauses, and absent formula calculation engines** on large datasets. osbxl’s optimized C++ core eliminates these bottlenecks:

| Feature / Metric | **osbxl** | **Apache POI (Java)** | **openpyxl (Python)** | **PhpSpreadsheet (PHP)** | **SheetJS / ExcelJS (JS)** |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Core Engine** | **Native C++** | Java | Python | PHP | JavaScript |
| **Large-Scale Data Speed** | **◎ Blazing Fast** | ◯ (Streaming SXSSF) | △ (Slow on large sheets) | ✕ (Severely slow) | ◯ / △ |
| **Memory Footprint** | **◎ Ultra-low (Up to 90% less)** | △ (High heap usage) | △ (Object overhead) | ✕ (Frequent 500 crashes) | △ (V8 heap limit) |
| **Formula Calculation** | **~500 functions (XLOOKUP etc.)** | 300–400 functions | ✕ (Cannot evaluate) | △ (Partial & heavy) | Limited / Paid add-on |
| **AI Code Generation** | **◎ Official Skill Included** | None (High hallucination) | None | None | None |
| **WebAssembly** | **◎ Native Support (Wasm 3.0)** | ✕ | △ (Huge Pyodide runtime) | ✕ | ◯ |

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
├── README.ja.md            # 日本語版 README
├── LICENSE                 # Licenses (MIT for samples, OSBXL agreement for engine)
├── LICENSE.ja.md           # 日本語版ライセンス条項
├── examples/               # Standalone runnable code examples
│   ├── python/             # Python samples
│   │   ├── quickstart.py             # Basic workbook creation & formulas
│   │   ├── chart_sample.py           # Column & line charts creation
│   │   ├── table_comments_sample.py  # Tables (ListObjects) & threaded comments
│   │   └── conditional_format_sample.py # Conditional formatting & data bars
│   ├── nodejs/             # Node.js samples
│   │   ├── quickstart.js             # Basic workbook creation & formulas
│   │   ├── chart_sample.js           # Column & line charts creation
│   │   └── table_comments_sample.js  # Tables (ListObjects) & threaded comments
│   ├── java/               # Java sample (QuickStart.java)
│   ├── php/                # PHP sample (quickstart.php)
│   └── wasm/               # WebAssembly sample (quickstart.wat, quickstart.wasm.js)
├── ja/                     # 日本語版 Web 形式リファレンス (general-01 ~ 11)
└── en/                     # English Web Reference (general-01 ~ 11)
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
- **Library Comparison & Benchmarks (vs POI, openpyxl, PhpSpreadsheet)**: [https://osb-office.com/en/pages/comparison.php](https://osb-office.com/en/pages/comparison.php)
- **AI Skill Guide & Download**: [https://osb-office.com/pages/ai-skill.php](https://osb-office.com/pages/ai-skill.php)
- **Interactive Web Samples**: [https://osb-office.com/sample/](https://osb-office.com/sample/)
- **Product Features**: [https://osb-office.com/pages/features.php](https://osb-office.com/pages/features.php)
- **Free Demo Download**: [https://osb-office.com/pages/download.php](https://osb-office.com/pages/download.php)

---

## 📄 License

- **Sample Code & AI Skills**: [MIT License](LICENSE) — Feel free to use, copy, and modify in your projects.
- **OSBXL Engine (Libraries)**: Commercial / Proprietary by osb-office.inc. See [LICENSE](LICENSE) or [LICENSE.ja.md](LICENSE.ja.md) for software license agreement details.

