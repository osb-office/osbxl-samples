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

## 📊 Feature Comparison Matrix & Real-world Benchmarks

A side-by-side comparison of advanced features (formula engine, PivotTables, charts, password encryption) and measured performance between popular open-source libraries and **osbxl**.

### 1. Feature Support Matrix

| Feature | osbxl (5 Languages) | Apache POI (Java) | openpyxl (Python) | PhpSpreadsheet (PHP) | ExcelJS / SheetJS (JS) |
|---|:---:|:---:|:---:|:---:|:---:|
| **.xlsx Read & Write** | **◎ Fast & Stable** | **◎** | **◎** | **◎** | **◎** |
| **Headless Formula Calculation**<br>(Evaluating values without Excel) | **◎ ~500 Functions**<br>(XLOOKUP, VLOOKUP, etc.) | **◯ 300–400 Functions**<br>(Evaluator) | **✕ Not Supported**<br>(None until Excel opens) | **△ Partial**<br>(Heavy & limited) | **✕ Not Supported**<br>(Requires paid SheetJS Pro) |
| **Tables & Structured References**<br>(e.g. `SUM(Table[Column])`) | **◎ Full Native Eval** | **△ No Calculation**<br>(Save only) | **✕ Cannot Eval** | **✕ Cannot Eval** | **✕ Cannot Eval** |
| **AutoFilter** | **◎ Advanced Filter** | **◯** | **◯** | **◯** | **◯** |
| **PivotTable Creation** | **◎ Standard API** | **◯ Low-level XML** | **△ Read-focused** | **✕ Not Supported** | **✕ Not Supported** |
| **Chart & Graph Generation** | **◎ 2D & 3D Charts** | **◯ Via XDDF** | **◯ 2D/3D Support** | **△ Limited rendering** | **✕–△ Minimal** |
| **Password Encryption / Decryption** | **◎ Built-in** | **◯ Standard** | **✕ Needs 3rd-party tool** | **✕ Needs extension** | **✕ Not Supported** |
| **Cell Comments (Notes & Threads)** | **◎ Both Notes & Threads** | **◯ Notes only** | **◯ Comments only** | **◯ Comments only** | **◯ Comments only** |
| **AI Codegen Optimization** | **◎ Official AI Skill**<br>(Zero hallucination) | **✕ Unofficial**<br>(Frequent fake APIs) | **✕ Unofficial** | **✕ Unofficial** | **✕ Unofficial** |

### 2. Standard 200-Row (800-Cell) Benchmark

Measured under the same development environment (Linux x86_64) generating a standard business report (**200 rows × 4 columns = 800 cells + SUM formula**).

| Language | Library | Time (sec) | Memory (MB) | Highlights |
|---|---|:---:|:---:|---|
| **Java** (JDK 1.8) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | 0.622 s<br>**0.139 s** | 15.19 MB<br>**16.87 MB** | **osbxl is ~4.5x faster**.<br>POI carries heavy XML schema init overhead; osbxl starts and finishes instantly. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.059 s<br>**0.035 s** | 109.07 MB<br>**125.61 MB** | **nodeosbxl outperforms ExcelJS in the 0.03s range (~1.7x faster)**.<br>Under identical serialization mode (setOnlyMode), nodeosbxl beats ExcelJS with only 18 MB V8 heap consumption. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | **0.018 s**<br>0.032 s | 53.73 MB<br>70.84 MB | Under matching write-only mode, **pyosbxl finishes in 0.03s**.<br>Python heap usage is kept down to just 12 MB. |
| **PHP** (8.5) | PhpSpreadsheet 5.10.0<br>**phposbxl** | 0.117 s<br>**0.122 s** | 26.00 MB<br>**2.00 MB** | Similar speed, but **phposbxl uses 13x less memory (only 2 MB)**.<br>Critical advantage under concurrent web downloads. |

### 3. Large-Scale 50,000-Row (200,000 Cells) Stress Test

Stress test results scaling data up to **50,000 rows × 4 columns = 200,000 cells + SUM formula**. While popular open-source libraries crash from out-of-memory errors, osbxl completes cleanly across all platforms.

| Runtime / Limits | Library | Execution Time | Memory Usage | Highlights & Fault Analysis |
|---|---|:---:|:---:|---|
| **PHP** (8.5 / 128M) | PhpSpreadsheet 5.10.0<br>**phposbxl** | ✕ Crashed<br>**7.194 s** | ✕ OOM (>128MB limit)<br>**32.00 MB** | **PhpSpreadsheet crashes fatally** (`Allowed memory size exhausted`).<br>**phposbxl runs safely using only 32 MB** (generating a 1.93 MB .xlsx). |
| **Java** (JDK 1.8 / 256M) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | ✕ Crashed<br>**15.740 s** | ✕ OOM (>256MB heap)<br>**42.13 MB** | **POI crashes fatally** (`OutOfMemoryError: Java heap space`).<br>**osbxl finishes cleanly within 42 MB** without requiring rewriting to restricted SXSSF streaming mode. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 1.165 s<br>**2.329 s** | 766.54 MB<br>**1,799 MB**<br>*(V8 heap: 20.9 MB)* | Under identical write-only conditions (setOnlyMode), nodeosbxl clocks in at 2.3s.<br>While ExcelJS inflates V8 heap (>420 MB), **nodeosbxl uses only 20.9 MB of V8 heap**. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 2.470 s<br>**2.433 s** | 378.28 MB<br>**1,439 MB**<br>*(Python heap: 28 MB)* | Under identical write-only conditions (setOnlyMode), **pyosbxl matches openpyxl at 2.4s**.<br>Python runtime heap usage is only 28 MB. |

> 💡 **Memory Structure & Language Heap Safety**:  
> ExcelJS and openpyxl keep every cell, row, and style as language-level objects in the runtime heap, rapidly consuming Node.js and Python heap limits (ExcelJS >420 MB, openpyxl >370 MB).<br>
> In contrast, osbxl transfers 2D matrix data immediately into the C++ core engine. **Node.js V8 heap usage remains at a modest 20.9 MB for 50,000 rows, and Python heap is only 28 MB**, entirely preventing language runtime OOM heap crashes. The process Peak RSS represents temporary native working buffers used by the C++ core for multi-threaded XML generation and ZIP compression.  
> 📖 **Full Feature Matrix & Detailed Analysis**: See [Excel Library Feature Comparison & Benchmark Report (en/benchmark-comparison.md)](en/benchmark-comparison.md).

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
- **Excel Library Performance & Features**: [https://osb-office.com/en/pages/comparison.php](https://osb-office.com/en/pages/comparison.php)
- **osbxl vs Famous Libraries Detailed Comparison (200 & 50,000 rows)**: [https://osb-office.com/en/pages/members/benchmark_report.php](https://osb-office.com/en/pages/members/benchmark_report.php)
- **Load Test Benchmark Report**: [https://osb-office.com/en/pages/members/tips.php](https://osb-office.com/en/pages/members/tips.php)
- **AI Skill Guide & Download**: [https://osb-office.com/pages/ai-skill.php](https://osb-office.com/pages/ai-skill.php)
- **Interactive Web Samples**: [https://osb-office.com/sample/](https://osb-office.com/sample/)
- **Product Features**: [https://osb-office.com/pages/features.php](https://osb-office.com/pages/features.php)
- **Free Demo Download**: [https://osb-office.com/pages/download.php](https://osb-office.com/pages/download.php)

---

## 📄 License

- **Sample Code & AI Skills**: [MIT License](LICENSE) — Feel free to use, copy, and modify in your projects.
- **OSBXL Engine (Libraries)**: Commercial / Proprietary by osb-office.inc. See [LICENSE](LICENSE) or [LICENSE.ja.md](LICENSE.ja.md) for software license agreement details.

