# osbxl-samples

> **[🇯🇵 日本語版の README はこちら (README.ja.md)](README.ja.md)**

Official verified code samples and patterns for **[osbxl (https://osb-office.com)](https://osb-office.com/)** — an AI-ready, ultra-fast C++ core multi-platform spreadsheet calculation and Excel generation engine.

[![Official Website](https://img.shields.io/badge/Official_Website-osb--office.com-blue?style=flat-square)](https://osb-office.com/)
[![Library Comparison](https://img.shields.io/badge/Benchmark-vs%20POI%20%7C%20openpyxl%20%7C%20PhpSpreadsheet-red?style=flat-square)](https://osb-office.com/en/pages/comparison.php)
[![AI-Ready](https://img.shields.io/badge/AI--Ready-Cursor%20%7C%20Copilot%20%7C%20Gemini-orange?style=flat-square)](https://osb-office.com/pages/ai-skill.php)
[![Languages](https://img.shields.io/badge/Languages-Python%20%7C%20Node.js%20%7C%20Java%20%7C%20PHP%20%7C%20Wasm-green?style=flat-square)](https://osb-office.com/pages/features.php)

---

## 📊 Comparison with Popular Excel Libraries

> 📈 **[New! Multi-Platform Benchmark Report (v1.5.1)](benchmarks/BENCHMARK_REPORT_3OS.en.md)**: Real-world stress testing across GCP C3/C4, macOS M4, and Windows. See how osbxl completes 50,000 rows in just 0.1s–0.5s where Apache POI and PhpSpreadsheet crash with OOM.
> 
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

**Last Updated: 2026-10-18**

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

Measured under representative production and development environments generating a standard business report (**200 rows × 4 columns = 800 cells + SUM formula**).

| Language | Target Library | GCP C4 (Linux) | macOS (Apple M4) | Memory (MB) | Highlights |
| :--- | :--- | :---: | :---: | :---: | :--- |
| **PHP** (8.5) | PhpSpreadsheet 5.10.0<br>**phposbxl** | 0.017 s<br>**0.007 s** | 0.018 s<br>**0.005 s** | 26.0 MB<br>**2.0 MB** | **~2.4x–3.5x faster & 13x less memory**.<br>Web API response finishes in milliseconds. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 0.011 s<br>**0.007 s** | 0.005 s<br>**0.003 s** | 53.2 MB<br>**12.0 MB (Py)** | **~1.7x faster (down to 3–7 ms)**.<br>Python heap usage is kept down to just 12 MB. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.011 s<br>**0.008 s** | 0.005 s<br>**0.004 s** | 104 MB<br>**18.0 MB (V8)** | **Finishes in blazing fast 4–8 ms**.<br>Never blocks the event loop with minimal V8 heap. |
| **Java** (JDK 24/8) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | 0.542 s<br>**0.110 s** | 0.311 s<br>**0.057 s** | 15.1 MB<br>**16.9 MB** | **~4.9x–5.5x faster**.<br>Eliminates heavy XML schema init overhead. |

### 3. Mid-Scale 10,000-Row (40,000-Cell) Benchmark

Benchmark results for standard business reporting (such as monthly logs, ledger exports) generating **10,000 rows × 4 columns = 40,000 cells + SUM formula**. osbxl completes in mere tens of milliseconds.

| Language | Target Library | GCP C4 (Linux) | macOS (Apple M4) | Speedup Factor | Highlights |
| :--- | :--- | :---: | :---: | :---: | :--- |
| **PHP** (8.5) | PhpSpreadsheet 5.10.0<br>**phposbxl** | 0.760 s<br>**0.055 s** | 0.862 s<br>**0.037 s** | **13.9x–23.4x faster** | While PhpSpreadsheet takes nearly a second, phposbxl returns in just 37–55 ms. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 0.324 s<br>**0.058 s** | 0.166 s<br>**0.030 s** | **~5.5x faster** | 10k rows completed in 30–58 ms, drastically cutting batch wait times. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.239 s<br>**0.056 s** | 0.109 s<br>**0.028 s** | **3.9x–4.3x faster** | Down to 28 ms; preserves event loop responsiveness even under high traffic. |
| **Java** (JDK 24/8) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | 1.335 s<br>**0.208 s** | 0.620 s<br>**0.098 s** | **6.3x–6.4x faster** | While POI takes 0.6s–1.3s, osbxl completes smoothly in ~0.1s–0.2s. |

### 4. Large-Scale 50,000-Row (200,000 Cells) Stress Test

Stress test results scaling data up to **50,000 rows × 4 columns = 200,000 cells + SUM formula**. While popular open-source libraries crash from out-of-memory errors, osbxl completes cleanly across all platforms in 0.1s–0.5s.

| Runtime / Limits | Target Library | GCP C4 (Linux) | macOS (Apple M4) | Memory Usage | Highlights & Fault Analysis |
| :--- | :--- | :---: | :---: | :---: | :--- |
| **PHP** (8.5 / 128M) | PhpSpreadsheet 5.10.0<br>**phposbxl** | ✕ OOM Crash<br>**0.346 s** | ✕ OOM Crash<br>**0.189 s** | ✕ >128MB limit<br>**32.0 MB** | **PhpSpreadsheet crashes fatally** (`Allowed memory size exhausted`).<br>**phposbxl runs safely in 0.1s–0.3s using only 32 MB**. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 1.662 s<br>**0.378 s** | 0.850 s<br>**0.183 s** | 378 MB<br>**28.0 MB (Py)** | **4.39x faster on C4, 4.64x faster on M4**.<br>Python heap is only 28 MB with zero risk of OOM. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.979 s<br>**0.329 s** | 0.502 s<br>**0.133 s** | 772 MB<br>**20.9 MB (V8)** | **2.97x faster on C4, 3.77x faster on M4**.<br>V8 heap capped at 20.9 MB, ideal for concurrent services. |
| **Java** (JDK 24/8 / 256M) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | ✕ OOM Crash<br>**0.553 s** | ✕ OOM Crash<br>**0.247 s** | ✕ >256MB heap<br>**36.2 MB** | **POI crashes fatally** (`OutOfMemoryError: Java heap space`).<br>**osbxl finishes cleanly in 0.2s–0.5s within 36.2 MB heap**. |

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

