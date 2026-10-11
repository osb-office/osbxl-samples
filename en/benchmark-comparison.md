# Excel Library Feature Comparison & Benchmark Performance Report

[日本語版 (benchmark-comparison.md)](../ja/benchmark-comparison.md) | [Sample Collection Home (index.md)](index.md)

This technical report details the **comprehensive feature matrix** and **real-world benchmarks** (from 200 rows up to 50,000 rows) comparing **osbxl** (Java, Python, Node.js, PHP, WebAssembly) against popular open-source spreadsheet libraries (Apache POI, openpyxl, PhpSpreadsheet, ExcelJS / SheetJS).

---

## 1. Feature Support Matrix

A side-by-side comparison baseline covering core capabilities required in production systems:

| Category | Feature Item | osbxl (5 Languages) | Apache POI (Java) | openpyxl (Python) | PhpSpreadsheet (PHP) | ExcelJS / SheetJS (JS) |
|---|---|:---:|:---:|:---:|:---:|:---:|
| **Basic I/O** | .xlsx Read & Write | **◎** | **◎** | **◎** | **◎** | **◎** |
| | Memory Stability on Big Data | **◎** (Flat footprint) | **△** (XSSF OOM; SXSSF constrained) | **△** (High lag on scale) | **✕** (Fatal OOM at 128MB) | **◯ / △** (V8 Heap limit) |
| **Formulas** | Formula Writing (`setFormula`) | **◎** | **◎** | **◎** | **◎** | **◎** |
| | **Headless Formula Evaluation**<br>(Evaluating values without Excel) | **◎ ~500 Functions**<br>(XLOOKUP, VLOOKUP, etc.) | **◯ ~300–400 Functions**<br>(Evaluator) | **✕ Not Supported**<br>(Always None until Excel opens) | **△ Partial only**<br>(Slow & limited functions) | **✕ Not Supported**<br>(Requires paid SheetJS Pro) |
| **Tables** | Table Creation (`ListObject`) | **◎** | **◯** (CTTable low-level XML) | **◯** (Table class) | **◯** (Table class) | **◯** (addTable) |
| | Totals Row & Custom Formulas | **◎** (`setTotalRowFunction`) | **◯** | **△** (Restricted) | **△** (Complex setup) | **△** (Basic sum only) |
| | **Structured References**<br>(e.g. `SUM(Table[Column])`) | **◎ Full Native Eval** | **△ Save only**<br>(Cannot calculate) | **✕ Cannot Eval** | **✕ Cannot Eval** | **✕ Cannot Eval** |
| **Filtering** | **AutoFilter & Criteria Evaluation** | **◎ Fast C++ Eval** | **◯** | **△** (Manual script needed) | **△** | **✕** (Range set only) |
| | Multi-column Sort | **◎** | **△** | **△** | **△** | **✕** |
| **Advanced** | **PivotTable Creation** | **◎** (`PivotTable`) | **◯** (Low-level XML) | **△** (Read-focused) | **✕ Not Supported** | **✕ Not Supported** |
| | **Chart & Graph Generation** | **◎ 2D/3D Charts** | **◯** (Via XDDF) | **◯** (2D/3D support) | **△** (Render bugs) | **✕–△** (Minimal) |
| | Conditional Formatting | **◎ Color scales, DataBars** | **◯** | **◯** | **◯** | **◯** |
| **Security** | **Password-Encrypted .xlsx Export** | **◎** (`save(..., pwd)`) | **◯** (`Biff8EncryptionKey`) | **✕ Not Supported**<br>(Requires external tool) | **✕ Not Supported**<br>(Requires extensions) | **✕ Not Supported** |
| | Decrypting & Reading Protected .xlsx | **◎** (`open(..., pwd)`) | **◯** | **✕ Not Supported** | **✕ Not Supported** | **✕ Not Supported** |
| **Others** | Comments & Threaded Notes | **◎ Both Notes & Threads** | **◯** (Notes only) | **◯** (Comments only) | **◯** (Comments only) | **◯** (Comments only) |
| | **Identical Cross-Language API** | **◎ 100% Identical (5 Langs)** | **✕** (Java only) | **✕** (Python only) | **✕** (PHP only) | **✕** (JS only) |
| | **Official AI Skill (No Hallucination)** | **◎ Zero Hallucination** | **✕** (Frequent fake APIs) | **✕** (Version confusion) | **✕** (High error rate) | **✕** (High error rate) |

---

## 2. Standard Business Workload Benchmark (200 rows / 800 cells)

Measured under identical development machine conditions (Linux x86_64) generating a standard business sales report (**200 rows × 4 columns = 800 cells + SUM formula**).

### Real-world Benchmark Measurements

| Runtime / Language | Library | Execution Time | Memory Usage | Highlights & Analysis |
|---|---|:---:|:---:|---|
| **Java** (JDK 1.8) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | 0.622 s<br>**0.139 s** | 15.19 MB<br>**16.87 MB** | **osbxl is ~4.5x faster**.<br>POI carries heavy XML schema init overhead; osbxl C++ core starts and finishes instantly. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.059 s<br>**0.035 s** | 109.07 MB<br>**125.61 MB** | **nodeosbxl outperforms ExcelJS in the 0.03s range (~1.7x faster)**.<br>Under identical serialization mode (setOnlyMode), nodeosbxl beats ExcelJS with only 18 MB V8 heap consumption. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | **0.018 s**<br>0.032 s | 53.73 MB<br>70.84 MB | Under matching write-only mode, **pyosbxl finishes in 0.03s**.<br>Python heap usage is kept down to just 12 MB. |
| **PHP** (8.5) | PhpSpreadsheet 5.10.0<br>**phposbxl** | 0.117 s<br>**0.122 s** | 26.00 MB<br>**2.00 MB** | Comparable speed, but **phposbxl uses 13x less memory (only 2 MB)**.<br>Decisive advantage for concurrent web downloads. |

---

## 3. Large-Scale Stress Test (50,000 rows / 200,000 cells)

Benchmark measurements scaling data up to **50,000 rows × 4 columns = 200,000 cells + SUM formula**. While popular open-source libraries crash from out-of-memory errors, osbxl completes cleanly across all platforms.

### 50,000-Row Measurement Results

| Runtime / Limits | Library | Execution Time | Memory Usage | Highlights & Fault Analysis |
|---|---|:---:|:---:|---|
| **PHP** (8.5 / 128M) | PhpSpreadsheet 5.10.0<br>**phposbxl** | ✕ Crashed<br>**7.194 s** | ✕ OOM (>128MB limit)<br>**32.00 MB** | **PhpSpreadsheet crashes fatally** (`Allowed memory size exhausted`).<br>**phposbxl runs safely using only 32 MB** (generating a 1.93 MB .xlsx). |
| **Java** (JDK 1.8 / 256M) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | ✕ Crashed<br>**15.740 s** | ✕ OOM (>256MB heap)<br>**42.13 MB** | **POI crashes fatally** (`OutOfMemoryError: Java heap space`).<br>**osbxl finishes cleanly within 42 MB** without requiring rewriting to restricted SXSSF streaming mode. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 1.165 s<br>**2.329 s** | 766.54 MB<br>**1,799 MB**<br>*(V8 heap: 20.9 MB)* | Under identical write-only conditions (setOnlyMode), nodeosbxl clocks in at 2.3s.<br>While ExcelJS inflates V8 heap (>420 MB), **nodeosbxl uses only 20.9 MB of V8 heap**. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 2.470 s<br>**2.433 s** | 378.28 MB<br>**1,439 MB**<br>*(Python heap: 28 MB)* | Under identical write-only conditions (setOnlyMode), **pyosbxl matches openpyxl at 2.4s**.<br>Python runtime heap usage is only 28 MB. |

> 💡 **Memory Structure & Language Heap Safety**:  
> - **Open-Source Libraries (ExcelJS / openpyxl)**:  
>   Keep every cell, row, and style as language-level objects in the runtime heap, rapidly consuming Node.js and Python heap limits (ExcelJS >420 MB, openpyxl >370 MB).
> - **osbxl (nodeosbxl / pyosbxl)**:  
>   Transfers 2D matrix data immediately into the C++ core engine. **Node.js V8 heap usage remains at a modest 20.9 MB for 50,000 rows, and Python heap is only 28 MB**, entirely preventing language runtime OOM heap crashes. The process Peak RSS represents temporary native working buffers used by the C++ core for multi-threaded XML generation and ZIP compression.

### Fatal Bottlenecks in Open-Source Libraries

- **PhpSpreadsheet (PHP 8.5)**:
  `Fatal error: Allowed memory size of 134217728 bytes exhausted`
  &rarr; Blew past the standard 128 MB PHP memory ceiling and crashed immediately during StringTable generation.
- **Apache POI (Java JDK 1.8)**:
  `java.lang.OutOfMemoryError: Java heap space`
  &rarr; Crashed under standard 256 MB heap. Forced to rewrite using `SXSSF` streaming mode (which forbids formula re-evaluation, reading back, and random access).

---

## 4. Summary of Critical Limitations in Existing OSS Libraries

### 1. openpyxl (Python)
- **Critical Flaw**: **Zero formula evaluation in headless environments**.
  - Cell values remain `None` unless opened in the Excel desktop application.
  - Cannot export password-encrypted workbooks without third-party tools.

### 2. PhpSpreadsheet (PHP)
- **Critical Flaw**: **Extravagant memory footprint and missing PivotTables**.
  - Consumes 26 MB for just 200 rows, guaranteeing `Allowed memory size exhausted` crashes at scale.
  - No PivotTable generation; no built-in password encryption/decryption.

### 3. Apache POI (Java)
- **Critical Flaw**: **OOM heap blowups and SXSSF API crippling**.
  - Default `XSSFWorkbook` crashes at 50,000 rows on 256 MB heaps.
  - SXSSF streaming loses random access, structured reference evaluation, and formula readback.

### 4. ExcelJS / SheetJS (Node.js)
- **Critical Flaw**: **Lack of formula calculation engines and PivotTables**.
  - Free versions are merely serializing writers without calculation capabilities.

---

## 5. Conclusion: Why osbxl?

1. **Daily Workloads (200 rows)**:
   - ~4.5x faster in Java, 13x less memory in PHP, 0.05–0.1s instant responses across all 5 languages.
2. **Absolute Functional Superiority**:
   - **Headless evaluation of ~500 Excel formulas**
   - **Structured references, PivotTables, and 2D/3D charts**
   - **Built-in AES password encryption and decryption**
   - No other spreadsheet library in the modern software landscape delivers these capabilities through a single unified API across 5 languages.

---

## 🔗 Official Links

- [Official Website](https://osb-office.com/)
- [Library Comparison & Benchmarks](https://osb-office.com/en/pages/comparison.php)
- [Official AI Skill Guide & Download](https://osb-office.com/pages/ai-skill.php)
- [Interactive Web Samples](https://osb-office.com/sample/)

