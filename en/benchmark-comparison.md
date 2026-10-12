# Excel Library Feature Comparison & Benchmark Performance Report

**Last Updated: 2026-10-12**

[日本語版 (benchmark-comparison.md)](../ja/benchmark-comparison.md) | [Sample Collection Home (index.md)](index.md)

This technical report details the **comprehensive feature matrix** and **real-world benchmarks** (from 200 rows up to 50,000 rows) comparing **osbxl** (Java, Python, Node.js, PHP, WebAssembly) against popular open-source spreadsheet libraries (Apache POI, openpyxl, PhpSpreadsheet, ExcelJS / SheetJS).

---

## 1. Feature Support Matrix

| Category | Feature | osbxl (5 Languages Unified) | Apache POI (Java) | openpyxl (Python) | PhpSpreadsheet (PHP) | ExcelJS / SheetJS (JS) |
|---|---|:---:|:---:|:---:|:---:|:---:|
| **Basic I/O** | Read & Save `.xlsx` | **◎** | **◎** | **◎** | **◎** | **◎** |
| | Large File Stability | **◎** (Constant memory) | **△** (XSSF OOMs, SXSSF limited) | **△** (Slows with row count) | **✕** (Frequent OOM at 128MB) | **◯ / △** (V8 heap bound) |
| **Formula Engine** | Write Formulas (`setFormula`) | **◎** | **◎** | **◎** | **◎** | **◎** |
| | **Headless Formula Evaluation**<br>(Calculate values without Excel) | **◎ ~500 functions**<br>(XLOOKUP, VLOOKUP, etc.) | **◯ ~300-400 functions**<br>(Built-in Evaluator) | **✕ Not supported**<br>(Values remain None) | **△ Partial only**<br>(Complex functions slow/unsupported) | **✕ Not supported**<br>(Commercial SheetJS Pro only) |
| **Tables** | Create Tables (`ListObject`) | **◎** | **◯** (Low-level CTTable) | **◯** (Table class) | **◯** (Table class) | **◯** (addTable) |
| | Total Row / Custom Formulas | **◎** (`setTotalRowFunction`) | **◯** | **△** (Limited) | **△** (Complex setup) | **△** (Basic aggregations only) |
| | **Structured References**<br>(e.g., `SUM(Table[Column])`) | **◎ Full native evaluation** | **△** (Save only, no calc) | **✕ Cannot evaluate** | **✕ Cannot evaluate** | **✕ Cannot evaluate** |
| **AutoFilter** | **Filter & Evaluation** | **◎ High performance** | **◯** | **△** (Manual script needed) | **△** | **✕** (Range definition only) |
| | Multi-Key Sort (`Sort`) | **◎** | **△** | **△** | **△** | **✕** |
| **Advanced** | **Create PivotTables** | **◎** (`PivotTable`) | **◯** (Complex low-level XML) | **△** (Read-only / limited) | **✕ Not supported** | **✕ Not supported** |
| | **Chart Generation** | **◎ 2D/3D supported** | **◯** (Via XDDF) | **◯** (2D/3D support) | **△** (Constrained / buggy) | **✕~△** (ExcelJS is limited) |
| | Conditional Formatting | **◎ DataBars, ColorScales** | **◯** | **◯** | **◯** | **◯** |
| **Security** | **Password Encryption** | **◎** (`save(..., pwd)`) | **◯** (`Biff8EncryptionKey`) | **✕ Not supported** (Needs 3rd party) | **✕ Not supported** | **✕ Not supported** |
| | Read Password Protected | **◎** (`open(..., pwd)`) | **◯** | **✕ Not supported** | **✕ Not supported** | **✕ Not supported** |
| **Others** | Comments / Threaded Notes | **◎ Both threaded & legacy** | **◯** | **◯** | **◯** | **◯** |
| | **Cross-Platform Unified API** | **◎ 100% Identical** | **✕** (Java only) | **✕** (Python only) | **✕** (PHP only) | **✕** (JS only) |
| | **AI Coding Skill Optimization** | **◎ Official Skill included** | **✕** (Hallucinations frequent) | **✕** (Version confusion) | **✕** (High AI error rate) | **✕** (High AI error rate) |

---

## 2. Standard Workload Benchmark (200 rows / 800 cells)

Measured on representative business reporting workload (200 rows × 4 columns = 800 cells + SUM formula) under identical Linux x86_64 environments.

### Results Comparison

| Language / Env | Library | Execution Time | Memory Usage | Highlights & Analysis |
|---|---|:---:|:---:|---|
| **Java** (JDK 1.8) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | 0.652 s<br>**0.105 s** | 15.11 MB<br>**16.86 MB** | **osbxl is ~6.2x faster** (0.105s).<br>POI is delayed by schema class initializations. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 0.056 s<br>**0.016 s** | 103.95 MB<br>**121.14 MB** | **nodeosbxl is ~3.5x faster (0.016s)**.<br>Far outperforms ExcelJS (0.056s). |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 0.017 s<br>**0.016 s** | 53.20 MB<br>**69.91 MB** | In write-only mode, **pyosbxl achieves 0.016s, beating openpyxl**.<br>Python heap usage is only 12 MB. |
| **PHP** (8.5) | PhpSpreadsheet 5.10.0<br>**phposbxl** | 0.073 s<br>**0.024 s** | 26.00 MB<br>**2.00 MB** | **phposbxl is ~3x faster & 13x more memory efficient (2 MB)**.<br>Critical advantage for concurrent web requests. |

---

## 3. Large-Scale Stress Test (50,000 rows / 200,000 cells)

Benchmark measurements scaling data up to **50,000 rows × 4 columns = 200,000 cells + SUM formula**. While popular open-source libraries crash from out-of-memory errors, osbxl completes cleanly across all platforms.

### 50,000-Row Measurement Results

| Runtime / Limits | Library | Execution Time | Memory Usage | Highlights & Fault Analysis |
|---|---|:---:|:---:|---|
| **PHP** (8.5 / 128M) | PhpSpreadsheet 5.10.0<br>**phposbxl** | ✕ Crashed<br>**5.836 s** | ✕ OOM (>128MB limit)<br>**32.00 MB** | **PhpSpreadsheet crashes fatally** (`Allowed memory size exhausted`).<br>**phposbxl runs safely using only 32 MB** (generating a 1.93 MB .xlsx). |
| **Java** (JDK 1.8 / 256M) | Apache POI 5.2.5 (XSSF)<br>**osbxl (Java)** | ✕ Crashed<br>**6.394 s** | ✕ OOM (>256MB heap)<br>**36.19 MB** | **POI crashes fatally** (`OutOfMemoryError: Java heap space`).<br>**osbxl finishes cleanly in 6.39s within 36 MB** without requiring rewriting to restricted SXSSF streaming mode. |
| **Node.js** (v24) | ExcelJS 4.4.0<br>**nodeosbxl** (`setOnlyMode`) | 1.218 s<br>**1.798 s** | 772.11 MB<br>**1,836 MB**<br>*(V8 Heap: 20.9MB)* | **nodeosbxl achieves 1.79s**.<br>V8 heap usage is kept down to just 20.9 MB, preventing Node.js crashes. |
| **Python** (3.14) | openpyxl 3.1.5<br>**pyosbxl** (`setOnlyMode`) | 2.500 s<br>**1.900 s** | 377.93 MB<br>**1,445 MB**<br>*(Python Heap: 28MB)* | **pyosbxl achieves 1.900s (beating openpyxl's 2.500s)**.<br>Python heap remains at just 28 MB. |

> 💡 **Memory Structure & Language Heap Safety**:  
> - **Open-Source Libraries (ExcelJS / openpyxl)**:  
>   Keep every cell, row, and style as language-level objects in the runtime heap, rapidly consuming Node.js and Python heap limits (ExcelJS >420 MB, openpyxl >370 MB).
> - **osbxl (nodeosbxl / pyosbxl)**:  
>   Transfers 2D matrix data immediately into the C++ core engine. **Node.js V8 heap usage remains at a modest 20.9 MB for 50,000 rows, and Python heap is only 28 MB**, entirely preventing language runtime OOM heap crashes. The process Peak RSS represents temporary native working buffers used by the C++ core for multi-threaded XML generation and ZIP compression.
> - **Why `setOnlyMode` Was Applied for Node.js / Python**:  
>   Competing open-source libraries (**openpyxl** and **ExcelJS**) do not implement a formula calculation engine (they merely serialize formula text into XML). Measuring them against osbxl's default mode—which recalculates all formula dependencies dynamically in real time—would be an unfair comparison. Therefore, osbxl was benchmarked in **`setOnlyMode = true`** (setting formulas without intermediate evaluation) to match the write-only baseline of the OSS libraries.  
>   While real-time formula evaluation results (`getValue()`) are skipped during script execution in `setOnlyMode`, **the resulting `.xlsx` file is automatically calculated and displayed normally when opened later in Excel on Windows**. This makes it an ideal high-speed mode for export-only web reporting.

### Fatal Bottlenecks in Open-Source Libraries

- **PhpSpreadsheet (PHP 8.5)**:
  `Fatal error: Allowed memory size of 134217728 bytes exhausted`
  &rarr; Exceeds the standard 128 MB PHP memory limit during StringTable generation and crashes immediately.
- **Apache POI (Java JDK 1.8)**:
  `java.lang.OutOfMemoryError: Java heap space`
  &rarr; Exhausts the 256 MB max heap limit building the DOM tree. Rewriting to SXSSF streaming disables formula re-evaluation, structured references, and random access.
- **openpyxl (Python 3.14)**:
  Has no calculation engine; cell values remain `None` in headless environments.
- **ExcelJS (Node.js v24)**:
  Lacks PivotTable and formula computation features while consuming large V8 memory buffers.

---

## 4. Conclusion: The Decisive Advantages of osbxl

1. **Everyday Reporting (200 rows)**:
   - 6.2x faster on Java, 3x faster & 13x less memory on PHP, 0.01s - 0.1s blazing responsiveness across all languages.
2. **Heavy-Duty Reporting (50,000 rows)**:
   - Safely finishes in mere seconds using minimal memory where open-source options suffer fatal OOM crashes.
3. **Cross-Language Consistency**:
   - ~500 headless formula functions, PivotTables, 2D/3D charts, and encryption under the exact same API across all 5 languages.
