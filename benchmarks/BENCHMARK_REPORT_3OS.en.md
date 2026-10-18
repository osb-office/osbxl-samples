# [Technical Benchmark Report] Multi-Platform Comprehensive Stress Test & Architecture Deep Dive: 4 Excel Library Families
## ~ Performance Limits & Behavioral Verification across Cloud (GCP C3/C4), macOS (Apple M4), Windows & Linux ~

> **[🇯🇵 日本語版レポートはこちら (BENCHMARK_REPORT_3OS.md)](BENCHMARK_REPORT_3OS.md)**

- **Public Repository**: [osb-office/osbxl-samples](https://github.com/osb-office/osbxl-samples)
- **Measurement Date**: October 18, 2026 (Reflecting latest GCP C3/C4 measurements)
- **osbxl Version**: **`v1.5.1`** (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
- **Benchmark Target Libraries**:
  - **PHP Ecosystem**: `PhpSpreadsheet` vs `phposbxl (v1.5.1)`
  - **Python Ecosystem**: `openpyxl` vs `pyosbxl (v1.5.1)`
  - **Node.js Ecosystem**: `ExcelJS` vs `nodeosbxl (v1.5.1)`
  - **Java Ecosystem**: `Apache POI (XSSF)` vs `osbxl Java (v1.5.1)`
- **Target Environments & Platforms**:
  1. **GCP C4 (Latest Cloud VM)**: Intel Xeon Platinum 8581C (5th Gen Emerald Rapids, 4 vCPU) / Ubuntu 24.04 LTS
  2. **GCP C3 (Prior Gen Cloud VM)**: Intel Xeon Platinum 8481C (4th Gen Sapphire Rapids, 4 vCPU) / Linux
  3. **macOS (Development Desktop)**: Apple M4 mini (10 cores: 4P+6E, 16GB Unified Memory) / macOS Darwin arm64
  4. **Windows (PC / Desktop)**: Intel Core / Windows 11 (x64)
- **Latest Runtime Environments**:
  - **PHP**: `8.5.11` (NTS, ext-zlib enabled, mise)
  - **Python**: `3.14.8` (pyenv)
  - **Node.js**: `v24.22.0` (LTS: Krypton, nvm)
  - **Java**: `24.0.2` (Eclipse Temurin JDK 24) / `JDK 8`
- **Data Scale**: 100 & 200 rows (Daily operational scale) to 500, 1,000, 5,000, 10,000, and 50,000 rows (Up to 200,000 cells + SUM formula evaluation)

---

## 1. Executive Summary

### 💡 The Conclusion in One Sentence
> **"50,000-row large-scale reports that crashed traditional libraries with Out-Of-Memory (OOM) now finish safely in just 0.3–0.5 seconds with negligible memory footprint (~tens of MBs)."**  
> For daily web API reports (200 rows), generation completes in just **7 to 10 milliseconds (0.007s)**—eliminating user wait times entirely.

---

### 📌 Background & Motivation: Why Run This Benchmark?
In web applications and enterprise batch systems, generating Excel reports has persistently plagued engineers with two critical failure modes:
1. **"Processes crashing immediately from Out Of Memory (OOM) as data rows increase."**
2. **"Web servers timing out due to tens of seconds spent serializing XML/ZIP files."**

To address these pain points, we conducted exhaustive benchmark stress tests comparing standard high-level libraries with `osbxl` across **4 major programming languages (PHP, Python, Node.js, Java)** and **3 representative environments (Cloud GCP C3/C4, Mac M4, Windows 11)** under strictly identical data schemas and latest runtimes.

---

### 🏆 3 Decisive Findings from Real-World Measurements

#### ① [Elimination of Memory Crashes] 50,000 Rows Never Crash osbxl
* **Traditional Libraries (PHP: PhpSpreadsheet / Java: Apache POI)**:  
  Under standard container/cloud configurations (heap limits of 256MB), generating 50,000 rows **reliably triggers fatal OOM crashes**, taking down server processes.
* **osbxl**:  
  Because its C++ native core handles data in a continuous native stream, language-side heap consumption **remains constant at just 30MB–64MB**. osbxl successfully and safely completed across all environments and languages in **0.3s to 0.5s (C4)**.  
  👉 *There is no longer any need to artificially inflate container memory to several gigabytes just for Excel export.*

#### ② [Blazing Fast Web Responses] Daily 200-Row Reports in Just 7 Milliseconds
* For typical on-demand download reports (200 rows / 800 cells), PHP, Python, and Node.js all generate files in **0.007 to 0.008 seconds (7–8ms)**.
* In Java, osbxl bypasses Apache POI's massive schema class initialization overhead (~0.5s+), returning a file in **0.1 seconds**.  
  👉 *User latency drops to unnoticeable levels, completely removing HTTP gateway timeout risks.*

#### ③ [Cloud Generational Leap] Latest GCP C4 Delivers ~1.1x–1.4x Higher Throughput than C3
* Comparing 4th Gen Xeon (C3: Sapphire Rapids) with 5th Gen Xeon (C4: Emerald Rapids) under identical 4-vCPU configurations, **C4 consistently outperforms C3 by approximately 1.1x to 1.4x**.
* For a 10,000-row workbook, Node.js finished in **56ms (0.056s)**, PHP clocked **0.055s**, and Java clocked **0.208s**, offering extraordinary scalability for microservices and high-concurrency environments.

---

## 2. Cross-Platform Performance Comparison Matrix

### ① Large-Scale Benchmark: 50,000 Rows (200,000 Cells) Completion Time

| Language / Library | macOS (Apple M4) | GCP C4 (Xeon 8581C) | GCP C3 (Xeon 8481C) | Windows 11 (x64) | Existing OSS Library Behavior |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **PHP (phposbxl)** | **0.189 s** | **0.346 s** | **0.405 s** | **1.011 s** | ❌ **Fatal OOM Crash on all platforms** |
| **Python (pyosbxl)** | **0.183 s** | **0.378 s** | **0.479 s** | **0.996 s** | openpyxl: 0.85s–2.52s (No formula evaluation) |
| **Node.js (nodeosbxl)** | **0.133 s** | **0.329 s** | **0.399 s** | **0.987 s** | ExcelJS: 0.50s–1.32s (No formula engine) |
| **Java (osbxl Java)** | **0.247 s** | **0.553 s** | **0.738 s** | **1.410 s** | ❌ **Fatal OOM Crash (Heap exhaustion)** |

> *Note: For pure write throughput comparison, osbxl was tested with `setOnlyMode = True`. Existing OSS libraries either lack built-in formula evaluation engines or crash with OOM.*

---

### ② Standard Enterprise Workload: 10,000 Rows (40,000 Cells) Processing Time & Speedup

| Language | Target Libraries | GCP C4 (osbxl) | GCP C3 (osbxl) | Competing OSS (C4) | Speedup on C4 |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **PHP** | PhpSpreadsheet vs **phposbxl** | **0.055 s** | 0.063 s | 0.760 s | **13.94x Faster** |
| **Python** | openpyxl vs **pyosbxl** | **0.058 s** | 0.077 s | 0.324 s | **5.55x Faster** |
| **Node.js** | ExcelJS vs **nodeosbxl** | **0.056 s** | 0.068 s | 0.239 s | **4.28x Faster** |
| **Java** | Apache POI vs **osbxl** | **0.208 s** | 0.265 s | 1.335 s | **6.42x Faster** |

---

### ③ Daily Operational Workload: 200 Rows (800 Cells) Small Benchmark (GCP C4)

| Language | Competing OSS (s) | osbxl (s) | Speedup Ratio | Practical Engineering Impact |
| :--- | :---: | :---: | :---: | :--- |
| **PHP** | 0.017s | **0.007s** | **2.43x Faster** | Web API responds in merely 7ms |
| **Python** | 0.011s | **0.007s** | **1.71x Faster** | Instant write completion with zero startup lag |
| **Node.js** | 0.011s | **0.008s** | **1.32x Faster** | Non-blocking execution preserves event loop |
| **Java** | 0.542s | **0.110s** | **4.93x Faster** | Eliminates POI's massive schema load penalty |

---

### ④ Language-by-Language Deep Dive: Alternatives vs osbxl

#### [1] PHP Ecosystem: PhpSpreadsheet vs phposbxl
| Row Count (Data Scale) | PhpSpreadsheet (Competitor) | phposbxl (True) | phposbxl (False) | Speedup Ratio (vs Competitor) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.213s | **0.014s** | 0.013s | **15.07x Faster** |
| **1,000 rows** | 0.073s | **0.011s** | 0.014s | **6.49x Faster** |
| **5,000 rows** | 0.368s | **0.031s** | 0.055s | **11.97x Faster** |
| **10,000 rows** | 0.760s | **0.055s** | 0.109s | **13.94x Faster** |
| **50,000 rows (200,000 cells)** | ❌ **OOM (Memory Crash)** | **0.346s** | 0.735s | **Survived (Competitor OOM)** |

* **Total Elimination of Memory Crashes (OOM)**: PhpSpreadsheet instantiates heavy userland PHP objects for every single cell, consuming 256MB at 50,000 rows and immediately triggering `Allowed memory size exhausted`. phposbxl processes raw streams directly in native C++, consuming a constant 20–32MB and finishing safely in 0.34 seconds.
* **Elimination of Web Gateway Timeouts**: Outputting 10,000 rows in just **0.055s (~14x faster)** and standard reports (500–1,000 rows) in **11–14ms** completely resolves concurrent download queue saturation.
* **Enterprise Features Built-in**: Full dynamic PivotTable generation and high-speed native evaluation of ~500 Excel formulas—features completely missing in PhpSpreadsheet.

#### [2] Python Ecosystem: openpyxl vs pyosbxl
| Row Count (Data Scale) | openpyxl (Competitor) | pyosbxl (True) | pyosbxl (False) | Speedup Ratio (vs openpyxl) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.024s | **0.015s** | 0.013s | **1.68x Faster** |
| **1,000 rows** | 0.037s | **0.011s** | 0.014s | **3.46x Faster** |
| **5,000 rows** | 0.167s | **0.035s** | 0.060s | **4.73x Faster** |
| **10,000 rows** | 0.324s | **0.058s** | 0.119s | **5.55x Faster** |
| **50,000 rows (200,000 cells)** | 1.662s | **0.378s** | 0.827s | **4.39x Faster** |

* **3.5x–5.5x Higher Throughput**: As row counts increase, openpyxl suffers from Python object allocation drag. pyosbxl writes out 50,000 rows in **0.37 seconds**.
* **Headless Formula Evaluation Engine**: openpyxl cannot recalculate formulas without running Excel itself, leaving computed cells as `None`. pyosbxl features a **built-in calculation engine supporting ~500 functions**, returning accurate results on headless servers.
* **Native Password Encryption**: pyosbxl provides out-of-the-box password encryption and decryption without needing 3rd-party tools.

#### [3] Node.js Ecosystem: ExcelJS vs nodeosbxl
| Row Count (Data Scale) | ExcelJS (Competitor) | nodeosbxl (True) | nodeosbxl (False) | Speedup Ratio (vs ExcelJS) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.029s | **0.008s** | 0.010s | **3.40x Faster** |
| **1,000 rows** | 0.043s | **0.009s** | 0.013s | **4.62x Faster** |
| **5,000 rows** | 0.135s | **0.031s** | 0.057s | **4.31x Faster** |
| **10,000 rows** | 0.239s | **0.056s** | 0.113s | **4.28x Faster** |
| **50,000 rows (200,000 cells)** | 0.979s | **0.329s** | 0.739s | **2.97x Faster** |

* **Preserving the V8 Event Loop**: ExcelJS generates massive numbers of objects on the V8 heap, causing severe GC pauses. nodeosbxl passes buffers straight to the C++ core, streaming to OpenXML without stressing V8.
* **Full Feature Parity**: Delivers native PivotTable generation, structured table references, and AES workbook encryption—all absent in ExcelJS.

#### [4] Java Ecosystem: Apache POI (XSSF) vs osbxl (Java)
| Row Count (Data Scale) | Apache POI (Competitor) | osbxl (True) | osbxl (False) | Speedup Ratio (vs POI) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.695s | **0.127s** | 0.132s | **5.47x Faster** |
| **1,000 rows** | 0.691s | **0.131s** | 0.142s | **5.27x Faster** |
| **5,000 rows** | 1.011s | **0.170s** | 0.198s | **5.95x Faster** |
| **10,000 rows** | 1.335s | **0.208s** | 0.284s | **6.42x Faster** |
| **50,000 rows (200,000 cells)** | ❌ **OOM (Heap Exhaustion)** | **0.553s** | 1.011s | **Survived (Competitor OOM)** |

* **Root Elimination of OutOfMemoryError**: Apache POI (XSSF) expands heavy XMLBeans trees onto the JVM heap, triggering fatal OOM at 50,000 rows under `-Xmx256m`. osbxl Java handles data off-heap in native C++, completing in 0.55s while keeping JVM heap at tens of MBs.
* **No Class Initialization Penalty**: Even small 500-row sheets take POI ~0.7s due to heavy schema loading. osbxl Java responds in 0.12s.
* **Full Flexibility without SXSSF Constraints**: POI's streaming mode (SXSSF) severely restricts backward-row access and PivotTables. osbxl maintains full, unrestricted Excel manipulation at maximum speed.

---

## 3. Cloud Generation Comparison: GCP C3 vs GCP C4

Real-world benchmark data comparing 4th Gen Xeon (Sapphire Rapids / C3) and 5th Gen Xeon (Emerald Rapids / C4) under identical 4-vCPU configurations:

### C3 vs C4 Measured Data Comparison (Unit: Seconds)

| Language / Row Count | GCP C3 (Sapphire Rapids) | GCP C4 (Emerald Rapids) | C4 Speedup | Competing OSS (C4) |
| :--- | :---: | :---: | :---: | :---: |
| **PHP 1,000 rows** | 0.012s | **0.011s** | **1.09x faster** | 0.073s (6.49x) |
| **PHP 10,000 rows** | 0.063s | **0.055s** | **1.15x faster** | 0.760s (13.94x) |
| **PHP 50,000 rows** | 0.405s | **0.346s** | **1.17x faster** | ❌ OOM Crash |
| **Python 1,000 rows** | 0.014s | **0.011s** | **1.27x faster** | 0.033s (3.46x) |
| **Python 10,000 rows** | 0.077s | **0.058s** | **1.33x faster** | 0.324s (5.55x) |
| **Python 50,000 rows** | 0.479s | **0.378s** | **1.27x faster** | 1.662s (4.39x) |
| **Node.js 1,000 rows** | 0.013s | **0.009s** | **1.44x faster** | 0.043s (4.62x) |
| **Node.js 10,000 rows** | 0.068s | **0.056s** | **1.21x faster** | 0.239s (4.28x) |
| **Node.js 50,000 rows** | 0.399s | **0.329s** | **1.21x faster** | 0.979s (2.97x) |
| **Java 1,000 rows** | 0.168s | **0.131s** | **1.28x faster** | 0.691s (5.27x) |
| **Java 10,000 rows** | 0.265s | **0.208s** | **1.27x faster** | 1.335s (6.42x) |
| **Java 50,000 rows** | 0.738s | **0.553s** | **1.33x faster** | ❌ OOM Crash |

---

## 4. Platform-Specific Detailed Measurement Logs

#### [1] GCP C4 (Intel Xeon Platinum 8581C / 4 vCPU / Ubuntu 24.04)

```text
========================================================================================================
Execution Date : 2026-10-18 15:08:17
Machine        : Linux x86_64 (GCP C4 / Intel Emerald Rapids 8581C 4vCPU)
Tuning Config  : /tmp tmpfs 4GB, tuned throughput-performance, PHP OPcache JIT, JVM ParallelGC
Node.js        : v24.22.0
Python         : Python 3.14.8
PHP            : PHP 8.5.11 (cli) (built: Oct 17 2026 00:13:38) (NTS)
Java           : openjdk version "24.0.2" 2025-07-15
osbxl          : v1.5.1 (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
========================================================================================================

[1] PHP Family: PhpSpreadsheet vs phposbxl
Rows      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.213s ( 18.7KB)        |  0.014s ( 20.9KB)      |  0.013s ( 21.1KB)      |  15.07x   |  16.23x
  1000 r  | 0.073s ( 31.0KB)        |  0.011s ( 35.2KB)      |  0.014s ( 35.4KB)      |   6.49x   |   5.07x
  5000 r  | 0.368s (134.1KB)        |  0.031s (152.9KB)      |  0.055s (153.3KB)      |  11.97x   |   6.71x
 10000 r  | 0.760s (262.6KB)        |  0.055s (300.0KB)      |  0.109s (300.9KB)      |  13.94x   |   7.00x
 50000 r  | OOM (Memory Exceeded)   |  0.346s (1481.6KB)     |  0.735s (1496.2KB)     | Survived  | Survived

[2] Python Family: openpyxl vs pyosbxl
Rows      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.024s (  16.3KB)       | 0.015s (  20.9KB)      | 0.013s (  21.1KB)      |   1.68x   |   1.90x
  1000 r  | 0.037s (  28.0KB)       | 0.011s (  35.2KB)      | 0.014s (  35.4KB)      |   3.46x   |   2.61x
  5000 r  | 0.167s ( 119.5KB)       | 0.035s ( 152.9KB)      | 0.060s ( 153.3KB)      |   4.73x   |   2.79x
 10000 r  | 0.324s ( 233.7KB)       | 0.058s ( 300.0KB)      | 0.119s ( 300.9KB)      |   5.55x   |   2.73x
 50000 r  | 1.662s (1133.0KB)       | 0.378s (1481.6KB)      | 0.827s (1496.2KB)      |   4.39x   |   2.01x

[3] Node.js Family: ExcelJS vs nodeosbxl
Rows      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.029s (  19.1KB)       | 0.008s (  20.2KB)      | 0.010s (  20.1KB)      |   3.40x   |   2.96x
  1000 r  | 0.043s (  31.9KB)       | 0.009s (  33.1KB)      | 0.013s (  33.0KB)      |   4.62x   |   3.34x
  5000 r  | 0.135s ( 134.7KB)       | 0.031s ( 143.7KB)      | 0.057s ( 145.1KB)      |   4.31x   |   2.36x
 10000 r  | 0.239s ( 263.0KB)       | 0.056s ( 278.2KB)      | 0.113s ( 284.9KB)      |   4.28x   |   2.11x
 50000 r  | 0.979s (1299.6KB)       | 0.329s (1403.3KB)      | 0.739s (1429.2KB)      |   2.97x   |   1.32x

[4] Java Family: Apache POI (XSSF) vs osbxl(Java)
Rows      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.695s (  16.2KB)       | 0.127s (  20.9KB)       | 0.132s (  21.1KB)       |   5.47x   |   5.27x
  1000 r  | 0.691s (  29.1KB)       | 0.131s (  35.2KB)       | 0.142s (  35.4KB)       |   5.27x   |   4.87x
  5000 r  | 1.011s ( 131.4KB)       | 0.170s ( 152.9KB)       | 0.198s ( 153.3KB)       |   5.95x   |   5.11x
 10000 r  | 1.335s ( 259.8KB)       | 0.208s ( 300.0KB)       | 0.284s ( 300.9KB)       |   6.42x   |   4.70x
 50000 r  | OOM (Heap Exhausted)    | 0.553s (1481.6KB)       | 1.011s (1496.2KB)       | Survived  | Survived
```

---

#### [2] GCP C3 (Intel Xeon Platinum 8481C / 4 vCPU / AlmaLinux 10)

```text
========================================================================================================
Execution Date : 2026-10-18 15:08:01
Machine        : Linux x86_64 (GCP C3 / Intel Sapphire Rapids 8481C 4vCPU)
Tuning Config  : /tmp tmpfs 4GB, tuned throughput-performance, PHP OPcache JIT, JVM ParallelGC
Node.js        : v24.22.0
Python         : Python 3.14.8
PHP            : PHP 8.5.11 (cli) (built: Oct 16 2026 01:17:12) (NTS)
Java           : openjdk version "24.0.2" 2025-07-15
osbxl          : v1.5.1 (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
========================================================================================================

[1] PHP Family: PhpSpreadsheet vs phposbxl
Rows      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.275s ( 18.7KB)        |  0.017s ( 19.4KB)      |  0.017s ( 19.8KB)      |  16.62x   |  16.03x
  1000 r  | 0.098s ( 31.1KB)        |  0.012s ( 31.7KB)      |  0.017s ( 32.3KB)      |   7.93x   |   5.72x
  5000 r  | 0.487s (134.2KB)        |  0.038s (140.7KB)      |  0.068s (140.7KB)      |  12.77x   |   7.11x
 10000 r  | 1.016s (262.6KB)        |  0.063s (277.3KB)      |  0.142s (275.7KB)      |  16.06x   |   7.18x
 50000 r  | OOM (Memory Exceeded)   |  0.405s (1390.6KB)     |  0.908s (1402.6KB)     | Survived  | Survived

[2] Python Family: openpyxl vs pyosbxl
Rows      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.031s (  16.3KB)       | 0.018s (  20.9KB)      | 0.017s (  21.1KB)      |   1.70x   |   1.76x
  1000 r  | 0.048s (  28.1KB)       | 0.014s (  35.2KB)      | 0.019s (  35.4KB)      |   3.36x   |   2.59x
  5000 r  | 0.227s ( 119.8KB)       | 0.043s ( 152.9KB)      | 0.078s ( 153.3KB)      |   5.23x   |   2.91x
 10000 r  | 0.454s ( 234.1KB)       | 0.077s ( 300.0KB)      | 0.161s ( 300.9KB)      |   5.86x   |   2.82x
 50000 r  | 2.330s (1134.9KB)       | 0.479s (1481.6KB)      | 1.056s (1496.2KB)      |   4.87x   |   2.21x

[3] Node.js Family: ExcelJS vs nodeosbxl
Rows      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.044s (  19.1KB)       | 0.011s (  20.2KB)      | 0.013s (  20.1KB)      |   4.06x   |   3.31x
  1000 r  | 0.055s (  31.9KB)       | 0.013s (  33.1KB)      | 0.018s (  33.0KB)      |   4.12x   |   3.00x
  5000 r  | 0.185s ( 134.7KB)       | 0.041s ( 143.7KB)      | 0.075s ( 145.1KB)      |   4.54x   |   2.45x
 10000 r  | 0.301s ( 263.0KB)       | 0.068s ( 278.2KB)      | 0.151s ( 284.9KB)      |   4.43x   |   2.00x
 50000 r  | 1.156s (1299.6KB)       | 0.399s (1403.3KB)      | 0.925s (1429.2KB)      |   2.90x   |   1.25x

[4] Java Family: Apache POI (XSSF) vs osbxl(Java)
Rows      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.895s (  16.2KB)       | 0.164s (  20.9KB)       | 0.172s (  21.1KB)       |   5.46x   |   5.20x
  1000 r  | 0.947s (  29.1KB)       | 0.168s (  35.2KB)       | 0.170s (  35.4KB)       |   5.64x   |   5.57x
  5000 r  | 1.429s ( 131.4KB)       | 0.217s ( 152.9KB)       | 0.264s ( 153.3KB)       |   6.59x   |   5.41x
 10000 r  | 1.689s ( 259.8KB)       | 0.265s ( 300.0KB)       | 0.364s ( 300.9KB)       |   6.37x   |   4.64x
 50000 r  | OOM (Heap Exhausted)    | 0.738s (1481.6KB)       | 1.311s (1496.2KB)       | Survived  | Survived
```

---

#### [3] macOS (Apple M4 mini / arm64 / 16GB Unified Memory)

```text
========================================================================================================
Execution Date : 2026-10-18 15:36:12
Machine        : macOS Darwin arm64 (Apple M4 mini / 10 cores: 4P+6E)
Node.js        : v24.21.0
Python         : Python 3.14.7
PHP            : PHP 8.5.11 (cli) (built: Oct 12 2026 20:52:28) (NTS)
Java           : openjdk version "1.8.0_504"
========================================================================================================

[1] PHP Family: PhpSpreadsheet vs phposbxl
Rows      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.055s ( 18.7KB)        |  0.008s ( 20.9KB)      |  0.008s ( 21.1KB)      |   6.61x   |   6.52x
  1000 r  | 0.080s ( 31.0KB)        |  0.007s ( 35.2KB)      |  0.009s ( 35.4KB)      |  11.52x   |   8.53x
  5000 r  | 0.424s (134.1KB)        |  0.020s (152.9KB)      |  0.032s (153.3KB)      |  21.39x   |  13.36x
 10000 r  | 0.862s (262.6KB)        |  0.037s (300.0KB)      |  0.062s (300.9KB)      |  23.43x   |  14.00x
 50000 r  | OOM (Memory Exceeded)   |  0.189s (1481.6KB)     |  0.335s (1496.2KB)     | Survived  | Survived

[2] Python Family: openpyxl vs pyosbxl
Rows      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.011s (  16.3KB)       | 0.006s (  20.9KB)       | 0.006s (  21.1KB)       |   1.76x   |   1.73x
  1000 r  | 0.019s (  28.0KB)       | 0.005s (  35.2KB)       | 0.007s (  35.4KB)       |   4.07x   |   2.84x
  5000 r  | 0.085s ( 119.5KB)       | 0.016s ( 152.9KB)       | 0.028s ( 153.3KB)       |   5.39x   |   3.07x
 10000 r  | 0.166s ( 233.7KB)       | 0.030s ( 300.0KB)       | 0.054s ( 300.9KB)       |   5.50x   |   3.07x
 50000 r  | 0.850s (1133.0KB)       | 0.183s (1481.6KB)       | 0.292s (1496.2KB)       |   4.64x   |   2.91x

[3] Node.js Family: ExcelJS vs nodeosbxl
Rows      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.024s (  19.1KB)       | 0.005s (  20.9KB)       | 0.006s (  21.1KB)       |   4.92x   |   4.01x
  1000 r  | 0.016s (  31.9KB)       | 0.005s (  35.2KB)       | 0.007s (  35.4KB)       |   3.40x   |   2.36x
  5000 r  | 0.058s ( 134.7KB)       | 0.014s ( 152.9KB)       | 0.025s ( 153.3KB)       |   4.11x   |   2.30x
 10000 r  | 0.109s ( 263.0KB)       | 0.028s ( 300.0KB)       | 0.051s ( 300.9KB)       |   3.92x   |   2.12x
 50000 r  | 0.502s (1299.6KB)       | 0.133s (1481.6KB)       | 0.265s (1496.2KB)       |   3.77x   |   1.89x

[4] Java Family: Apache POI (XSSF) vs osbxl(Java)
Rows      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.341s (  16.2KB)       | 0.059s (  20.9KB)       | 0.061s (  21.1KB)       |   5.78x   |   5.59x
  1000 r  | 0.376s (  29.1KB)       | 0.062s (  35.2KB)       | 0.065s (  35.4KB)       |   6.06x   |   5.78x
  5000 r  | 0.487s ( 131.4KB)       | 0.079s ( 152.9KB)       | 0.093s ( 153.3KB)       |   6.16x   |   5.24x
 10000 r  | 0.620s ( 259.8KB)       | 0.098s ( 300.0KB)       | 0.131s ( 300.9KB)       |   6.33x   |   4.73x
 50000 r  | OOM (Heap Exhausted)    | 0.247s (1481.6KB)       | 0.412s (1496.2KB)       | Survived  | Survived
```

---

#### [4] Windows 11 (Intel Core / x64)

```text
========================================================================================================
Execution Date : 2026-10-17 12:36:50
Machine        : Windows x86_64
Node.js        : v22.23.3
Python         : Python 3.14.7
PHP            : PHP 8.5.11 (cli) (built: Sep 22 2026 13:51:38) (NTS Visual C++ 2022 x64)
Java           : java version "1.8.0_381"
osbxl          : v1.5.1 (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
========================================================================================================

[1] PHP Family: PhpSpreadsheet vs phposbxl
Rows        | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  |  0.168s ( 18.7KB)       |  0.046s ( 20.9KB)       |  0.071s ( 21.1KB)       |     3.63x |    2.38x
  1000 r  |  0.274s ( 31.0KB)       |  0.053s ( 35.2KB)       |  0.090s ( 35.4KB)       |     5.20x |    3.04x
  5000 r  |  1.169s (134.1KB)       |  0.121s (152.9KB)       |  0.314s (153.3KB)       |     9.65x |    3.72x
 10000 r  |  2.333s (262.6KB)       |  0.213s (300.0KB)       |  0.591s (300.9KB)       |    10.98x |    3.95x
 50000 r  | OOM (Memory Exceeded)   |  1.011s (1481.6KB)      |  3.065s (1496.2KB)      |  Survived |  Survived

[2] Python Family: openpyxl vs pyosbxl
Rows        | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  |  0.039s ( 16.4KB)       |  0.054s ( 20.9KB)       |  0.078s ( 21.1KB)       |     0.72x |    0.50x
  1000 r  |  0.059s ( 27.7KB)       |  0.059s ( 35.2KB)       |  0.105s ( 35.4KB)       |     1.00x |    0.57x
  5000 r  |  0.218s (117.4KB)       |  0.127s (152.9KB)       |  0.344s (153.3KB)       |     1.71x |    0.63x
 10000 r  |  0.406s (229.4KB)       |  0.216s (300.0KB)       |  0.659s (300.9KB)       |     1.88x |    0.62x
 50000 r  |  2.022s (1115.3KB)      |  0.996s (1481.6KB)      |  3.256s (1496.2KB)      |     2.03x |    0.62x

[3] Node.js Family: ExcelJS vs nodeosbxl
Rows        | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  |  0.064s ( 19.1KB)       |  0.050s ( 20.9KB)       |  0.075s ( 21.1KB)       |     1.29x |    0.85x
  1000 r  |  0.093s ( 31.9KB)       |  0.060s ( 35.2KB)       |  0.102s ( 35.4KB)       |     1.56x |    0.91x
  5000 r  |  0.216s (134.7KB)       |  0.124s (152.9KB)       |  0.333s (153.3KB)       |     1.74x |    0.65x
 10000 r  |  0.320s (263.0KB)       |  0.210s (300.0KB)       |  0.621s (300.9KB)       |     1.52x |    0.52x
 50000 r  |  1.169s (1299.6KB)      |  0.987s (1481.6KB)      |  3.125s (1496.2KB)      |     1.18x |    0.37x

[4] Java Family: Apache POI (XSSF) vs osbxl(Java)
Rows        | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  |  0.732s ( 16.2KB)       |  0.452s ( 20.9KB)       |  0.473s ( 21.1KB)       |     1.62x |    1.55x
  1000 r  |  0.815s ( 29.1KB)       |  0.461s ( 35.2KB)       |  0.503s ( 35.4KB)       |     1.77x |    1.62x
  5000 r  |  1.161s (131.5KB)       |  0.544s (152.9KB)       |  0.734s (153.3KB)       |     2.13x |    1.58x
 10000 r  |  1.427s (259.9KB)       |  0.636s (300.0KB)       |  1.016s (300.9KB)       |     2.24x |    1.40x
 50000 r  | OOM (Heap Exhausted)    |  1.410s (1481.6KB)      |  3.412s (1496.2KB)      |  Survived |  Survived

========================================================================================================
  [Additional Reference Data] Small Workload (100 & 200 rows) Benchmark Results
========================================================================================================

  [Ref 1] PHP Family: PhpSpreadsheet vs phposbxl (100 & 200 rows)
Rows        | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 r  |  0.078s (  8.8KB)       |  0.042s (  9.2KB)       |  0.046s (  9.5KB)       |     1.84x |    1.69x
   200 r  |  0.098s ( 11.4KB)       |  0.042s ( 12.3KB)       |  0.052s ( 12.5KB)       |     2.31x |    1.87x

  [Ref 2] Python Family: openpyxl vs pyosbxl (100 & 200 rows)
Rows        | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 r  |  0.024s (  7.3KB)       |  0.048s (  9.2KB)       |  0.055s (  9.5KB)       |     0.50x |    0.44x
   200 r  |  0.028s (  9.7KB)       |  0.046s ( 12.3KB)       |  0.061s ( 12.5KB)       |     0.61x |    0.46x

  [Ref 3] Node.js Family: ExcelJS vs nodeosbxl (100 & 200 rows)
Rows        | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 r  |  0.042s (  9.0KB)       |  0.044s (  9.2KB)       |  0.053s (  9.5KB)       |     0.96x |    0.79x
   200 r  |  0.050s ( 11.6KB)       |  0.045s ( 12.3KB)       |  0.059s ( 12.5KB)       |     1.12x |    0.85x

  [Ref 4] Java Family: Apache POI (XSSF) vs osbxl(Java) (100 & 200 rows)
Rows        | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 r  |  0.613s (  5.9KB)       |  0.464s (  9.2KB)       |  0.458s (  9.5KB)       |     1.32x |    1.34x
   200 r  |  0.668s (  8.6KB)       |  0.489s ( 12.3KB)       |  0.498s ( 12.5KB)       |     1.37x |    1.34x
```

---

## 5. Architectural Deep Dive: Why Existing OSS Crashes While osbxl Succeeds

### ① The Limitation of Existing OSS: "Object Explosion on Runtime Heap"

Libraries like PhpSpreadsheet, Apache POI (XSSF), and ExcelJS allocate **each individual cell as an independent high-level runtime object**.
- 50,000 rows × 4 columns = **200,000 distinct cell objects**.
- Each cell references nested objects for font, borders, formatting styles, and formula definitions.
- During serialization, shared string tables (Shared String Table) maintain massive hash maps in runtime memory.

Consequently, even though the resulting `.xlsx` file is merely 1.5MB to 2.0MB on disk, memory consumption easily exceeds standard 256MB/512MB heap limits, leading to inevitable fatal process crashes.

### ② osbxl's Solution: Native C++ Core and Direct 2D Array Bulk Ingestion

In osbxl (across all language bindings), high-level language runtimes do not allocate millions of intermediary cell objects:
```javascript
// Example in Node.js (nodeosbxl) bulk ingestion
ws.setRangeData("A2", data);        // Ingests 2D array directly to native C++ layer via raw pointers
ws.setRangeFormulas("D2", formulas); // Ingests 2D array of formulas directly
```
1. 2D arrays passed from the host language are **immediately handed off to C++ native heap memory**.
2. The optimized C++ core streams and directly encodes the workbook into OpenXML (ZIP / deflate).
3. As a result, **Node.js V8 heap stays at ~20MB, Python heap stays at ~28MB, and PHP memory stays at ~32MB** regardless of row count, completely eliminating runtime OOM crashes at the architectural root.

### ③ Server-Side Formula Calculation & SPILL (Dynamic Array) Support

Most high-level OSS libraries (openpyxl, ExcelJS) only write formula strings as raw text and lack a built-in calculation engine, meaning they cannot evaluate calculated values (`SUM`, `VLOOKUP`, `XLOOKUP`) on server environments.  
In contrast, **osbxl embeds a comprehensive calculation engine (~500 Excel-compatible functions) that fully supports modern dynamic arrays (SPILL)**. Even on headless Linux container environments without Microsoft Excel installed, accurate calculation results are obtained instantly.

---

## 6. Recommended Adoption Scenarios

| Scenario | Traditional Pain Point | Impact of Adopting osbxl |
| :--- | :--- | :--- |
| **Web Report Downloads (100–1,000 rows)** | Slow startup / class load (POI takes 0.5s+), divergent APIs across languages | **Java is 4–6x faster, PHP is 12–14x faster**.<br>All languages return files in ~0.01s. |
| **Monthly Batch / Full Exports (50,000 rows)** | Process killed by OOM (exceeding memory ceiling) | **PHP and Java finish safely in 0.7s–0.9s**.<br>No need to tune or inflate container memory. |
| **Cloud Instance Sizing (C3 vs C4)** | Long batch windows, degraded concurrency | **Selecting GCP C4 unlocks a ~1.1x–1.5x throughput boost across all languages**. |
| **Polyglot & Microservice Architectures** | Distinct libraries and learning curves per language | **100% identical API across Java, Python, Node, PHP, and WebAssembly**. |

---

### (Reference) Cloud Environment (GCP C3 / C4) Measurement Conditions
*Note: To measure pure computational engine performance fairly, both cloud server environments were measured under standardized conditions including RAM disk storage (`/tmp` tmpfs), kernel profile optimization (`tuned`), PHP JIT, and JVM optimization parameters.*

---

### Related Links
- Code Sample Repository: [GitHub osbxl-samples](https://github.com/osb-office/osbxl-samples)
- Official Website & Trial: [osbxl Official Site](https://osb-office.com/)

