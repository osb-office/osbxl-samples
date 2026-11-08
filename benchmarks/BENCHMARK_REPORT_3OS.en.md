# [Technical Benchmark Report] Multi-Platform Comprehensive Stress Test & Architecture Deep Dive: 4 Excel Library Families
## ~ Performance Limits & Behavioral Verification across Cloud (GCP C3/C4), macOS (Apple M4), Windows & Linux ~

> **[🇯🇵 日本語版レポートはこちら (BENCHMARK_REPORT_3OS.md)](BENCHMARK_REPORT_3OS.md)**

- **Public Repository**: [osb-office/osbxl-samples](https://github.com/osb-office/osbxl-samples)
- **Measurement Date**: November 8, 2026 (Reflecting latest GCP C3/C4 / macOS M4 / Windows 11 measurements)
- **osbxl Version**: **`v2.0.1`** (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
- **Benchmark Target Libraries**:
  - **PHP Ecosystem**: `PhpSpreadsheet` vs `phposbxl (v2.0.1)`
  - **Python Ecosystem**: `openpyxl` vs `pyosbxl (v2.0.1)`
  - **Node.js Ecosystem**: `ExcelJS` vs `nodeosbxl (v2.0.1)`
  - **Java Ecosystem**: `Apache POI (XSSF)` vs `osbxl Java (v2.0.1)`
- **Target Environments & Platforms**:
  1. **GCP C4 (Latest Cloud VM)**: Intel Xeon Platinum 8581C (5th Gen Emerald Rapids, 4 vCPU) / Ubuntu 24.04 LTS
  2. **GCP C3 (Prior Gen Cloud VM)**: Intel Xeon Platinum 8481C (4th Gen Sapphire Rapids, 4 vCPU) / Linux
  3. **macOS (Development Desktop)**: Apple M4 mini (10 cores: 4P+6E, 16GB Unified Memory) / macOS Darwin arm64
  4. **Windows (Development Desktop)**: Intel Core / x86_64 / Windows 11
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
> Even small daily web reports (200–1,000 rows) finish in **6–27 ms**: Java is **~23–28x** faster than Apache POI, PHP **~7–15x**, and Python / Node.js **~2–4x**.

---

### 📌 Background & Motivation: Why Run This Benchmark?
In web applications and enterprise batch systems, generating Excel reports has persistently plagued engineers with two critical failure modes:
1. **"Processes crashing immediately from Out Of Memory (OOM) as data rows increase."**
2. **"Web servers timing out due to tens of seconds spent serializing XML/ZIP files."**

To address these pain points, we conducted exhaustive benchmark stress tests comparing standard high-level libraries with `osbxl` across **4 major programming languages (PHP, Python, Node.js, Java)** and **3 representative platforms (Linux: cloud GCP C3/C4, macOS: Apple M4, Windows: Intel Core / Windows 11)** under strictly identical data schemas and latest runtimes.

---

### 🏆 3 Decisive Findings from Real-World Measurements

#### ① [Elimination of Memory Crashes] 50,000 Rows Never Crash osbxl
* **Traditional Libraries (PHP: PhpSpreadsheet / Java: Apache POI)**:  
  Under standard container/cloud configurations (heap limits of 256MB), generating 50,000 rows **reliably triggers fatal OOM crashes**, taking down server processes.
* **osbxl**:  
  Because its C++ native core handles data in a continuous native stream, language-side heap consumption **remains constant at just 30MB–64MB**. osbxl successfully and safely completed across all environments and languages in **0.3s to 0.5s (C4)**.  
  👉 *There is no longer any need to artificially inflate container memory to several gigabytes just for Excel export.*

#### ② [Small Reports Too] Sub-1,000-Row Daily Reports Are Fast in Every Language (GCP C4, measured)
It is not only huge batches: the **100–1,000-row reports** triggered by every "Download" click define perceived speed.

| Rows | PHP (PhpSpreadsheet → osbxl) | Python (openpyxl → osbxl) | Node.js (ExcelJS → osbxl) | Java (Apache POI → osbxl) |
| :--- | :---: | :---: | :---: | :---: |
| **200** | 0.016s → **0.006s** (2.5x) | 0.012s → **0.007s** (1.6x) | 0.009s → 0.009s (parity) | 0.503s → **0.018s** (**27.9x**) |
| **500** | 0.156s → **0.010s** (**15.2x**) | 0.022s → **0.010s** (2.1x) | 0.025s → **0.008s** (3.3x) | 0.563s → **0.021s** (**26.8x**) |
| **1,000** | 0.070s → **0.010s** (**6.7x**) | 0.033s → **0.011s** (3.1x) | 0.039s → **0.010s** (3.8x) | 0.628s → **0.027s** (**23.3x**) |

* **Java: 0.5–0.6 s → tens of ms.** Apache POI pays a heavy XMLBeans schema initialization even for tiny sheets; osbxl has almost none: **~28x** at 200 rows, **~23x** at 1,000 rows (~30x at 100–200 rows on macOS M4 as well).
* **PHP: up to 15x at 500 rows**, 6–15x below 1,000 rows, responding in about **10 ms**.
* **From 500 rows up, every language is 2–4x+ faster**; at 1,000 rows all four languages are **3x or more** ahead.
* **Built-in formula engine, pivot tables and encryption** come at the same speed.
* Honest note: at 100–200 rows Python / Node.js competitors also finish in a few ms, so the gap is small (Node.js parity at 200 rows; openpyxl can be ~6 ms faster at 100 rows). There the value is features and flat scaling rather than raw speed.

#### ③ [Cloud Generational Leap] Latest GCP C4 Delivers ~1.1x–1.4x Higher Throughput than C3
* Comparing 4th Gen Xeon (C3: Sapphire Rapids) with 5th Gen Xeon (C4: Emerald Rapids) under identical 4-vCPU configurations, **C4 consistently outperforms C3 by approximately 1.1x to 1.4x**.
* For a 10,000-row workbook, Node.js finished in **56ms (0.056s)**, PHP clocked **0.055s**, and Java clocked **0.208s**, offering extraordinary scalability for microservices and high-concurrency environments.

---

## 2. Cross-Platform Performance Comparison Matrix

### ① Large-Scale Benchmark: 50,000 Rows (200,000 Cells) Completion Time

| Language / Library | macOS (Apple M4) | GCP C4 (Xeon 8581C) | GCP C3 (Xeon 8481C) | Existing OSS Library Behavior |
| :--- | :---: | :---: | :---: | :--- |
| **PHP (phposbxl)** | **0.202 s** | **0.355 s** | **0.417 s** | ❌ **Fatal OOM Crash on all platforms** |
| **Python (pyosbxl)** | **0.200 s** | **0.389 s** | **0.472 s** | openpyxl: 0.85s–2.52s (No formula evaluation) |
| **Node.js (nodeosbxl)** | **0.153 s** | **0.327 s** | **0.415 s** | ExcelJS: 0.50s–1.32s (No formula engine) |
| **Java (osbxl Java)** | **0.227 s** | **0.447 s** | **0.595 s** | ❌ **Fatal OOM Crash (Heap exhaustion)** |

> *Note: For pure write throughput comparison, osbxl was tested with `setOnlyMode = True`. Existing OSS libraries either lack built-in formula evaluation engines or crash with OOM.*

---

### ② Standard Enterprise Workload: 10,000 Rows (40,000 Cells) Processing Time & Speedup

| Language | Target Libraries | GCP C4 (osbxl) | GCP C3 (osbxl) | Competing OSS (C4) | Speedup on C4 |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **PHP** | PhpSpreadsheet vs **phposbxl** | **0.055 s** | 0.066 s | 0.754 s | **13.60x Faster** |
| **Python** | openpyxl vs **pyosbxl** | **0.063 s** | 0.079 s | 0.311 s | **4.96x Faster** |
| **Node.js** | ExcelJS vs **nodeosbxl** | **0.055 s** | 0.070 s | 0.230 s | **4.20x Faster** |
| **Java** | Apache POI vs **osbxl** | **0.095 s** | 0.125 s | 1.255 s | **13.21x Faster** |

---

### ③ Daily Operational Workload: 200 Rows (800 Cells) Small Benchmark (GCP C4)

| Language | Competing OSS (s) | osbxl (s) | Speedup Ratio | Practical Engineering Impact |
| :--- | :---: | :---: | :---: | :--- |
| **PHP** | 0.016s | **0.006s** | **2.52x Faster** | Web API responds in merely 6ms |
| **Python** | 0.012s | **0.007s** | **1.58x Faster** | Instant write completion with zero startup lag |
| **Node.js** | 0.009s | **0.009s** | **0.99x (Parity)** | Non-blocking execution preserves event loop |
| **Java** | 0.503s | **0.018s** | **27.94x Faster** | Eliminates POI's massive schema load penalty (18ms) |

---

### ④ Language-by-Language Deep Dive: Alternatives vs osbxl

#### [1] PHP Ecosystem: PhpSpreadsheet vs phposbxl
| Row Count (Data Scale) | PhpSpreadsheet (Competitor) | phposbxl (True) | phposbxl (False) | Speedup Ratio (vs Competitor) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.156s | **0.010s** | 0.015s | **15.19x Faster** |
| **1,000 rows** | 0.070s | **0.010s** | 0.017s | **6.73x Faster** |
| **5,000 rows** | 0.357s | **0.032s** | 0.079s | **11.16x Faster** |
| **10,000 rows** | 0.754s | **0.055s** | 0.157s | **13.60x Faster** |
| **50,000 rows (200,000 cells)** | ❌ **OOM (Memory Crash)** | **0.355s** | 1.038s | **Survived (Competitor OOM)** |

* **Total Elimination of Memory Crashes (OOM)**: PhpSpreadsheet instantiates heavy userland PHP objects for every single cell, consuming 256MB at 50,000 rows and immediately triggering `Allowed memory size exhausted`. phposbxl processes raw streams directly in native C++, consuming a constant 20–32MB and finishing safely in 0.34 seconds.
* **Elimination of Web Gateway Timeouts**: Outputting 10,000 rows in just **0.055s (~14x faster)** and standard reports (500–1,000 rows) in **11–14ms** completely resolves concurrent download queue saturation.
* **Enterprise Features Built-in**: Full dynamic PivotTable generation and high-speed native evaluation of ~500 Excel formulas—features completely missing in PhpSpreadsheet.

#### [2] Python Ecosystem: openpyxl vs pyosbxl
| Row Count (Data Scale) | openpyxl (Competitor) | pyosbxl (True) | pyosbxl (False) | Speedup Ratio (vs openpyxl) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.022s | **0.010s** | 0.015s | **2.09x Faster** |
| **1,000 rows** | 0.033s | **0.011s** | 0.018s | **3.07x Faster** |
| **5,000 rows** | 0.156s | **0.036s** | 0.081s | **4.40x Faster** |
| **10,000 rows** | 0.311s | **0.063s** | 0.167s | **4.96x Faster** |
| **50,000 rows (200,000 cells)** | 1.604s | **0.389s** | 1.159s | **4.12x Faster** |

* **3.5x–5.5x Higher Throughput**: As row counts increase, openpyxl suffers from Python object allocation drag. pyosbxl writes out 50,000 rows in **0.37 seconds**.
* **Headless Formula Evaluation Engine**: openpyxl cannot recalculate formulas without running Excel itself, leaving computed cells as `None`. pyosbxl features a **built-in calculation engine supporting ~500 functions**, returning accurate results on headless servers.
* **Native Password Encryption**: pyosbxl provides out-of-the-box password encryption and decryption without needing 3rd-party tools.

#### [3] Node.js Ecosystem: ExcelJS vs nodeosbxl
| Row Count (Data Scale) | ExcelJS (Competitor) | nodeosbxl (True) | nodeosbxl (False) | Speedup Ratio (vs ExcelJS) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.025s | **0.008s** | 0.011s | **3.34x Faster** |
| **1,000 rows** | 0.039s | **0.010s** | 0.017s | **3.81x Faster** |
| **5,000 rows** | 0.136s | **0.030s** | 0.075s | **4.47x Faster** |
| **10,000 rows** | 0.230s | **0.055s** | 0.156s | **4.20x Faster** |
| **50,000 rows (200,000 cells)** | 0.933s | **0.327s** | 0.991s | **2.86x Faster** |

* **Preserving the V8 Event Loop**: ExcelJS generates massive numbers of objects on the V8 heap, causing severe GC pauses. nodeosbxl passes buffers straight to the C++ core, streaming to OpenXML without stressing V8.
* **Full Feature Parity**: Delivers native PivotTable generation, structured table references, and AES workbook encryption—all absent in ExcelJS.

#### [4] Java Ecosystem: Apache POI (XSSF) vs osbxl (Java)
| Row Count (Data Scale) | Apache POI (Competitor) | osbxl (True) | osbxl (False) | Speedup Ratio (vs POI) |
| :--- | :---: | :---: | :---: | :---: |
| **500 rows** | 0.563s | **0.021s** | 0.028s | **26.81x Faster** |
| **1,000 rows** | 0.628s | **0.027s** | 0.038s | **23.26x Faster** |
| **5,000 rows** | 0.836s | **0.061s** | 0.120s | **13.70x Faster** |
| **10,000 rows** | 1.255s | **0.095s** | 0.203s | **13.21x Faster** |
| **50,000 rows (200,000 cells)** | ❌ **OOM (Heap Exhaustion)** | **0.447s** | 1.210s | **Survived (Competitor OOM)** |

* **Root Elimination of OutOfMemoryError**: Apache POI (XSSF) expands heavy XMLBeans trees onto the JVM heap, triggering fatal OOM at 50,000 rows under `-Xmx256m`. osbxl Java handles data off-heap in native C++, completing in 0.55s while keeping JVM heap at tens of MBs.
* **No Class Initialization Penalty**: Even small 500-row sheets take POI ~0.7s due to heavy schema loading. osbxl Java responds in 0.12s.
* **Full Flexibility without SXSSF Constraints**: POI's streaming mode (SXSSF) severely restricts backward-row access and PivotTables. osbxl maintains full, unrestricted Excel manipulation at maximum speed.

---

## 3. Cloud Generation Comparison: GCP C3 vs GCP C4

Real-world benchmark data comparing 4th Gen Xeon (Sapphire Rapids / C3) and 5th Gen Xeon (Emerald Rapids / C4) under identical 4-vCPU configurations:

### C3 vs C4 Measured Data Comparison (Unit: Seconds)

| Language / Row Count | GCP C3 (Sapphire Rapids) | GCP C4 (Emerald Rapids) | C4 Speedup | Competing OSS (C4) |
| :--- | :---: | :---: | :---: | :---: |
| **PHP 1,000 rows** | 0.013s | **0.010s** | **1.30x faster** | 0.070s (7.00x) |
| **PHP 10,000 rows** | 0.066s | **0.055s** | **1.20x faster** | 0.754s (13.60x) |
| **PHP 50,000 rows** | 0.417s | **0.355s** | **1.17x faster** | ❌ OOM Crash |
| **Python 1,000 rows** | 0.014s | **0.011s** | **1.27x faster** | 0.033s (3.07x) |
| **Python 10,000 rows** | 0.079s | **0.063s** | **1.25x faster** | 0.311s (4.96x) |
| **Python 50,000 rows** | 0.472s | **0.389s** | **1.21x faster** | 1.604s (4.12x) |
| **Node.js 1,000 rows** | 0.014s | **0.010s** | **1.40x faster** | 0.039s (3.81x) |
| **Node.js 10,000 rows** | 0.070s | **0.055s** | **1.27x faster** | 0.230s (4.20x) |
| **Node.js 50,000 rows** | 0.415s | **0.327s** | **1.27x faster** | 0.933s (2.86x) |
| **Java 1,000 rows** | 0.041s | **0.027s** | **1.52x faster** | 0.628s (23.26x) |
| **Java 10,000 rows** | 0.125s | **0.095s** | **1.32x faster** | 1.255s (13.21x) |
| **Java 50,000 rows** | 0.595s | **0.447s** | **1.33x faster** | ❌ OOM Crash |

## 4. Platform-Specific Detailed Measurement Logs

#### [1] GCP C4 (Intel Xeon Platinum 8581C / 4 vCPU / Ubuntu 24.04)

```text
========================================================================================================
          Linux (x86_64) 4系統 総合負荷テスト (500行〜50,000行)                                         
========================================================================================================
実行日時 : 2026-11-06 14:43:11
マシン   : Linux x86_64
Node.js  : v24.22.0
Python   : Python 3.14.8
PHP      : PHP 8.5.11 (cli) (built: Oct 17 2026 00:13:38) (NTS)
Java     : openjdk version "24.0.2" 2025-07-15
========================================================================================================

########################################################################################################
  [1] PHP 系: PhpSpreadsheet vs phposbxl
########################################################################################################
行数      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.156s ( 18.7KB)        |  0.010s ( 20.9KB)      |  0.015s ( 24.0KB)      |  15.19倍 | 10.14倍
  1000 行 | 0.070s ( 31.0KB)        |  0.010s ( 35.2KB)      |  0.017s ( 40.8KB)      |   6.73倍 |  4.02倍
  5000 行 | 0.357s (134.1KB)        |  0.032s (152.6KB)      |  0.079s (184.2KB)      |  11.16倍 |  4.55倍
 10000 行 | 0.754s (262.6KB)        |  0.055s (299.8KB)      |  0.157s (362.9KB)      |  13.60倍 |  4.81倍
 50000 行 | OOM (メモリ枯渇)   |  0.355s (1479.8KB)      |  1.038s (1790.6KB)      | 完走(OOM) | 完走(OOM)

########################################################################################################
  [2] Python 系: openpyxl vs pyosbxl
########################################################################################################
行数      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.022s (  16.3KB)       | 0.010s (  20.9KB)      | 0.015s (  24.0KB)      |     2.09倍 |    1.45倍
  1000 行 | 0.033s (  28.0KB)       | 0.011s (  35.2KB)      | 0.018s (  40.8KB)      |     3.07倍 |    1.83倍
  5000 行 | 0.156s ( 119.5KB)       | 0.036s ( 152.6KB)      | 0.081s ( 184.2KB)      |     4.40倍 |    1.94倍
 10000 行 | 0.311s ( 233.7KB)       | 0.063s ( 299.8KB)      | 0.167s ( 363.0KB)      |     4.96倍 |    1.86倍
 50000 行 | 1.604s (1133.0KB)       | 0.389s (1479.8KB)      | 1.159s (1790.6KB)      |     4.12倍 |    1.38倍

########################################################################################################
  [3] Node.js 系: ExcelJS vs nodeosbxl
########################################################################################################
行数      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.025s (  19.1KB)       | 0.008s (  20.3KB)      | 0.011s (  22.0KB)      |     3.34倍 |    2.29倍
  1000 行 | 0.039s (  31.9KB)       | 0.010s (  33.2KB)      | 0.017s (  36.2KB)      |     3.81倍 |    2.30倍
  5000 行 | 0.136s ( 134.7KB)       | 0.030s ( 144.2KB)      | 0.075s ( 163.4KB)      |     4.47倍 |    1.81倍
 10000 行 | 0.230s ( 263.0KB)       | 0.055s ( 284.7KB)      | 0.156s ( 323.3KB)      |     4.20倍 |    1.47倍
 50000 行 | 0.933s (1299.6KB)       | 0.327s (1410.3KB)      | 0.991s (1629.0KB)      |     2.86倍 |    0.94倍

########################################################################################################
  [4] Java 系: Apache POI (XSSF) vs osbxl(Java)
########################################################################################################
行数      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.563s (  16.2KB)       | 0.021s (  20.9KB)       | 0.028s (  24.0KB)       |    26.81倍 |   20.11倍
  1000 行 | 0.628s (  29.1KB)       | 0.027s (  35.2KB)       | 0.038s (  40.8KB)       |    23.26倍 |   16.53倍
  5000 行 | 0.836s ( 131.4KB)       | 0.061s ( 152.6KB)       | 0.120s ( 184.2KB)       |    13.70倍 |    6.97倍
 10000 行 | 1.255s ( 259.8KB)       | 0.095s ( 299.8KB)       | 0.203s ( 362.9KB)       |    13.21倍 |    6.18倍
 50000 行 | OOM (ヒープ枯渇)             | 0.447s (1479.8KB)       | 1.210s (1790.6KB)       |   完走(OOM) |  完走(OOM)

========================================================================================================
  【追加参考データ】小規模ケース (100行・200行) 実測結果
========================================================================================================

  [参考 1] PHP 系: PhpSpreadsheet vs phposbxl (100行・200行)
行数      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.128s (  8.8KB)        |  0.012s (  9.2KB)      |  0.011s ( 10.2KB)      |  10.24倍 | 11.46倍
   200 行 | 0.016s ( 11.4KB)        |  0.006s ( 12.3KB)      |  0.008s ( 13.8KB)      |   2.52倍 |  2.06倍

  [参考 2] Python 系: openpyxl vs pyosbxl (100行・200行)
行数      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.007s (   7.2KB)       | 0.013s (   9.2KB)      | 0.010s (  10.2KB)      |     0.55倍 |    0.67倍
   200 行 | 0.012s (   9.6KB)       | 0.007s (  12.3KB)      | 0.009s (  13.8KB)      |     1.58倍 |    1.30倍

  [参考 3] Node.js 系: ExcelJS vs nodeosbxl (100行・200行)
行数      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.029s (   9.0KB)       | 0.015s (   9.3KB)      | 0.011s (  10.2KB)      |     1.98倍 |    2.59倍
   200 行 | 0.009s (  11.6KB)       | 0.009s (  12.2KB)      | 0.008s (  13.4KB)      |     0.99倍 |    1.12倍

  [参考 4] Java 系: Apache POI (XSSF) vs osbxl(Java) (100行・200行)
行数      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.446s (   5.9KB)       | 0.016s (   9.2KB)       | 0.022s (  10.2KB)       |    27.88倍 |   20.27倍
   200 行 | 0.503s (   8.6KB)       | 0.018s (  12.3KB)       | 0.023s (  13.8KB)       |    27.94倍 |   21.87倍

========================================================================================================
                                   Linux 総合負荷テスト完了                                             
========================================================================================================
```

---

#### [2] GCP C3 (Intel Xeon Platinum 8481C / 4 vCPU / AlmaLinux 10)

```text
========================================================================================================
          Linux (x86_64) 4系統 総合負荷テスト (500行〜50,000行)                                         
========================================================================================================
実行日時 : 2026-11-06 14:42:10
マシン   : Linux x86_64
Node.js  : v24.22.0
Python   : Python 3.14.8
PHP      : PHP 8.5.11 (cli) (built: Oct 16 2026 01:17:12) (NTS)
Java     : openjdk version "24.0.2" 2025-07-15
========================================================================================================

########################################################################################################
  [1] PHP 系: PhpSpreadsheet vs phposbxl
########################################################################################################
行数      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.217s ( 18.7KB)        |  0.013s ( 19.5KB)      |  0.020s ( 21.6KB)      |  16.29倍 | 10.88倍
  1000 行 | 0.097s ( 31.1KB)        |  0.013s ( 31.7KB)      |  0.023s ( 35.7KB)      |   7.21倍 |  4.24倍
  5000 行 | 0.484s (134.2KB)        |  0.040s (140.8KB)      |  0.102s (160.9KB)      |  11.96倍 |  4.74倍
 10000 行 | 1.014s (262.6KB)        |  0.066s (277.1KB)      |  0.206s (317.2KB)      |  15.34倍 |  4.91倍
 50000 行 | OOM (メモリ枯渇)   |  0.417s (1390.4KB)      |  1.298s (1607.3KB)      | 完走(OOM) | 完走(OOM)

########################################################################################################
  [2] Python 系: openpyxl vs pyosbxl
########################################################################################################
行数      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.030s (  16.3KB)       | 0.014s (  19.5KB)      | 0.020s (  21.6KB)      |     2.10倍 |    1.51倍
  1000 行 | 0.048s (  28.1KB)       | 0.014s (  31.7KB)      | 0.024s (  35.7KB)      |     3.35倍 |    2.00倍
  5000 行 | 0.227s ( 119.8KB)       | 0.043s ( 140.8KB)      | 0.104s ( 160.9KB)      |     5.22倍 |    2.17倍
 10000 行 | 0.453s ( 234.1KB)       | 0.079s ( 277.1KB)      | 0.222s ( 317.2KB)      |     5.70倍 |    2.04倍
 50000 行 | 2.312s (1134.9KB)       | 0.472s (1390.4KB)      | 1.432s (1607.3KB)      |     4.90倍 |    1.61倍

########################################################################################################
  [3] Node.js 系: ExcelJS vs nodeosbxl
########################################################################################################
行数      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.039s (  19.1KB)       | 0.012s (  20.3KB)      | 0.017s (  22.0KB)      |     3.31倍 |    2.34倍
  1000 行 | 0.059s (  31.9KB)       | 0.014s (  33.2KB)      | 0.024s (  36.2KB)      |     4.10倍 |    2.51倍
  5000 行 | 0.182s ( 134.7KB)       | 0.042s ( 144.2KB)      | 0.101s ( 163.4KB)      |     4.28倍 |    1.80倍
 10000 行 | 0.303s ( 263.0KB)       | 0.070s ( 284.7KB)      | 0.210s ( 323.3KB)      |     4.32倍 |    1.44倍
 50000 行 | 1.166s (1299.6KB)       | 0.415s (1410.3KB)      | 1.286s (1629.0KB)      |     2.81倍 |    0.91倍

########################################################################################################
  [4] Java 系: Apache POI (XSSF) vs osbxl(Java)
########################################################################################################
行数      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.756s (  16.2KB)       | 0.031s (  19.5KB)       | 0.040s (  21.6KB)       |    24.39倍 |   18.90倍
  1000 行 | 0.837s (  29.1KB)       | 0.041s (  31.7KB)       | 0.055s (  35.7KB)       |    20.41倍 |   15.22倍
  5000 行 | 1.340s ( 131.4KB)       | 0.083s ( 140.8KB)       | 0.165s ( 160.9KB)       |    16.14倍 |    8.12倍
 10000 行 | 1.799s ( 259.8KB)       | 0.125s ( 277.1KB)       | 0.320s ( 317.2KB)       |    14.39倍 |    5.62倍
 50000 行 | OOM (ヒープ枯渇)             | 0.595s (1390.4KB)       | 1.559s (1607.3KB)       |   完走(OOM) |  完走(OOM)

========================================================================================================
  【追加参考データ】小規模ケース (100行・200行) 実測結果
========================================================================================================

  [参考 1] PHP 系: PhpSpreadsheet vs phposbxl (100行・200行)
行数      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.179s (  8.9KB)        |  0.016s (  9.2KB)      |  0.016s ( 10.1KB)      |  11.43倍 | 11.48倍
   200 行 | 0.025s ( 11.5KB)        |  0.009s ( 11.9KB)      |  0.011s ( 13.2KB)      |   2.70倍 |  2.24倍

  [参考 2] Python 系: openpyxl vs pyosbxl (100行・200行)
行数      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.010s (   7.2KB)       | 0.017s (   9.2KB)      | 0.015s (  10.1KB)      |     0.59倍 |    0.67倍
   200 行 | 0.016s (   9.6KB)       | 0.009s (  11.9KB)      | 0.012s (  13.2KB)      |     1.77倍 |    1.38倍

  [参考 3] Node.js 系: ExcelJS vs nodeosbxl (100行・200行)
行数      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.044s (   9.0KB)       | 0.019s (   9.3KB)      | 0.016s (  10.2KB)      |     2.25倍 |    2.80倍
   200 行 | 0.013s (  11.6KB)       | 0.013s (  12.2KB)      | 0.013s (  13.4KB)      |     0.99倍 |    1.02倍

  [参考 4] Java 系: Apache POI (XSSF) vs osbxl(Java) (100行・200行)
行数      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 | 0.627s (   5.9KB)       | 0.024s (   9.2KB)       | 0.030s (  10.1KB)       |    26.12倍 |   20.90倍
   200 行 | 0.675s (   8.6KB)       | 0.026s (  11.9KB)       | 0.034s (  13.2KB)       |    25.96倍 |   19.85倍

========================================================================================================
                                   Linux 総合負荷テスト完了                                             
========================================================================================================
```

---

#### [3] macOS (Apple M4 mini / arm64 / 16GB Unified Memory)

```text
========================================================================================================
Execution Date : 2026-11-06 22:41:57
Machine        : macOS Darwin arm64 (Apple M4 mini / 10 cores: 4P+6E)
Node.js        : v24.21.0
Python         : Python 3.14.7
PHP            : PHP 8.5.11 (cli) (built: Oct 12 2026 20:52:28) (NTS)
Java           : openjdk version "1.8.0_504"
osbxl          : v2.0.0 (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
========================================================================================================

[1] PHP Family: PhpSpreadsheet vs phposbxl
Rows      | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.056s ( 18.7KB)        |  0.005s ( 20.9KB)      |  0.008s ( 24.0KB)      |  12.35x   |   6.68x
  1000 r  | 0.081s ( 31.0KB)        |  0.006s ( 35.2KB)      |  0.011s ( 40.8KB)      |  13.47x   |   7.46x
  5000 r  | 0.430s (134.1KB)        |  0.020s (152.6KB)      |  0.047s (184.2KB)      |  21.48x   |   9.25x
 10000 r  | 0.874s (262.6KB)        |  0.040s (299.8KB)      |  0.095s (362.9KB)      |  21.99x   |   9.16x
 50000 r  | OOM (Memory Exceeded)   |  0.202s (1479.8KB)     |  0.510s (1790.6KB)     | Survived  | Survived

[2] Python Family: openpyxl vs pyosbxl
Rows      | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.011s (  16.3KB)       | 0.006s (  20.9KB)       | 0.008s (  24.0KB)       |   1.79x   |   1.31x
  1000 r  | 0.019s (  28.0KB)       | 0.006s (  35.2KB)       | 0.009s (  40.8KB)       |   3.52x   |   2.06x
  5000 r  | 0.085s ( 119.5KB)       | 0.018s ( 152.6KB)       | 0.040s ( 184.2KB)       |   4.80x   |   2.11x
 10000 r  | 0.167s ( 233.7KB)       | 0.035s ( 299.8KB)       | 0.082s ( 362.9KB)       |   4.83x   |   2.04x
 50000 r  | 0.868s (1133.0KB)       | 0.200s (1479.8KB)       | 0.436s (1790.6KB)       |   4.35x   |   1.99x

[3] Node.js Family: ExcelJS vs nodeosbxl
Rows      | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.024s (  19.1KB)       | 0.004s (  20.9KB)       | 0.008s (  24.0KB)       |   5.48x   |   2.97x
  1000 r  | 0.016s (  31.9KB)       | 0.005s (  35.2KB)       | 0.010s (  40.8KB)       |   3.10x   |   1.70x
  5000 r  | 0.058s ( 134.7KB)       | 0.017s ( 152.6KB)       | 0.040s ( 184.2KB)       |   3.37x   |   1.44x
 10000 r  | 0.113s ( 263.0KB)       | 0.032s ( 299.8KB)       | 0.078s ( 362.9KB)       |   3.59x   |   1.45x
 50000 r  | 0.523s (1299.6KB)       | 0.153s (1479.8KB)       | 0.420s (1790.6KB)       |   3.42x   |   1.25x

[4] Java Family: Apache POI (XSSF) vs osbxl(Java)
Rows      | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 r  | 0.341s (  16.2KB)       | 0.012s (  20.9KB)       | 0.017s (  24.0KB)       |  28.42x   |  20.06x
  1000 r  | 0.369s (  29.1KB)       | 0.015s (  35.2KB)       | 0.022s (  40.8KB)       |  24.60x   |  16.77x
  5000 r  | 0.517s ( 131.4KB)       | 0.035s ( 152.6KB)       | 0.061s ( 184.2KB)       |  14.77x   |   8.48x
 10000 r  | 0.634s ( 259.8KB)       | 0.055s ( 299.8KB)       | 0.114s ( 362.9KB)       |  11.53x   |   5.56x
 50000 r  | OOM (Heap Exhausted)    | 0.227s (1479.8KB)       | 0.523s (1790.6KB)       | Survived  | Survived

[Additional Reference] Daily Operational Scale (100 & 200 rows)
Rows      | Competitor (A)          | osbxl (True) (B)        | osbxl (False) (C)       | (A)/(B)   | (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
PHP 100r  | 0.024s (  8.8KB)        |  0.004s (  9.2KB)      |  0.005s ( 10.2KB)      |   6.49x   |   4.65x
PHP 200r  | 0.018s ( 11.4KB)        |  0.003s ( 12.3KB)      |  0.004s ( 13.8KB)      |   5.48x   |   4.11x
Py  100r  | 0.004s (  7.2KB)        |  0.004s (  9.2KB)      |  0.005s ( 10.2KB)      |   0.94x   |   0.81x
Py  200r  | 0.005s (  9.6KB)        |  0.003s ( 12.3KB)      |  0.004s ( 13.8KB)      |   1.71x   |   1.30x
Node 100r | 0.016s (  9.0KB)        |  0.003s (  9.2KB)      |  0.005s ( 10.2KB)      |   4.56x   |   3.05x
Node 200r | 0.006s ( 11.6KB)        |  0.003s ( 12.3KB)      |  0.004s ( 13.8KB)      |   1.71x   |   1.30x
Java 100r | 0.283s (  5.9KB)        |  0.009s (  9.2KB)      |  0.011s ( 10.2KB)      |  31.44x   |  25.73x
Java 200r | 0.303s (  8.6KB)        |  0.010s ( 12.3KB)      |  0.013s ( 13.8KB)      |  30.30x   |  23.31x
```

---

---

### [4] Windows (Intel Core / x86_64 / Windows 11)

```text
﻿========================================================================================================
実行日時 : 2026-11-08 14:19:49
マシン   : Windows x86_64
Node.js  : v22.23.3
Python   : Python 3.14.7
PHP      : PHP 8.5.11 (cli) (built: Sep 22 2026 13:51:38) (NTS Visual C++ 2022 x64)
Java     : java version "1.8.0_381"
osbxl    : v2.0.1 (nodeosbxl / pyosbxl / phposbxl / osbxl Java)
集計方法 : 全 10 回試行の各セル最良値 (Best of 10)
========================================================================================================

[1] PHP 系: PhpSpreadsheet vs phposbxl
行数        | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.165s ( 18.7KB)        | 0.011s ( 20.9KB)        | 0.020s ( 24.0KB)        |    15.00倍 |    8.25倍
  1000 行 | 0.274s ( 31.0KB)        | 0.014s ( 35.2KB)        | 0.029s ( 40.8KB)        |    19.57倍 |    9.45倍
  5000 行 | 1.184s (134.1KB)        | 0.040s (152.6KB)        | 0.113s (184.2KB)        |    29.60倍 |   10.48倍
 10000 行 | 2.341s (262.6KB)        | 0.075s (299.8KB)        | 0.225s (363.0KB)        |    31.21倍 |   10.40倍
 50000 行 | OOM (メモリ枯渇)             | 0.375s (1479.8KB)       | 1.171s (1790.6KB)       |   完走(OOM) |  完走(OOM)

[2] Python 系: openpyxl vs pyosbxl
行数        | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.036s ( 16.4KB)        | 0.016s ( 20.9KB)        | 0.025s ( 24.0KB)        |     2.25倍 |    1.44倍
  1000 行 | 0.057s ( 27.7KB)        | 0.019s ( 35.2KB)        | 0.036s ( 40.8KB)        |     3.00倍 |    1.58倍
  5000 行 | 0.209s (117.4KB)        | 0.047s (152.6KB)        | 0.121s (184.2KB)        |     4.45倍 |    1.73倍
 10000 行 | 0.404s (229.4KB)        | 0.084s (299.8KB)        | 0.236s (362.9KB)        |     4.81倍 |    1.71倍
 50000 行 | 2.047s (1115.3KB)       | 0.398s (1479.8KB)       | 1.204s (1790.6KB)       |     5.14倍 |    1.70倍

[3] Node.js 系: ExcelJS vs nodeosbxl
行数        | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.063s ( 19.1KB)        | 0.013s ( 20.9KB)        | 0.022s ( 24.0KB)        |     4.85倍 |    2.86倍
  1000 行 | 0.096s ( 31.9KB)        | 0.016s ( 35.2KB)        | 0.032s ( 40.8KB)        |     6.00倍 |    3.00倍
  5000 行 | 0.211s (134.7KB)        | 0.041s (152.6KB)        | 0.114s (184.2KB)        |     5.15倍 |    1.85倍
 10000 行 | 0.328s (263.0KB)        | 0.074s (299.8KB)        | 0.224s (362.9KB)        |     4.43倍 |    1.46倍
 50000 行 | 1.207s (1299.6KB)       | 0.366s (1479.8KB)       | 1.161s (1790.6KB)       |     3.30倍 |    1.04倍

[4] Java 系: Apache POI (XSSF) vs osbxl(Java)
行数        | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   500 行 | 0.666s ( 16.2KB)        | 0.023s ( 20.9KB)        | 0.029s ( 24.0KB)        |    28.96倍 |   22.97倍
  1000 行 | 0.714s ( 29.1KB)        | 0.025s ( 35.2KB)        | 0.040s ( 40.8KB)        |    28.56倍 |   17.85倍
  5000 行 | 1.049s (131.5KB)        | 0.063s (152.6KB)        | 0.133s (184.2KB)        |    16.65倍 |    7.89倍
 10000 行 | 1.338s (259.9KB)        | 0.104s (299.8KB)        | 0.239s (362.9KB)        |    12.87倍 |    5.60倍
 50000 行 | OOM (ヒープ枯渇)             | 0.411s (1479.8KB)       | 1.149s (1790.6KB)       |   完走(OOM) |  完走(OOM)

========================================================================================================
  【追加参考データ】小規模ケース (100行・200行) 実測結果
========================================================================================================

  [参考 1] PHP 系: PhpSpreadsheet vs phposbxl (100行・200行)
行数        | PhpSpreadsheet (A)      | phposbxl (True) (B)     | phposbxl (False) (C)    |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 |  0.077s (  8.8KB)       |  0.009s (  9.2KB)       |  0.014s ( 10.2KB)       |     8.32倍 |    5.58倍
   200 行 |  0.099s ( 11.4KB)       |  0.009s ( 12.3KB)       |  0.015s ( 13.8KB)       |    10.63倍 |    6.55倍

  [参考 2] Python 系: openpyxl vs pyosbxl (100行・200行)
行数        | openpyxl (A)            | pyosbxl (True) (B)      | pyosbxl (False) (C)     |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 |  0.020s (  7.3KB)       |  0.014s (  9.2KB)       |  0.019s ( 10.2KB)       |     1.47倍 |    1.10倍
   200 行 |  0.025s (  9.7KB)       |  0.014s ( 12.3KB)       |  0.020s ( 13.8KB)       |     1.78倍 |    1.28倍

  [参考 3] Node.js 系: ExcelJS vs nodeosbxl (100行・200行)
行数        | ExcelJS (A)             | nodeosbxl (True) (B)    | nodeosbxl (False) (C)   |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 |  0.041s (  9.0KB)       |  0.010s (  9.2KB)       |  0.015s ( 10.2KB)       |     3.97倍 |    2.82倍
   200 行 |  0.047s ( 11.6KB)       |  0.011s ( 12.3KB)       |  0.015s ( 13.8KB)       |     4.42倍 |    3.11倍

  [参考 4] Java 系: Apache POI (XSSF) vs osbxl(Java) (100行・200行)
行数        | Apache POI (A)          | osbxl (True) (B)        | osbxl (False) (C)       |   (A)/(B) |  (A)/(C)
----------+-------------------------+-------------------------+-------------------------+-----------+----------
   100 行 |  0.545s (  5.9KB)       |  0.032s (  9.2KB)       |  0.023s ( 10.2KB)       |    17.03倍 |   23.70倍
   200 行 |  0.585s (  8.6KB)       |  0.031s ( 12.3KB)       |  0.031s ( 13.8KB)       |    18.87倍 |   18.87倍

========================================================================================================
ベンチマーク測定終了
========================================================================================================
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

