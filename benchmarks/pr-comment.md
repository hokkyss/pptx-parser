## 🚀 CI Benchmark & Bundle Size Telemetry

> Automated evaluation for commit hash: `local` | Target: `main`

### 📦 Production Bundle Footprint (Bundlephobia-Aligned)

| Package | Minified | Gzip Size | Brotli Size | DTS Size | Budget | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | 1.99 KB | **622.00 B** | 546.00 B | 305.00 B | 5 KB | ✅ PASS |
| **`@hokkyss/pptx-reader`** | 97.60 KB | **26.53 KB** | 23.47 KB | 608.00 B | 30 KB | ✅ PASS |
| **`@hokkyss/pptx-writer`** | 128.57 KB | **32.26 KB** | 28.22 KB | 107.00 B | 35 KB | ✅ PASS |
| **`@hokkyss/pptx (Full SDK)`** | 196.33 KB | **50.30 KB** | 44.52 KB | 895.00 B | 70 KB | ✅ PASS |

### 📊 Baseline Comparison (Δ Delta)

| Package | Δ Minified | Δ Gzip | Δ Brotli | Status |
| :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | +401 B (+24.6%) | +185 B (+42.3%) | +164 B (+42.9%) | 🔴 INCREASED |
| **`@hokkyss/pptx-reader`** | +870 B (+0.9%) | +325 B (+1.2%) | +224 B (+0.9%) | 🔴 INCREASED |
| **`@hokkyss/pptx-writer`** | +292 B (+0.2%) | +87 B (+0.3%) | +101 B (+0.4%) | 🔴 INCREASED |
| **`@hokkyss/pptx (Full SDK)`** | +622 B (+0.3%) | +288 B (+0.6%) | +275 B (+0.6%) | 🔴 INCREASED |

### 🌲 First-Party Tree-Shaking Efficacy (Client Bundling)

| Import Scenario | Minified | Gzip Size | Brotli Size | Compression Ratio |
| :--- | :--- | :--- | :--- | :--- |
| **Core Units & Types Only** | 502.00 B | **262.00 B** | 231.00 B | 47.8% |
| **Reader Parser Only** | 32.20 KB | **9.08 KB** | 8.05 KB | 71.8% |
| **Writer Serializer Only** | 60.97 KB | **13.16 KB** | 11.55 KB | 78.4% |
| **Full Presentation Engine** | 28.98 KB | **7.08 KB** | 6.25 KB | 75.6% |

### ⚡ Microsecond Performance & Latency Matrix (Vitest)

| Benchmark Workload | Package | Frequency (ops/s) | Mean Latency | p99 Latency | Margin of Error |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **9,312.3** | 0.1074 ms | 0.1785 ms | $\pm$ 0.4% |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **4,805,006.0** | 0.0002 ms | 0.0003 ms | $\pm$ 0.87% |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,684,633.0** | 0.0006 ms | 0.0008 ms | $\pm$ 0.14% |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **925,100.7** | 0.0011 ms | 0.0014 ms | $\pm$ 0.17% |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,411.3** | 0.7086 ms | 2.6033 ms | $\pm$ 3.83% |
| **Single Slide with Shape & Text** | `pptx` | **1,200.9** | 0.8327 ms | 3.963 ms | $\pm$ 6.11% |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **1,189.7** | 0.8405 ms | 4.146 ms | $\pm$ 7.94% |
| **DrawingML Column Chart Generation** | `pptx` | **972.7** | 1.0281 ms | 4.2158 ms | $\pm$ 14.12% |
| **10-Slide Full Enterprise Deck** | `pptx` | **165.3** | 6.0497 ms | 47.649 ms | $\pm$ 24.13% |
| **50-Slide Batch Scale** | `pptx` | **120.4** | 8.3049 ms | 49.148 ms | $\pm$ 21.55% |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
