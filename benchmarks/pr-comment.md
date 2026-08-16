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
| **`@hokkyss/pptx-core`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |
| **`@hokkyss/pptx-reader`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |
| **`@hokkyss/pptx-writer`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |
| **`@hokkyss/pptx (Full SDK)`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |

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
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **9,656.0** | 0.1036 ms | 0.1459 ms | $\pm$ 0.23% |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **4,708,741.4** | 0.0002 ms | 0.0004 ms | $\pm$ 2.2% |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,643,666.2** | 0.0006 ms | 0.0009 ms | $\pm$ 0.25% |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **882,755.9** | 0.0011 ms | 0.0018 ms | $\pm$ 0.43% |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,378.8** | 0.7253 ms | 2.6588 ms | $\pm$ 5.02% |
| **Single Slide with Shape & Text** | `pptx` | **1,832.0** | 0.5458 ms | 0.7706 ms | $\pm$ 0.8% |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **1,209.2** | 0.827 ms | 4.7131 ms | $\pm$ 6.5% |
| **DrawingML Column Chart Generation** | `pptx` | **685.0** | 1.4599 ms | 11.0127 ms | $\pm$ 26.05% |
| **10-Slide Full Enterprise Deck** | `pptx` | **278.5** | 3.5909 ms | 10.0702 ms | $\pm$ 7.23% |
| **50-Slide Batch Scale** | `pptx` | **148.5** | 6.7323 ms | 20.7122 ms | $\pm$ 8.67% |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
