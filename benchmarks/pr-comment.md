## 🚀 CI Benchmark & Bundle Size Telemetry

> Automated evaluation for commit hash: `local` | Target: `main`

### 📦 Production Bundle Footprint (Bundlephobia-Aligned)

| Package | Minified | Gzip Size | Brotli Size | DTS Size | Budget | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | 1.99 KB | **622.00 B** | 546.00 B | 305.00 B | 5 KB | ✅ PASS |
| **`@hokkyss/pptx-reader`** | 98.29 KB | **26.81 KB** | 23.72 KB | 608.00 B | 30 KB | ✅ PASS |
| **`@hokkyss/pptx-writer`** | 130.23 KB | **32.65 KB** | 28.54 KB | 107.00 B | 35 KB | ✅ PASS |
| **`@hokkyss/pptx (Full SDK)`** | 199.04 KB | **51.05 KB** | 45.16 KB | 997.00 B | 70 KB | ✅ PASS |

### 📊 Baseline Comparison (Δ Delta)

| Package | Δ Minified | Δ Gzip | Δ Brotli | Status |
| :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | +401 B (+24.6%) | +185 B (+42.3%) | +164 B (+42.9%) | 🔴 INCREASED |
| **`@hokkyss/pptx-reader`** | +1.54 KB (+1.6%) | +610 B (+2.3%) | +484 B (+2.0%) | 🔴 INCREASED |
| **`@hokkyss/pptx-writer`** | +1.94 KB (+1.5%) | +491 B (+1.5%) | +434 B (+1.5%) | 🔴 INCREASED |
| **`@hokkyss/pptx (Full SDK)`** | +3.31 KB (+1.7%) | +1.03 KB (+2.1%) | +933 B (+2.1%) | 🔴 INCREASED |

### 🌲 First-Party Tree-Shaking Efficacy (Client Bundling)

| Import Scenario | Minified | Gzip Size | Brotli Size | Compression Ratio |
| :--- | :--- | :--- | :--- | :--- |
| **Core Units & Types Only** | 502.00 B | **262.00 B** | 231.00 B | 47.8% |
| **Reader Parser Only** | 32.89 KB | **9.36 KB** | 8.30 KB | 71.6% |
| **Writer Serializer Only** | 62.63 KB | **13.59 KB** | 11.89 KB | 78.3% |
| **Full Presentation Engine** | 29.33 KB | **7.13 KB** | 6.30 KB | 75.7% |

### ⚡ Microsecond Performance & Latency Matrix (Vitest)

| Benchmark Workload | Package | Frequency (ops/s) | Mean Latency | p99 Latency | Margin of Error |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **8,810.0** | 0.1135 ms | 0.2264 ms | $\pm$ 0.73% |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **4,551,807.0** | 0.0002 ms | 0.0004 ms | $\pm$ 0.97% |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,521,302.3** | 0.0007 ms | 0.0015 ms | $\pm$ 0.88% |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **903,729.3** | 0.0011 ms | 0.0016 ms | $\pm$ 0.33% |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,270.0** | 0.7874 ms | 3.14 ms | $\pm$ 5.12% |
| **Single Slide with Shape & Text** | `pptx` | **1,519.7** | 0.658 ms | 1.826 ms | $\pm$ 3.5% |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **1,223.5** | 0.8173 ms | 4.3238 ms | $\pm$ 12.71% |
| **DrawingML Column Chart Generation** | `pptx` | **1,311.0** | 0.7628 ms | 2.3797 ms | $\pm$ 4.07% |
| **10-Slide Full Enterprise Deck** | `pptx` | **339.0** | 2.9501 ms | 6.2696 ms | $\pm$ 4.01% |
| **50-Slide Batch Scale** | `pptx` | **81.1** | 12.3338 ms | 107.9666 ms | $\pm$ 41.2% |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
