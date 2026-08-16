## 🚀 CI Benchmark & Bundle Size Telemetry

> Automated evaluation for commit hash: `local` | Target: `main`

### 📦 Production Bundle Footprint (Bundlephobia-Aligned)

| Package | Minified | Gzip Size | Brotli Size | DTS Size | Budget | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | 4.07 KB | **1.42 KB** | 1.25 KB | 353.00 B | 5 KB | ✅ PASS |
| **`@hokkyss/pptx-reader`** | 101.10 KB | **27.61 KB** | 24.43 KB | 691.00 B | 30 KB | ✅ PASS |
| **`@hokkyss/pptx-writer`** | 133.86 KB | **33.37 KB** | 29.29 KB | 175.00 B | 35 KB | ✅ PASS |
| **`@hokkyss/pptx (Full SDK)`** | 214.82 KB | **55.47 KB** | 49.08 KB | 1.07 KB | 70 KB | ✅ PASS |

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
| **Reader Parser Only** | 38.89 KB | **11.04 KB** | 9.83 KB | 71.6% |
| **Writer Serializer Only** | 69.39 KB | **15.35 KB** | 13.42 KB | 77.9% |
| **Full Presentation Engine** | 30.26 KB | **7.37 KB** | 6.52 KB | 75.7% |

### ⚡ Microsecond Performance & Latency Matrix (Vitest)

| Benchmark Workload | Package | Frequency (ops/s) | Mean Latency | p99 Latency | Margin of Error |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **7,896.9** | 0.1266 ms | 0.175 ms | $\pm$ 0.24% |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **4,142,451.6** | 0.0002 ms | 0.0006 ms | $\pm$ 2.26% |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,636,534.4** | 0.0006 ms | 0.0009 ms | $\pm$ 1.09% |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **897,506.3** | 0.0011 ms | 0.0014 ms | $\pm$ 0.21% |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,350.3** | 0.7406 ms | 2.7403 ms | $\pm$ 4.02% |
| **Single Slide with Shape & Text** | `pptx` | **1,418.4** | 0.705 ms | 2.4461 ms | $\pm$ 3.84% |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **961.4** | 1.0401 ms | 5.1173 ms | $\pm$ 10.27% |
| **DrawingML Column Chart Generation** | `pptx` | **1,183.6** | 0.8449 ms | 3.4944 ms | $\pm$ 5.24% |
| **10-Slide Full Enterprise Deck** | `pptx` | **324.7** | 3.0803 ms | 9.9807 ms | $\pm$ 6.14% |
| **50-Slide Batch Scale** | `pptx` | **141.6** | 7.0628 ms | 76.2813 ms | $\pm$ 28.48% |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
