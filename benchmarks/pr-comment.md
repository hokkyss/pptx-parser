## 🚀 CI Benchmark & Bundle Size Telemetry

> Automated evaluation for commit hash: `local` | Target: `main`

### 📦 Production Bundle Footprint (Bundlephobia-Aligned)

| Package | Minified | Gzip Size | Brotli Size | DTS Size | Budget | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | 4.07 KB | **1.42 KB** | 1.25 KB | 353.00 B | 5 KB | ✅ PASS |
| **`@hokkyss/pptx-reader`** | 101.10 KB | **27.61 KB** | 24.43 KB | 691.00 B | 30 KB | ✅ PASS |
| **`@hokkyss/pptx-writer`** | 133.86 KB | **33.37 KB** | 29.29 KB | 175.00 B | 35 KB | ✅ PASS |
| **`@hokkyss/pptx (Full SDK)`** | 212.33 KB | **54.76 KB** | 48.41 KB | 1.07 KB | 70 KB | ✅ PASS |

### 📊 Baseline Comparison (Δ Delta)

| Package | Δ Minified | Δ Gzip | Δ Brotli | Status |
| :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |
| **`@hokkyss/pptx-reader`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |
| **`@hokkyss/pptx-writer`** | 0 B (0.0%) | 0 B (0.0%) | 0 B (0.0%) | 🟢 OPTIMIZED |
| **`@hokkyss/pptx (Full SDK)`** | -2.49 KB (-1.2%) | -729 B (-1.3%) | -685 B (-1.4%) | 🟢 REDUCED |

### 🌲 First-Party Tree-Shaking Efficacy (Client Bundling)

| Import Scenario | Minified | Gzip Size | Brotli Size | Compression Ratio |
| :--- | :--- | :--- | :--- | :--- |
| **Core Units & Types Only** | 502.00 B | **262.00 B** | 231.00 B | 47.8% |
| **Reader Parser Only** | 38.89 KB | **11.04 KB** | 9.83 KB | 71.6% |
| **Writer Serializer Only** | 69.39 KB | **15.35 KB** | 13.42 KB | 77.9% |
| **Full Presentation Engine** | 29.85 KB | **7.27 KB** | 6.44 KB | 75.6% |

### ⚡ Microsecond Performance & Latency Matrix (Vitest)

| Benchmark Workload | Package | Frequency (ops/s) | Mean Latency | p99 Latency | Margin of Error |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **7,702.6** | 0.1298 ms | 0.2132 ms | $\pm$ 0.56% |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **4,476,096.0** | 0.0002 ms | 0.0005 ms | $\pm$ 1.83% |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,409,103.7** | 0.0007 ms | 0.0016 ms | $\pm$ 1.6% |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **768,506.8** | 0.0013 ms | 0.0026 ms | $\pm$ 8.12% |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,282.9** | 0.7795 ms | 3.5574 ms | $\pm$ 5.05% |
| **Single Slide with Shape & Text** | `pptx` | **976.7** | 1.0239 ms | 4.5016 ms | $\pm$ 6.21% |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **1,208.5** | 0.8275 ms | 3.6478 ms | $\pm$ 10.37% |
| **DrawingML Column Chart Generation** | `pptx` | **658.1** | 1.5195 ms | 22.7373 ms | $\pm$ 26.28% |
| **10-Slide Full Enterprise Deck** | `pptx` | **240.0** | 4.1673 ms | 14.0686 ms | $\pm$ 10.54% |
| **50-Slide Batch Scale** | `pptx` | **92.9** | 10.7678 ms | 33.9334 ms | $\pm$ 16.31% |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
