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

### ⚡ Microsecond Performance & Latency Delta (vs `main`)

| Benchmark Workload | Package | PR Throughput | Base Throughput | Δ Throughput | PR Latency | Base Latency | Δ Latency | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **6,627.0** | 6,627.0 | **-0.0%** | 0.1509 ms | 0.1509 ms | 0.0% | 🟢 PARITY |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **4,570,456.4** | 4,570,456.0 | **+0.0%** | 0.0002 ms | 0.0002 ms | 0.0% | 🟢 PARITY |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,515,909.1** | 1,515,909.0 | **+0.0%** | 0.0007 ms | 0.0007 ms | 0.0% | 🟢 PARITY |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **840,416.9** | 840,417.0 | **-0.0%** | 0.0012 ms | 0.0012 ms | 0.0% | 🟢 PARITY |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,393.8** | 1,394.0 | **-0.0%** | 0.7175 ms | 0.7175 ms | 0.0% | 🟢 PARITY |
| **Single Slide with Shape & Text** | `pptx` | **1,566.4** | 1,566.0 | **+0.0%** | 0.6384 ms | 0.6384 ms | 0.0% | 🟢 PARITY |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **1,583.7** | 1,584.0 | **-0.0%** | 0.6314 ms | 0.6314 ms | 0.0% | 🟢 PARITY |
| **DrawingML Column Chart Generation** | `pptx` | **1,378.1** | 1,378.0 | **+0.0%** | 0.7256 ms | 0.7256 ms | 0.0% | 🟢 PARITY |
| **10-Slide Full Enterprise Deck** | `pptx` | **309.3** | 309.0 | **+0.1%** | 3.2329 ms | 3.2329 ms | 0.0% | 🟢 PARITY |
| **50-Slide Batch Scale** | `pptx` | **136.3** | 136.0 | **+0.2%** | 7.3353 ms | 7.3353 ms | 0.0% | 🟢 PARITY |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
