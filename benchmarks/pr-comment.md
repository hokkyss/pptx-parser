## 🚀 CI Benchmark & Bundle Size Telemetry

> Automated evaluation for commit hash: `local` | Target: `main`

### 📦 Production Bundle Footprint (Bundlephobia-Aligned)

| Package | Minified | Gzip Size | Brotli Size | DTS Size | Budget | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`@hokkyss/pptx-core`** | 1.59 KB | **437.00 B** | 382.00 B | 305.00 B | 5 KB | ✅ PASS |
| **`@hokkyss/pptx-reader`** | 95.41 KB | **25.83 KB** | 22.89 KB | 608.00 B | 30 KB | ✅ PASS |
| **`@hokkyss/pptx-writer`** | 125.94 KB | **31.69 KB** | 27.74 KB | 97.00 B | 35 KB | ✅ PASS |
| **`@hokkyss/pptx (Full SDK)`** | 190.99 KB | **48.93 KB** | 43.32 KB | 895.00 B | 70 KB | ✅ PASS |

### 🌲 First-Party Tree-Shaking Efficacy (Client Bundling)

| Import Scenario | Minified | Gzip Size | Brotli Size | Compression Ratio |
| :--- | :--- | :--- | :--- | :--- |
| **Core Units & Types Only** | 245.00 B | **127.00 B** | 95.00 B | 48.2% |
| **Reader Parser Only** | 30.45 KB | **8.55 KB** | 7.59 KB | 71.9% |
| **Writer Serializer Only** | 58.75 KB | **12.72 KB** | 11.16 KB | 78.4% |
| **Full Presentation Engine** | 27.67 KB | **6.72 KB** | 5.95 KB | 75.7% |

### ⚡ Microsecond Performance & Latency Matrix (Vitest)

| Benchmark Workload | Package | Frequency (ops/s) | Mean Latency | p99 Latency | Margin of Error |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Parse In-Memory PPTX Package to AST** | `pptx-reader` | **9,379.6** | 0.1066 ms | 0.166 ms | $\pm$ 0.45% |
| **Serialize Shape to DrawingML XML** | `pptx-writer` | **5,073,981.6** | 0.0002 ms | 0.0004 ms | $\pm$ 1.33% |
| **Serialize Table to DrawingML XML** | `pptx-writer` | **1,989,735.9** | 0.0005 ms | 0.0008 ms | $\pm$ 0.36% |
| **Serialize Chart to DrawingML XML** | `pptx-writer` | **836,008.1** | 0.0012 ms | 0.0023 ms | $\pm$ 3.28% |
| **Full writePptx ZIP Assembly** | `pptx-writer` | **1,469.3** | 0.6806 ms | 2.6893 ms | $\pm$ 4.17% |
| **Single Slide with Shape & Text** | `pptx` | **1,820.7** | 0.5492 ms | 0.7986 ms | $\pm$ 0.82% |
| **Enterprise Data Table (10 rows x 5 cols)** | `pptx` | **1,270.6** | 0.787 ms | 2.9784 ms | $\pm$ 6.1% |
| **DrawingML Column Chart Generation** | `pptx` | **1,243.4** | 0.8042 ms | 3.0161 ms | $\pm$ 8.93% |
| **10-Slide Full Enterprise Deck** | `pptx` | **357.8** | 2.795 ms | 4.8388 ms | $\pm$ 3.58% |
| **50-Slide Batch Scale** | `pptx` | **164.7** | 6.0717 ms | 17.3659 ms | $\pm$ 6.61% |

---
*Generated automatically by `pnpm bench:json` on Node.js v24.13.1 (darwin arm64)*
