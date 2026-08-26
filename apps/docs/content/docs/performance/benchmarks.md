---
title: "Runtime Benchmarks"
description: "Sub-millisecond latency profile evaluated across Apple Silicon and V8 engines."
order: 2
section: "performance"
---

# Runtime Benchmarks

Benchmarked on Apple Silicon (M-series / Node.js v22):

| Benchmark Workload | Target Engine | Operations / Sec | Mean Execution Latency | p99 Latency |
| :--- | :--- | :--- | :--- | :--- |
| **Parse Binary PPTX to AST** | `@hokkyss/pptx-reader` | **9,380 ops/s** | **0.106 ms** | 0.166 ms |
| **Shape -> DrawingML XML** | `@hokkyss/pptx-writer` | **5,070,000 ops/s** | **0.0002 ms** | 0.0004 ms |
| **Table -> DrawingML XML** | `@hokkyss/pptx-writer` | **1,990,000 ops/s** | **0.0005 ms** | 0.0008 ms |
| **Chart -> DrawingML XML** | `@hokkyss/pptx-writer` | **836,000 ops/s** | **0.0012 ms** | 0.0023 ms |
| **Single Slide Full Assembly** | `@hokkyss/pptx` | **1,820 ops/s** | **0.549 ms** | 0.798 ms |
| **10-Slide Full Presentation** | `@hokkyss/pptx` | **358 ops/s** | **2.795 ms** | 4.838 ms |
| **50-Slide Batch Generation** | `@hokkyss/pptx` | **165 ops/s** | **6.071 ms** | 17.36 ms |
