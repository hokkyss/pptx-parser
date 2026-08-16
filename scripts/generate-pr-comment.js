import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

function formatNumber(num, decimals = 2) {
  return Number(num).toLocaleString('en-US', {
    maximumFractionDigits: decimals,
    minimumFractionDigits: decimals,
  });
}

function generateMarkdown(report) {
  let md = '';

  md += `## 🚀 CI Benchmark & Bundle Size Telemetry\n\n`;
  md += `> Automated evaluation for commit hash: \`${process.env.GITHUB_SHA ? process.env.GITHUB_SHA.slice(0, 7) : 'local'}\` | Target: \`${process.env.GITHUB_REF_NAME || 'main'}\`\n\n`;

  // 1. Bundle Size Table
  md += `### 📦 Production Bundle Footprint (Bundlephobia-Aligned)\n\n`;
  md += `| Package | Minified | Gzip Size | Brotli Size | DTS Size | Budget | Status |\n`;
  md += `| :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n`;

  for (const [pkgName, data] of Object.entries(report.bundleSize?.packages || {})) {
    const status = data.passed ? '✅ PASS' : '⚠️ OVER';
    md += `| **\`${pkgName}\`** | ${data.minifiedFormatted} | **${data.gzipFormatted}** | ${data.brotliFormatted} | ${data.dtsFormatted} | ${data.targetBudgetKb} KB | ${status} |\n`;
  }

  // 1b. Baseline Comparison (Δ Delta)
  const baselinePath = resolve(process.cwd(), 'benchmarks/bundle-baseline.json');
  if (existsSync(baselinePath)) {
    try {
      const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
      const formatDelta = (delta, baseVal) => {
        if (!baseVal || delta === 0) return '0 B (0.0%)';
        const sign = delta > 0 ? '+' : '-';
        const pct = ((delta / baseVal) * 100).toFixed(1);
        const abs = Math.abs(delta);
        const formattedBytes = abs < 1024 ? `${abs} B` : `${(abs / 1024).toFixed(2)} KB`;
        return `${sign}${formattedBytes} (${delta > 0 ? '+' : ''}${pct}%)`;
      };

      md += `\n### 📊 Baseline Comparison (Δ Delta)\n\n`;
      md += `| Package | Δ Minified | Δ Gzip | Δ Brotli | Status |\n`;
      md += `| :--- | :--- | :--- | :--- | :--- |\n`;

      for (const [pkgName, curr] of Object.entries(report.bundleSize?.packages || {})) {
        const base = baseline[pkgName];
        if (!base) continue;
        const baseMin = base.minifiedBytes ?? base.rawBytes ?? base.rawSize ?? 0;
        const baseGzip = base.gzipBytes ?? base.gzipSize ?? 0;
        const baseBrotli = base.brotliBytes ?? base.brotliSize ?? 0;

        const deltaMin = curr.minifiedBytes - baseMin;
        const deltaGzip = curr.gzipBytes - baseGzip;
        const deltaBrotli = curr.brotliBytes - baseBrotli;

        const status = deltaGzip > 0 ? '🔴 INCREASED' : deltaGzip < 0 ? '🟢 REDUCED' : '🟢 OPTIMIZED';
        md += `| **\`${pkgName}\`** | ${formatDelta(deltaMin, baseMin)} | ${formatDelta(deltaGzip, baseGzip)} | ${formatDelta(deltaBrotli, baseBrotli)} | ${status} |\n`;
      }
    } catch {
      // ignore baseline parse failure
    }
  }

  // 2. Tree-Shaking Table
  if (report.bundleSize?.treeShaking?.length) {
    md += `\n### 🌲 First-Party Tree-Shaking Efficacy (Client Bundling)\n\n`;
    md += `| Import Scenario | Minified | Gzip Size | Brotli Size | Compression Ratio |\n`;
    md += `| :--- | :--- | :--- | :--- | :--- |\n`;
    for (const item of report.bundleSize.treeShaking) {
      md += `| **${item.scenario}** | ${item.minified || item.raw} | **${item.gzip}** | ${item.brotli} | ${item.ratio} |\n`;
    }
  }

  // 3. Vitest Performance Table
  if (report.performance?.summary?.length) {
    md += `\n### ⚡ Microsecond Performance & Latency Matrix (Vitest)\n\n`;
    md += `| Benchmark Workload | Package | Frequency (ops/s) | Mean Latency | p99 Latency | Margin of Error |\n`;
    md += `| :--- | :--- | :--- | :--- | :--- | :--- |\n`;

    for (const b of report.performance.summary) {
      md += `| **${b.name}** | \`${b.package.replace('@hokkyss/', '')}\` | **${formatNumber(b.opsPerSec, 1)}** | ${b.meanMs} ms | ${b.p99Ms} ms | $\\pm$ ${b.rmePercent}% |\n`;
    }
  }

  md += `\n---\n`;
  md += `*Generated automatically by \`pnpm bench:json\` on Node.js ${report.system?.node || process.version} (${report.system?.platform || process.platform} ${report.system?.arch || process.arch})*\n`;

  return md;
}

function main() {
  const reportPath = resolve(process.cwd(), 'benchmarks/benchmark-report.json');
  if (!existsSync(reportPath)) {
    console.error(`❌ Benchmark report not found at ${reportPath}. Run "pnpm bench:json" first.`);
    process.exit(1);
  }

  const report = JSON.parse(readFileSync(reportPath, 'utf8'));
  const md = generateMarkdown(report);

  const outputPath = resolve(process.cwd(), 'benchmarks/pr-comment.md');
  writeFileSync(outputPath, md);
  console.log(`✅ Generated PR Markdown Comment at: ${outputPath}`);
  console.log('\n' + md);
}

main();
