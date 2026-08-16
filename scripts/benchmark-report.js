import { execSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, unlinkSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { getBundleSizeData } from './benchmark-bundle-size.js';

async function generateBenchmarkReport() {
  console.log(`\n========================================================================================`);
  console.log(` 📊 GENERATING UNIFIED BENCHMARK JSON REPORT`);
  console.log(`========================================================================================\n`);

  const benchmarkDir = resolve(process.cwd(), 'benchmarks');
  if (!existsSync(benchmarkDir)) mkdirSync(benchmarkDir, { recursive: true });

  const tempJson = resolve(benchmarkDir, `tmp_vitest_${Date.now()}.json`);

  // 1. Collect Bundle Size & Tree-Shaking Data
  console.log('📦 Analyzing Monorepo Bundle Sizes & Tree-Shaking Metrics...');
  const bundleData = await getBundleSizeData();

  // 2. Collect Vitest Performance Benchmarks
  console.log('⚡ Running Vitest Benchmarks Across Packages...');
  const vitestResults = {};

  const packages = [
    { name: '@hokkyss/pptx-reader', pkgDir: 'packages/pptx-reader' },
    { name: '@hokkyss/pptx-writer', pkgDir: 'packages/pptx-writer' },
    { name: '@hokkyss/pptx', pkgDir: 'packages/pptx' },
  ];

  for (const pkg of packages) {
    try {
      const outputJsonPath = resolve(benchmarkDir, `tmp_${pkg.name.replace(/[^a-zA-Z0-9]/g, '_')}.json`);
      execSync(`pnpm --filter ${pkg.name} exec vitest bench --outputJson ${outputJsonPath}`, {
        stdio: 'pipe',
        cwd: process.cwd(),
      });

      if (existsSync(outputJsonPath)) {
        const rawJson = JSON.parse(readFileSync(outputJsonPath, 'utf8'));
        vitestResults[pkg.name] = rawJson;
        unlinkSync(outputJsonPath);
      }
    } catch (err) {
      console.warn(`⚠️ Warning: Failed to run benchmarks for ${pkg.name}:`, err.message);
    }
  }

  // 3. Extract Clean Benchmark Summary
  const parsedPerfBenchmarks = [];
  for (const [pkgName, data] of Object.entries(vitestResults)) {
    for (const file of data.files || []) {
      for (const group of file.groups || []) {
        for (const bench of group.benchmarks || []) {
          parsedPerfBenchmarks.push({
            package: pkgName,
            suite: group.fullName,
            name: bench.name,
            opsPerSec: Number(bench.hz.toFixed(2)),
            meanMs: Number(bench.mean.toFixed(4)),
            minMs: Number(bench.min.toFixed(4)),
            maxMs: Number(bench.max.toFixed(4)),
            p75Ms: Number(bench.p75.toFixed(4)),
            p99Ms: Number(bench.p99.toFixed(4)),
            rmePercent: Number(bench.rme.toFixed(2)),
            samples: bench.sampleCount || bench.samples?.length || 0,
          });
        }
      }
    }
  }

  // 4. Construct Unified JSON Report
  const finalReport = {
    schemaVersion: '1.0.0',
    generatedAt: new Date().toISOString(),
    system: {
      node: process.version,
      platform: process.platform,
      arch: process.arch,
    },
    bundleSize: {
      packages: bundleData.packages,
      treeShaking: bundleData.treeShaking,
    },
    performance: {
      summary: parsedPerfBenchmarks,
      rawVitest: vitestResults,
    },
  };

  const args = process.argv.slice(2);
  const outputIdx = args.indexOf('--output');
  const targetFile = outputIdx !== -1 && args[outputIdx + 1]
    ? resolve(process.cwd(), args[outputIdx + 1])
    : resolve(benchmarkDir, 'benchmark-report.json');

  writeFileSync(targetFile, JSON.stringify(finalReport, null, 2));

  console.log(`\n✅ Unified Benchmark JSON Report successfully generated!`);
  console.log(`💾 File: ${targetFile}`);
  console.log(`📊 Benchmarks recorded: ${parsedPerfBenchmarks.length}`);
  console.log(`📦 Monorepo packages analyzed: ${Object.keys(bundleData.packages).length}\n`);
}

generateBenchmarkReport().catch(console.error);
