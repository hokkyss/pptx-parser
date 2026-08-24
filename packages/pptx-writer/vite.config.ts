import { resolve } from 'node:path';
import { defineConfig, esmExternalRequirePlugin } from 'vite';
import dts from 'vite-plugin-dts';
import pkg from './package.json' with { type: 'json' };

export default defineConfig({
  build: {
    emptyOutDir: true,
    lib: {
      entry: resolve(import.meta.dirname, 'lib/index.ts'),
      fileName: () => 'index.js',
      formats: ['es'],
    },
    ssr: true,
  },
  resolve: {
    alias: {
      '@hokkyss/pptx-core': resolve(import.meta.dirname, '../pptx-core/lib/index.ts'),
      '@hokkyss/pptx-reader': resolve(import.meta.dirname, '../pptx-reader/lib/index.ts'),
    },
  },
  plugins: [
    esmExternalRequirePlugin({
      external: [
        ...Object.keys(pkg.dependencies || {}),
        ...Object.keys(pkg.peerDependencies || {}),
      ]
        .map<RegExp | string>((key) => new RegExp(`^${key}(/.+)*`))
        .concat('node:console'),
    }),
    dts({
      entryRoot: './lib',
      exclude: ['tests/**', 'scripts/**'],
    }),
  ],
});
