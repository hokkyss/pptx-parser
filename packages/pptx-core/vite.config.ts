import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import dts from 'vite-plugin-dts';

export default defineConfig({
  build: {
    lib: {
      entry: resolve(import.meta.dirname, 'lib/index.ts'),
      fileName: () => 'index.js',
      formats: ['es'],
    },
    outDir: 'dist',
    sourcemap: true,
  },
  plugins: [
    dts({
      entryRoot: './lib',
    }),
  ],
});
