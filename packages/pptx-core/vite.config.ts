import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import dts from 'vite-plugin-dts';
import pkg from './package.json';
import viteVersionPlugin from '@monorepo/vite-plugins/version';

export default defineConfig({
  build: {
    lib: {
      entry: {
        index: resolve(import.meta.dirname, 'lib/index.ts'),
      },
      formats: ['es'],
    },
    outDir: 'dist',
    sourcemap: true,
  },
  plugins: [
    dts({
      entryRoot: './lib',
    }),
    viteVersionPlugin(pkg.version),
  ],
});
