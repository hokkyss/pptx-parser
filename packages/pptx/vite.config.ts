import { resolve } from 'node:path';
import { defineConfig, esmExternalRequirePlugin } from 'vite';
import dts from 'vite-plugin-dts';
import pkg from './package.json' with { type: 'json' };
import viteVersionPlugin from '@monorepo/vite-plugins/version';

export default defineConfig({
  build: {
    emptyOutDir: true,
    lib: {
      entry: {
        index: resolve(import.meta.dirname, 'lib/index.ts'),
      },
      formats: ['es'],
    },
    ssr: true,
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
    viteVersionPlugin(pkg.version),
  ],
});
