import getConfig from '@monorepo/eslint';
import { defineConfig } from 'eslint/config';

export default defineConfig(
  getConfig({
    environment: 'isomorphic',
    ignores: ['public/**'],
    outDir: 'dist',
    react: true,
    tanstackQuery: true,
    tanstackRouter: true,
    tsconfigRootDir: import.meta.dirname,
  }),
);
