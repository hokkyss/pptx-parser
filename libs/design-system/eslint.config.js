import getConfig from '@monorepo/eslint';
import { defineConfig } from 'eslint/config';

export default defineConfig(
  {
    ignores: ['lib/components/ui/**', 'dist/**'],
  },
  getConfig({
    environment: 'browser',
    outDir: 'dist',
    react: true,
    tanstackQuery: false,
    tanstackRouter: false,
    tsconfigRootDir: import.meta.dirname,
  }),
);
