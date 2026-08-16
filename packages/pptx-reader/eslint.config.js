import getConfig from '@monorepo/eslint';
import { defineConfig } from 'eslint/config';

export default defineConfig({
  extends: [
    getConfig({
      environment: 'isomorphic',
      outDir: 'dist',
      react: false,
      tanstackQuery: false,
      tanstackRouter: false,
      tsconfigPath: ['./tsconfig.json', './tsconfig.test.json'],
      tsconfigRootDir: import.meta.dirname,
    }),
  ],
  ignores: ['dist/**', 'types/vite-env.d.ts', 'scripts/**', 'parsed_output.json', '**/*.json'],
  rules: {
    'perfectionist/sort-objects': 'off',
    'perfectionist/sort-interfaces': 'off',
    'perfectionist/sort-object-types': 'off',
    'perfectionist/sort-classes': 'off',
    'perfectionist/sort-imports': 'off',
    'perfectionist/sort-named-imports': 'off',
    'perfectionist/sort-exports': 'off',
    'perfectionist/sort-modules': 'off',
    'sort-keys': 'off',
    'jsdoc/escape-inline-tags': 'off',
    'jsdoc/require-returns': 'off',
    'jsdoc/require-param': 'off',
    'jsdoc/check-param-names': 'off',
    '@typescript-eslint/no-base-to-string': 'off',
  },
});
