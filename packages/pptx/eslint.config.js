import getConfig from '@monorepo/eslint';
import { defineConfig } from 'eslint/config';

export default defineConfig(
  getConfig({
    environment: 'isomorphic',
    outDir: 'dist',
    react: false,
    tanstackQuery: false,
    tanstackRouter: false,
    tsconfigPath: ['./tsconfig.json', './tsconfig.test.json'],
    tsconfigRootDir: import.meta.dirname,
  }),
  {
    rules: {
      '@typescript-eslint/no-base-to-string': 'off',
      'jsdoc/check-param-names': 'off',
      'jsdoc/escape-inline-tags': 'off',
      'jsdoc/require-param': 'off',
      'jsdoc/require-param-description': 'off',
      'jsdoc/require-returns': 'off',
      'perfectionist/sort-classes': 'off',
      'perfectionist/sort-exports': 'off',
      'perfectionist/sort-imports': 'off',
      'perfectionist/sort-interfaces': 'off',
      'perfectionist/sort-modules': 'off',
      'perfectionist/sort-named-imports': 'off',
      'perfectionist/sort-object-types': 'off',
      'perfectionist/sort-objects': 'off',
      'sort-keys': 'off',
    },
  },
);
