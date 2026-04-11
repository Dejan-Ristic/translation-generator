import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import { defineConfig } from 'eslint/config';
import globals from 'globals';
import prettier from 'eslint-config-prettier';

export default defineConfig([
  // Base JS rules
  js.configs.recommended,

  // TypeScript recommended rules (includes parser + plugin)
  ...tseslint.configs.recommended,

  // Your project config
  {
    files: ['**/*.{js,mjs,cjs,ts,mts,cts}'],

    languageOptions: {
      globals: {
        ...globals.node,
        ...globals.browser,
      },
    },

    rules: {
      // your overrides here
    },
  },

  // disables ESLint rules that conflict with Prettier
  prettier,
]);
