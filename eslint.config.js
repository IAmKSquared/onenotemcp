import js from '@eslint/js';
import prettierConfig from 'eslint-config-prettier';
import jsdoc from 'eslint-plugin-jsdoc';

export default [
  // Apply recommended rules
  js.configs.recommended,

  // JSDoc plugin configuration
  {
    plugins: {
      jsdoc,
    },
    rules: {
      ...jsdoc.configs.recommended.rules,
      'jsdoc/require-description': 'warn',
      'jsdoc/require-param-description': 'warn',
      'jsdoc/require-returns-description': 'warn',
      'jsdoc/check-alignment': 'error',
      'jsdoc/check-param-names': 'error',
      'jsdoc/check-tag-names': 'error',
      'jsdoc/check-types': 'warn',
      'jsdoc/require-jsdoc': [
        'warn',
        {
          require: {
            FunctionDeclaration: true,
            MethodDefinition: true,
            ClassDeclaration: true,
            ArrowFunctionExpression: false,
            FunctionExpression: true,
          },
        },
      ],
    },
  },

  // General JavaScript rules
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        console: 'readonly',
        process: 'readonly',
        Buffer: 'readonly',
        __dirname: 'readonly',
        __filename: 'readonly',
        setTimeout: 'readonly',
        setInterval: 'readonly',
        clearTimeout: 'readonly',
        clearInterval: 'readonly',
        global: 'readonly',
      },
    },
    rules: {
      // Code quality
      'no-unused-vars': [
        'warn',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
      'no-console': 'off', // We use console for server output
      'no-debugger': 'warn',
      'no-unreachable': 'error',
      'no-constant-condition': ['error', { checkLoops: false }],

      // Best practices
      eqeqeq: ['error', 'always'],
      curly: ['error', 'all'],
      'no-eval': 'error',
      'no-implied-eval': 'error',
      'no-multi-spaces': 'warn',
      'no-throw-literal': 'error',
      'prefer-const': 'warn',
      'no-var': 'error',

      // Security
      'no-new-func': 'error',
    },
  },

  // Disable Prettier conflicting rules (must be last)
  prettierConfig,

  // Ignore patterns
  {
    ignores: ['node_modules/**', '.git/**', 'coverage/**', 'dist/**', '*.min.js'],
  },
];
