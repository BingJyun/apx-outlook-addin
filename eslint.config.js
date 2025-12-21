// eslint.config.js
export default [
  {
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        Office: "readonly",
        window: "readonly",
        document: "readonly",
      },
    },
    rules: {
      /* === 基本安全 === */
      "no-eval": "error",
      "no-implied-eval": "error",
      "no-new-func": "error",
      "no-console": ["warn", { allow: ["error"] }],

      /* === 乾淨程式碼 === */
      "no-var": "error",
      "prefer-const": "error",
      "no-unused-vars": ["error", { argsIgnorePattern: "^_" }],
      "no-undef": "error",
      "no-shadow": "error",

      /* === 非同步規範 === */
      "no-promise-executor-return": "error",
      "require-await": "error",

      /* === 可讀性 === */
      "eqeqeq": ["error", "always"],
      "curly": ["error", "all"],
      "consistent-return": "error",

      /* === Magic 防止 === */
      "no-magic-numbers": [
        "error",
        {
          ignore: [0, 1],
          ignoreArrayIndexes: true,
          enforceConst: true,
          detectObjects: false,
        },
      ],

      /* === 模組化 === */
      "no-duplicate-imports": "error",

      /* === Outlook / Office.js 特別限制 === */
      "no-restricted-globals": ["error", "localStorage", "sessionStorage"],
    },
  },
];