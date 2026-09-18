const js = require("@eslint/js");
const globals = require("globals");
module.exports = [
  { ignores: ["dist/**", "node_modules/**", ".serena/**"] },
  js.configs.recommended,
  {
    files: ["**/*.js", "**/*.cjs"],
    languageOptions: {
      sourceType: "commonjs",
      globals: {
        ...globals.node,
        ...globals.browser,
        Office: "readonly",
        __ANALYTICS_ORIGIN__: "readonly",
      },
    },
    rules: {
      "no-unused-vars": ["error", { argsIgnorePattern: "^_", caughtErrorsIgnorePattern: "^_" }],
      "no-console": "error",
    },
  },
];
