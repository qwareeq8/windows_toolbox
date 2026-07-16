import js from "@eslint/js";
import globals from "globals";
import reactHooks from "eslint-plugin-react-hooks";

export default [
  {
    ignores: ["coverage/**", "dist/**", "node_modules/**"],
  },
  js.configs.recommended,
  {
    files: ["**/*.{js,jsx}"],
    languageOptions: {
      ecmaVersion: "latest",
      globals: {
        ...globals.browser,
        __APP_VERSION__: "readonly",
        QWebChannel: "readonly",
      },
      parserOptions: {
        ecmaFeatures: { jsx: true },
        sourceType: "module",
      },
    },
    plugins: {
      "react-hooks": reactHooks,
    },
    rules: {
      "react-hooks/exhaustive-deps": "warn",
      "react-hooks/rules-of-hooks": "error",
    },
  },
  {
    files: ["src/__tests__/**/*.{js,jsx}", "src/test-setup.js"],
    languageOptions: {
      globals: globals.vitest,
    },
  },
  {
    files: ["eslint.config.js", "vite.config.js"],
    languageOptions: {
      globals: globals.node,
    },
  },
];
