// eslint.config.mjs
import tsparser from "@typescript-eslint/parser";
import obsidianmd from "eslint-plugin-obsidianmd";

export default [
  {
    ignores: [
      "**/*.js",
      "**/*.mjs",
      "node_modules/**",
      "main.js",
      "*.config.*",
    ],
  },
  {
    files: ["src/**/*.ts"],
    languageOptions: {
      parser: tsparser,
      parserOptions: {
        project: "./tsconfig.json",
        tsconfigRootDir: import.meta.dirname,
      },
      globals: {
        // Node.js globals
        require: "readonly",
        process: "readonly",
        console: "readonly",
        // Browser globals
        document: "readonly",
        window: "readonly",
      },
    },
    plugins: {
      obsidianmd,
    },
    rules: {
      // Use only the Obsidian plugin rules
      ...obsidianmd.configs.recommended.rules,

      // Disable TypeScript ESLint rules that are too strict for Obsidian plugins
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-floating-promises": "off",
      "@typescript-eslint/no-require-imports": "off",
      "@typescript-eslint/await-thenable": "off",
      "@typescript-eslint/no-unused-vars": "off",
      "no-undef": "off",
    },
  },
];
