module.exports = {
  semi: false,
  singleQuote: true,
  printWidth: 80,
  trailingComma: 'all',
  jsxSingleQuote: false,
  tabWidth: 2,
  endOfLine: 'auto',
  plugins: [require('@trivago/prettier-plugin-sort-imports')],
  overrides: [
    {
      files: '*.json',
      options: {
        tabWidth: 2,
      },
    },
    {
      files: '*.js',
      options: {
        semi: false,
        singleQuote: true,
      },
    },
  ],
}
