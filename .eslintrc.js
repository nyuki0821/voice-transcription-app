module.exports = {
  env: {
    browser: true,
    es2021: true,
    googleappsscript: true
  },
  extends: [
    'google'
  ],
  parserOptions: {
    ecmaVersion: 12,
    sourceType: 'script'
  },
  globals: {
    // Google Apps Script globals
    Logger: 'readonly',
    DriveApp: 'readonly',
    GmailApp: 'readonly',
    SpreadsheetApp: 'readonly',
    UrlFetchApp: 'readonly',
    Utilities: 'readonly',
    PropertiesService: 'readonly',
    ScriptApp: 'readonly'
  },
  rules: {
    'max-len': ['error', { code: 120 }],
    'require-jsdoc': 'warn',
    'valid-jsdoc': 'warn'
  }
}; 