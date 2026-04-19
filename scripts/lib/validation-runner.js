/**
 * Validation Runner — orchestrates config file validation.
 *
 * Loads schemas, reads config files (UTF-8 with Latin-1 fallback),
 * parses them, and runs structural/type/semantic validators.
 * Collects all diagnostics and determines the exit code.
 *
 * @module validation-runner
 */

import { readFileSync } from 'fs';
import { basename } from 'path';
import { createDiagnostic } from './data-models.js';
import { parseIni } from './ini-parser.js';
import { loadSchemas } from './schema-registry.js';
import { validateStructure } from './validators/structural.js';
import { validateTypes } from './validators/type.js';
import { validateSemantics } from './validators/semantic.js';

/**
 * Read a file as UTF-8 with Latin-1 fallback.
 * Tries UTF-8 first; if the result contains the Unicode replacement
 * character (U+FFFD), re-reads as Latin-1.
 *
 * @param {string} filePath
 * @returns {string} File content
 */
function readFileWithFallback(filePath) {
  const utf8Content = readFileSync(filePath, 'utf-8');
  if (utf8Content.includes('\uFFFD')) {
    return readFileSync(filePath, 'latin1');
  }
  return utf8Content;
}

/**
 * Run validation on config files.
 * @param {Object} options
 * @param {string} options.schemaDir - Path to schema JSON directory
 * @param {string[]} options.configFiles - Array of config file paths to validate
 * @returns {{ diagnostics: import('./data-models.js').Diagnostic[], exitCode: number }}
 */
export function runValidation(options) {
  const { schemaDir, configFiles } = options;
  const diagnostics = [];

  // 1. Load schemas
  const { schemas, diagnostics: schemaDiagnostics } = loadSchemas(schemaDir);
  diagnostics.push(...schemaDiagnostics);

  // 2. Process each config file
  for (const filePath of configFiles) {
    // 2a. Read the file
    let content;
    try {
      content = readFileWithFallback(filePath);
    } catch (err) {
      // 2b/2c. File not found or permission/IO error
      const message = err.code === 'ENOENT'
        ? `Config file not found: ${filePath}`
        : `Cannot read config file: ${err.message}`;
      diagnostics.push(
        createDiagnostic('error', filePath, 0, null, null, message)
      );
      continue;
    }

    // 2d. Parse the INI content
    const { document, diagnostics: parseDiagnostics } = parseIni(content, filePath);
    diagnostics.push(...parseDiagnostics);

    // 2e. Look up schema by uppercase filename
    const configName = basename(filePath).toUpperCase();
    const schema = schemas.get(configName);

    // 2f. If schema found, run all three validators
    if (schema) {
      diagnostics.push(...validateStructure(document, schema, filePath));
      diagnostics.push(...validateTypes(document, schema, filePath));
      diagnostics.push(...validateSemantics(document, schema, filePath));
    }
  }

  // 3. Determine exit code: non-zero if any diagnostic has severity 'error'
  const exitCode = diagnostics.some(d => d.severity === 'error') ? 1 : 0;

  // 4. Return results
  return { diagnostics, exitCode };
}
