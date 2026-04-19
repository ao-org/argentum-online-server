/**
 * Schema Registry — loads and indexes JSON schema files from a directory.
 *
 * Each `.json` file in the schema directory is parsed as a ConfigSchema
 * and indexed by the uppercase config file name derived from the filename.
 * For example, `Server.ini.json` ? key `"SERVER.INI"`.
 *
 * @module schema-registry
 */

import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';
import { createDiagnostic } from './data-models.js';

/**
 * Derive the uppercase config file name from a schema filename.
 * E.g., "Server.ini.json" ? "SERVER.INI"
 * @param {string} filename - Schema file name (e.g., "Server.ini.json")
 * @returns {string} Uppercase config file name
 */
function configNameFromFilename(filename) {
  // Strip the trailing ".json" to get the config file name
  return filename.replace(/\.json$/i, '').toUpperCase();
}

/**
 * Load all schema JSON files from a directory.
 * @param {string} schemaDir - Path to schema directory
 * @returns {{ schemas: Map<string, ConfigSchema>, diagnostics: Diagnostic[] }}
 */
export function loadSchemas(schemaDir) {
  /** @type {Map<string, import('./data-models.js').ConfigSchema>} */
  const schemas = new Map();
  /** @type {import('./data-models.js').Diagnostic[]} */
  const diagnostics = [];

  // Attempt to read the directory
  let files;
  try {
    files = readdirSync(schemaDir);
  } catch (err) {
    diagnostics.push(
      createDiagnostic('error', schemaDir, 0, null, null,
        `Cannot read schema directory: ${err.message}`)
    );
    return { schemas, diagnostics };
  }

  // Filter to .json files only
  const jsonFiles = files.filter(f => f.endsWith('.json'));

  for (const filename of jsonFiles) {
    const filePath = join(schemaDir, filename);
    let content;
    try {
      content = readFileSync(filePath, 'utf-8');
    } catch (err) {
      diagnostics.push(
        createDiagnostic('error', filePath, 0, null, null,
          `Cannot read schema file: ${err.message}`)
      );
      continue;
    }

    let schema;
    try {
      schema = JSON.parse(content);
    } catch (err) {
      diagnostics.push(
        createDiagnostic('error', filePath, 0, null, null,
          `Malformed JSON in schema file: ${err.message}`)
      );
      continue;
    }

    // Prefer the schema's `file` property for the key; fall back to filename
    const key = (schema.file)
      ? schema.file.toUpperCase()
      : configNameFromFilename(filename);
    schemas.set(key, schema);
  }

  return { schemas, diagnostics };
}
