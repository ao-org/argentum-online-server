/**
 * Data models and factory functions for the config file validation tool.
 *
 * All models mirror the design document interfaces:
 * - IniDocument, IniSection, IniEntry (INI parsing)
 * - Diagnostic (validation output)
 * - ConfigSchema, SchemaEntry (schema definitions)
 */

/**
 * Create an IniDocument.
 * @param {import('./data-models.js').IniSection[]} sections - Ordered sections
 * @param {import('./data-models.js').IniEntry[]} orphans - Key-value pairs before any section header
 * @returns {{ sections: IniSection[], orphans: IniEntry[] }}
 */
export function createIniDocument(sections = [], orphans = []) {
  return { sections, orphans };
}

/**
 * Create an IniSection.
 * @param {string} name - Uppercase-normalized section name
 * @param {number} line - 1-based line number of the [SECTION] header
 * @param {import('./data-models.js').IniEntry[]} entries - Ordered key-value entries
 * @returns {{ name: string, line: number, entries: IniEntry[] }}
 */
export function createIniSection(name, line, entries = []) {
  return { name, line, entries };
}

/**
 * Create an IniEntry.
 * @param {string} key - Uppercase-normalized key
 * @param {string} value - Raw string value (right of first '=')
 * @param {number} line - 1-based line number
 * @returns {{ key: string, value: string, line: number }}
 */
export function createIniEntry(key, value, line) {
  return { key, value, line };
}

/**
 * Create a Diagnostic.
 * @param {'error'|'warning'|'info'} severity
 * @param {string} file - File path
 * @param {number} line - 1-based line number (0 if not applicable)
 * @param {string|null} section - Section name or null
 * @param {string|null} key - Key name or null
 * @param {string} message - Human-readable description
 * @returns {{ severity: string, file: string, line: number, section: string|null, key: string|null, message: string }}
 */
export function createDiagnostic(severity, file, line, section, key, message) {
  return { severity, file, line, section, key, message };
}

/**
 * Create a ConfigSchema.
 * @param {string} file - Config file name (e.g., "Server.ini")
 * @param {string} description - Human-readable description
 * @param {Object<string, { required: boolean, strict: boolean, keys: Object }>} sections - Section definitions
 * @param {Array} rules - Cross-field validation rules
 * @returns {{ file: string, description: string, sections: Object, rules: Array }}
 */
export function createConfigSchema(file, description = '', sections = {}, rules = []) {
  return { file, description, sections, rules };
}

/**
 * Create a SchemaEntry (internal, used by the schema generator).
 * @param {string} file - Config file this entry belongs to
 * @param {string} section - Uppercase section name
 * @param {string} key - Uppercase key name
 * @param {'string'|'integer'|'long'|'single'|'double'|'boolean'} type - Schema data type
 * @param {boolean} required - Whether the key is required
 * @param {string|null} defaultValue - Default value from VB6 source, or null
 * @param {string} source - Source file and line where this was extracted
 * @returns {{ file: string, section: string, key: string, type: string, required: boolean, defaultValue: string|null, source: string }}
 */
export function createSchemaEntry(file, section, key, type, required, defaultValue, source) {
  return { file, section, key, type, required, defaultValue, source };
}
