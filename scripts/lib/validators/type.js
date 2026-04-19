/**
 * Type Validator — checks that INI values match their declared schema types
 * and respect numeric bounds.
 *
 * Supported types:
 * - integer (CInt): valid integer string in [-32768, 32767]
 * - long (CLng): valid integer string in [-2147483648, 2147483647]
 * - single / double: valid finite floating-point number
 * - boolean: exactly "0" or "1"
 * - string: always valid
 *
 * @module validators/type
 */

import { createDiagnostic } from '../data-models.js';

/**
 * Mimic VB6's val() function: extract the leading numeric portion of a string,
 * ignoring any trailing non-numeric text (e.g., inline comments like "180 'minutes").
 * Returns the trimmed numeric prefix, or the original trimmed value if no numeric prefix found.
 *
 * @param {string} value - Raw INI value
 * @returns {string} Numeric prefix or original value
 */
export function vb6Val(value) {
  const trimmed = value.trim();
  // Match optional sign, digits, optional decimal part
  const match = trimmed.match(/^([+-]?\d+(?:\.\d+)?)/);
  if (match) {
    return match[1];
  }
  return trimmed;
}

/** @type {Record<string, { check: (v: string) => boolean, range?: [number, number], label: string }>} */
const TYPE_RULES = {
  integer: {
    label: 'integer',
    range: [-32768, 32767],
    check(v) {
      if (!/^-?\d+$/.test(v)) return false;
      const n = Number(v);
      return n >= -32768 && n <= 32767;
    },
  },
  long: {
    label: 'long',
    range: [-2147483648, 2147483647],
    check(v) {
      if (!/^-?\d+$/.test(v)) return false;
      const n = Number(v);
      return n >= -2147483648 && n <= 2147483647;
    },
  },
  single: {
    label: 'single',
    check(v) {
      if (v.trim() === '') return false;
      const n = parseFloat(v);
      return !isNaN(n) && isFinite(n);
    },
  },
  double: {
    label: 'double',
    check(v) {
      if (v.trim() === '') return false;
      const n = parseFloat(v);
      return !isNaN(n) && isFinite(n);
    },
  },
  boolean: {
    label: 'boolean',
    check(v) {
      return v === '0' || v === '1';
    },
  },
  string: {
    label: 'string',
    check() {
      return true;
    },
  },
};

/**
 * Validate values in an IniDocument against the types and bounds declared in a ConfigSchema.
 *
 * Only keys present in both the document AND the schema are validated.
 *
 * @param {import('../data-models.js').IniDocument} document
 * @param {import('../data-models.js').ConfigSchema} schema
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
export function validateTypes(document, schema, filePath) {
  const diagnostics = [];
  const schemaSections = schema.sections || {};

  // Index document sections by uppercase name (already normalized by parser)
  /** @type {Map<string, import('../data-models.js').IniSection[]>} */
  const docSectionsByName = new Map();
  for (const section of document.sections) {
    if (!docSectionsByName.has(section.name)) {
      docSectionsByName.set(section.name, [section]);
    } else {
      docSectionsByName.get(section.name).push(section);
    }
  }

  for (const [sectionName, sectionDef] of Object.entries(schemaSections)) {
    const upperSection = sectionName.toUpperCase();
    const occurrences = docSectionsByName.get(upperSection);
    if (!occurrences) continue;

    const schemaKeys = sectionDef.keys || {};

    for (const section of occurrences) {
      for (const entry of section.entries) {
        const keyDef = schemaKeys[entry.key] || schemaKeys[entry.key.toUpperCase()];
        if (!keyDef) continue;

        const typeName = (keyDef.type || 'string').toLowerCase();
        const rule = TYPE_RULES[typeName];
        if (!rule) continue;

        // Strip trailing text to mimic VB6 val() behavior for numeric types
        const effectiveValue = (typeName !== 'string' && typeName !== 'boolean')
          ? vb6Val(entry.value)
          : entry.value.trim();

        // Type check
        if (!rule.check(effectiveValue)) {
          const rangeHint = rule.range
            ? ` (allowed range: ${rule.range[0]} to ${rule.range[1]})`
            : '';
          diagnostics.push(
            createDiagnostic(
              'error',
              filePath,
              entry.line,
              upperSection,
              entry.key,
              `Expected type '${rule.label}', got '${effectiveValue}'${rangeHint}`
            )
          );
          // Skip bounds check if type itself is invalid
          continue;
        }

        // Bounds check (only for numeric types with min/max in schema)
        if (typeName === 'string' || typeName === 'boolean') continue;

        const numericValue = parseFloat(effectiveValue);
        const hasMin = keyDef.min !== undefined && keyDef.min !== null;
        const hasMax = keyDef.max !== undefined && keyDef.max !== null;

        if (hasMin && numericValue < keyDef.min) {
          diagnostics.push(
            createDiagnostic(
              'error',
              filePath,
              entry.line,
              upperSection,
              entry.key,
              `Value ${effectiveValue} is below minimum ${keyDef.min} (allowed range: ${keyDef.min} to ${hasMax ? keyDef.max : rule.range ? rule.range[1] : '?'})`
            )
          );
        } else if (hasMax && numericValue > keyDef.max) {
          diagnostics.push(
            createDiagnostic(
              'error',
              filePath,
              entry.line,
              upperSection,
              entry.key,
              `Value ${effectiveValue} is above maximum ${keyDef.max} (allowed range: ${hasMin ? keyDef.min : rule.range ? rule.range[0] : '-?'} to ${keyDef.max})`
            )
          );
        }
      }
    }
  }

  return diagnostics;
}
