/**
 * Structural Validator — checks for missing required elements, unknown elements,
 * and duplicate sections/keys in an IniDocument against a ConfigSchema.
 *
 * @module validators/structural
 */

import { createDiagnostic } from '../data-models.js';

/**
 * Validate the structure of a parsed IniDocument against a ConfigSchema.
 *
 * Checks performed:
 * 1. Missing required sections (error)
 * 2. Missing required keys within present sections (error)
 * 3. Unknown sections not declared in schema (warning)
 * 4. Unknown keys not declared in schema section, respecting strict flag (warning)
 * 5. Duplicate sections (warning per 2nd+ occurrence)
 * 6. Duplicate keys within a section (warning per 2nd+ occurrence)
 *
 * @param {import('../data-models.js').IniDocument} document
 * @param {import('../data-models.js').ConfigSchema} schema
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
export function validateStructure(document, schema, filePath) {
  const diagnostics = [];
  const schemaSections = schema.sections || {};

  // Build a map of document sections by uppercase name for quick lookup.
  // Also track duplicates: first occurrence goes in the map, subsequent are flagged.
  /** @type {Map<string, import('../data-models.js').IniSection[]>} */
  const docSectionsByName = new Map();

  for (const section of document.sections) {
    const name = section.name;
    if (!docSectionsByName.has(name)) {
      docSectionsByName.set(name, [section]);
    } else {
      docSectionsByName.get(name).push(section);
    }
  }

  // --- Duplicate section detection (warning per 2nd+ occurrence) ---
  for (const [name, occurrences] of docSectionsByName) {
    if (occurrences.length > 1) {
      // Skip the first occurrence, warn on each subsequent one
      for (let i = 1; i < occurrences.length; i++) {
        diagnostics.push(
          createDiagnostic(
            'warning',
            filePath,
            occurrences[i].line,
            name,
            null,
            `Duplicate section '${name}' (first defined on line ${occurrences[0].line})`
          )
        );
      }
    }
  }

  // --- Missing required sections (error per missing) ---
  for (const [sectionName, sectionDef] of Object.entries(schemaSections)) {
    const upperName = sectionName.toUpperCase();
    if (sectionDef.required && !docSectionsByName.has(upperName)) {
      diagnostics.push(
        createDiagnostic(
          'error',
          filePath,
          0,
          upperName,
          null,
          `Missing required section '${upperName}'`
        )
      );
    }
  }

  // --- Unknown sections (warning) ---
  // Build a set of section prefixes from count_sections rules so that
  // dynamically-indexed sections (e.g., TOGGLE1..TOGGLEN) are not flagged
  // as unknown.
  const dynamicPrefixes = new Set();
  if (schema.rules) {
    for (const rule of schema.rules) {
      if ((rule.type === 'count_sections' || rule.type === 'count_indexed') && rule.targetSectionPrefix) {
        dynamicPrefixes.add(rule.targetSectionPrefix.toUpperCase());
      }
    }
  }

  const schemaUpperNames = new Set(
    Object.keys(schemaSections).map(n => n.toUpperCase())
  );

  for (const [name, occurrences] of docSectionsByName) {
    if (!schemaUpperNames.has(name)) {
      // Check if this section matches a dynamic prefix (e.g., TOGGLE1 matches TOGGLE)
      let isDynamic = false;
      for (const prefix of dynamicPrefixes) {
        if (name.startsWith(prefix) && /^\d+$/.test(name.slice(prefix.length))) {
          isDynamic = true;
          break;
        }
      }
      if (isDynamic) continue;

      // Warn on the first occurrence only (duplicates already warned above)
      diagnostics.push(
        createDiagnostic(
          'warning',
          filePath,
          occurrences[0].line,
          name,
          null,
          `Unknown section '${name}' not declared in schema`
        )
      );
    }
  }

  // --- Per-section key validation ---
  for (const [sectionName, sectionDef] of Object.entries(schemaSections)) {
    const upperName = sectionName.toUpperCase();
    const occurrences = docSectionsByName.get(upperName);
    if (!occurrences) {
      // Section not present — already handled by missing required check
      continue;
    }

    const schemaKeys = sectionDef.keys || {};
    const schemaUpperKeys = new Set(
      Object.keys(schemaKeys).map(k => k.toUpperCase())
    );

    // Determine strict mode: default is true if omitted
    const strict = sectionDef.strict !== false;

    // Process each occurrence of this section
    for (const section of occurrences) {
      // Track keys seen in this section occurrence for duplicate detection
      /** @type {Map<string, number>} key name ? first line */
      const seenKeys = new Map();

      for (const entry of section.entries) {
        const key = entry.key;

        // --- Duplicate key detection (warning per 2nd+ occurrence) ---
        if (seenKeys.has(key)) {
          diagnostics.push(
            createDiagnostic(
              'warning',
              filePath,
              entry.line,
              upperName,
              key,
              `Duplicate key '${key}' in section '${upperName}' (first defined on line ${seenKeys.get(key)})`
            )
          );
        } else {
          seenKeys.set(key, entry.line);
        }

        // --- Unknown key detection (warning if strict) ---
        if (!schemaUpperKeys.has(key) && strict) {
          diagnostics.push(
            createDiagnostic(
              'warning',
              filePath,
              entry.line,
              upperName,
              key,
              `Unknown key '${key}' in section '${upperName}'`
            )
          );
        }
      }

      // --- Missing required keys (error per missing) ---
      for (const [keyName, keyDef] of Object.entries(schemaKeys)) {
        const upperKey = keyName.toUpperCase();
        if (keyDef.required && !seenKeys.has(upperKey)) {
          diagnostics.push(
            createDiagnostic(
              'error',
              filePath,
              section.line,
              upperName,
              upperKey,
              `Missing required key '${upperKey}' in section '${upperName}'`
            )
          );
        }
      }
    }
  }

  return diagnostics;
}
