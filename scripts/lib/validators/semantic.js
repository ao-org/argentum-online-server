/**
 * Semantic Validator — checks cross-field rules: min/max consistency,
 * count-indexed sections, and admin entry format.
 *
 * Rule types processed from schema.rules[]:
 * - min_max: minKey value ? maxKey value within a section
 * - count_sections: count key N matches number of entries in target section
 * - admin_entry: values match Name|*@* format
 * - count_indexed: count key N means sections [PREFIX1]..[PREFIXN] exist,
 *   optionally with required keys
 *
 * @module validators/semantic
 */

import { createDiagnostic } from '../data-models.js';

/**
 * Admin entry pattern: non-empty name, pipe, string containing @.
 * @type {RegExp}
 */
const ADMIN_ENTRY_RE = /^[^|]+\|[^@]*@/;

/**
 * Build a lookup map of document sections by uppercase name.
 * @param {import('../data-models.js').IniDocument} document
 * @returns {Map<string, import('../data-models.js').IniSection[]>}
 */
function indexSections(document) {
  const map = new Map();
  for (const section of document.sections) {
    if (!map.has(section.name)) {
      map.set(section.name, [section]);
    } else {
      map.get(section.name).push(section);
    }
  }
  return map;
}

/**
 * Find the first entry with the given key in a section (first occurrence).
 * @param {Map<string, import('../data-models.js').IniSection[]>} sectionMap
 * @param {string} sectionName - Uppercase section name
 * @param {string} key - Uppercase key name
 * @returns {import('../data-models.js').IniEntry|null}
 */
function findEntry(sectionMap, sectionName, key) {
  const occurrences = sectionMap.get(sectionName.toUpperCase());
  if (!occurrences) return null;
  for (const section of occurrences) {
    for (const entry of section.entries) {
      if (entry.key === key.toUpperCase()) return entry;
    }
  }
  return null;
}

/**
 * Handle a min_max rule: verify minKey ? maxKey within the given section.
 * @param {Object} rule
 * @param {Map<string, import('../data-models.js').IniSection[]>} sectionMap
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
function handleMinMax(rule, sectionMap, filePath) {
  const diagnostics = [];
  const section = rule.section.toUpperCase();
  const minEntry = findEntry(sectionMap, section, rule.minKey);
  const maxEntry = findEntry(sectionMap, section, rule.maxKey);

  if (!minEntry || !maxEntry) return diagnostics;

  const minVal = Number(minEntry.value);
  const maxVal = Number(maxEntry.value);

  if (isNaN(minVal) || isNaN(maxVal)) return diagnostics;

  if (minVal > maxVal) {
    diagnostics.push(
      createDiagnostic(
        'error',
        filePath,
        minEntry.line,
        section,
        rule.minKey.toUpperCase(),
        `${rule.minKey.toUpperCase()} (${minVal}) must be less than or equal to ${rule.maxKey.toUpperCase()} (${maxVal})`
      )
    );
  }

  return diagnostics;
}

/**
 * Handle a count_sections rule: verify count key N matches the number of
 * entries in the target section.
 * @param {Object} rule
 * @param {Map<string, import('../data-models.js').IniSection[]>} sectionMap
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
function handleCountSections(rule, sectionMap, filePath) {
  const diagnostics = [];
  const countEntry = findEntry(sectionMap, rule.countSection, rule.countKey);
  if (!countEntry) return diagnostics;

  const expectedCount = Number(countEntry.value);
  if (isNaN(expectedCount) || !Number.isInteger(expectedCount)) return diagnostics;

  const targetSection = rule.targetSectionPrefix.toUpperCase();
  const occurrences = sectionMap.get(targetSection);

  // Count entries in the target section
  let actualCount = 0;
  if (occurrences) {
    for (const section of occurrences) {
      actualCount += section.entries.length;
    }
  }

  if (actualCount !== expectedCount) {
    diagnostics.push(
      createDiagnostic(
        'error',
        filePath,
        countEntry.line,
        rule.countSection.toUpperCase(),
        rule.countKey.toUpperCase(),
        `${rule.countKey.toUpperCase()} declares ${expectedCount} entries but section '${targetSection}' has ${actualCount}`
      )
    );
  }

  return diagnostics;
}

/**
 * Handle an admin_entry rule: verify all values in the section match Name|*@* format.
 * @param {Object} rule
 * @param {Map<string, import('../data-models.js').IniSection[]>} sectionMap
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
function handleAdminEntry(rule, sectionMap, filePath) {
  const diagnostics = [];
  const sectionName = rule.section.toUpperCase();
  const occurrences = sectionMap.get(sectionName);
  if (!occurrences) return diagnostics;

  for (const section of occurrences) {
    for (const entry of section.entries) {
      if (!ADMIN_ENTRY_RE.test(entry.value)) {
        diagnostics.push(
          createDiagnostic(
            'error',
            filePath,
            entry.line,
            sectionName,
            entry.key,
            `Invalid admin entry format '${entry.value}' — expected 'Name|user@domain'`
          )
        );
      }
    }
  }

  return diagnostics;
}

/**
 * Handle a count_indexed rule: verify count key N means sections
 * [PREFIX1]..[PREFIXN] exist, optionally with required keys.
 * @param {Object} rule
 * @param {Map<string, import('../data-models.js').IniSection[]>} sectionMap
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
function handleCountIndexed(rule, sectionMap, filePath) {
  const diagnostics = [];
  const countEntry = findEntry(sectionMap, rule.countSection, rule.countKey);
  if (!countEntry) return diagnostics;

  const expectedCount = Number(countEntry.value);
  if (isNaN(expectedCount) || !Number.isInteger(expectedCount)) return diagnostics;

  const prefix = rule.targetSectionPrefix.toUpperCase();
  const requiredKeys = (rule.requiredKeys || []).map(k => k.toUpperCase());

  for (let i = 1; i <= expectedCount; i++) {
    const expectedName = `${prefix}${i}`;
    const occurrences = sectionMap.get(expectedName);

    if (!occurrences || occurrences.length === 0) {
      diagnostics.push(
        createDiagnostic(
          'error',
          filePath,
          countEntry.line,
          expectedName,
          null,
          `Missing indexed section '${expectedName}' (expected by ${rule.countKey.toUpperCase()}=${expectedCount})`
        )
      );
      continue;
    }

    // Check required keys in the first occurrence
    if (requiredKeys.length > 0) {
      const section = occurrences[0];
      const presentKeys = new Set(section.entries.map(e => e.key));
      for (const reqKey of requiredKeys) {
        if (!presentKeys.has(reqKey)) {
          diagnostics.push(
            createDiagnostic(
              'error',
              filePath,
              section.line,
              expectedName,
              reqKey,
              `Missing required key '${reqKey}' in indexed section '${expectedName}'`
            )
          );
        }
      }
    }
  }

  return diagnostics;
}

/** @type {Record<string, (rule: Object, sectionMap: Map, filePath: string) => import('../data-models.js').Diagnostic[]>} */
const RULE_HANDLERS = {
  min_max: handleMinMax,
  count_sections: handleCountSections,
  admin_entry: handleAdminEntry,
  count_indexed: handleCountIndexed,
};

/**
 * Validate cross-field semantic rules declared in a ConfigSchema's rules array.
 *
 * @param {import('../data-models.js').IniDocument} document
 * @param {import('../data-models.js').ConfigSchema} schema
 * @param {string} filePath
 * @returns {import('../data-models.js').Diagnostic[]}
 */
export function validateSemantics(document, schema, filePath) {
  const diagnostics = [];
  const rules = schema.rules || [];
  const sectionMap = indexSections(document);

  for (const rule of rules) {
    const handler = RULE_HANDLERS[rule.type];
    if (handler) {
      diagnostics.push(...handler(rule, sectionMap, filePath));
    }
  }

  return diagnostics;
}
