/**
 * INI Parser and Printer module.
 *
 * Mirrors the behavior of the VB6 `clsIniManager`:
 * - Section names and keys are normalized to uppercase.
 * - Lines starting with `'`, `;`, or `#` are comments.
 * - Key-value pairs are split on the first `=`.
 * - Produces warning diagnostics for structural issues.
 *
 * @module ini-parser
 */

import {
  createIniDocument,
  createIniSection,
  createIniEntry,
  createDiagnostic,
} from './data-models.js';

/**
 * Parse an INI-format string into an IniDocument.
 * @param {string} content - Raw file content
 * @param {string} filePath - File path (for diagnostic messages)
 * @returns {{ document: import('./data-models.js').IniDocument, diagnostics: import('./data-models.js').Diagnostic[] }}
 */
export function parseIni(content, filePath) {
  const diagnostics = [];
  const sections = [];
  const orphans = [];

  // Normalize line endings: convert \r\n to \n, then split
  const lines = content.replace(/\r\n/g, '\n').split('\n');

  /** @type {import('./data-models.js').IniSection | null} */
  let currentSection = null;

  /** Track seen section names (uppercase) for duplicate detection */
  const seenSections = new Map();

  /** Track seen keys per section (uppercase) for duplicate detection */
  let seenKeysInSection = new Map();

  for (let i = 0; i < lines.length; i++) {
    const lineNumber = i + 1;
    const raw = lines[i];
    const trimmed = raw.trim();

    // Skip empty / whitespace-only lines
    if (trimmed === '') {
      continue;
    }

    // Skip comment lines
    if (trimmed[0] === "'" || trimmed[0] === ';' || trimmed[0] === '#') {
      continue;
    }

    // Section header detection
    if (trimmed[0] === '[') {
      const closingBracket = trimmed.indexOf(']');
      if (closingBracket === -1) {
        // Malformed section header — `[` without `]`
        diagnostics.push(
          createDiagnostic(
            'warning',
            filePath,
            lineNumber,
            null,
            null,
            `Malformed section header (missing closing ']'): ${trimmed}`
          )
        );
        continue;
      }

      const sectionName = trimmed.substring(1, closingBracket).trim().toUpperCase();

      // Duplicate section detection
      if (seenSections.has(sectionName)) {
        diagnostics.push(
          createDiagnostic(
            'warning',
            filePath,
            lineNumber,
            sectionName,
            null,
            `Duplicate section '${sectionName}' (first defined on line ${seenSections.get(sectionName)})`
          )
        );
      }
      seenSections.set(sectionName, lineNumber);

      currentSection = createIniSection(sectionName, lineNumber);
      sections.push(currentSection);
      seenKeysInSection = new Map();
      continue;
    }

    // Key-value pair detection
    const eqIndex = raw.indexOf('=');
    if (eqIndex !== -1) {
      const key = raw.substring(0, eqIndex).trim().toUpperCase();
      const value = raw.substring(eqIndex + 1);

      if (currentSection === null) {
        // Orphaned key — before any section header
        orphans.push(createIniEntry(key, value, lineNumber));
        diagnostics.push(
          createDiagnostic(
            'warning',
            filePath,
            lineNumber,
            null,
            key,
            `Orphaned key '${key}' appears before any section header`
          )
        );
      } else {
        // Duplicate key detection within current section
        if (seenKeysInSection.has(key)) {
          diagnostics.push(
            createDiagnostic(
              'warning',
              filePath,
              lineNumber,
              currentSection.name,
              key,
              `Duplicate key '${key}' in section '${currentSection.name}' (first defined on line ${seenKeysInSection.get(key)})`
            )
          );
        }
        seenKeysInSection.set(key, lineNumber);

        currentSection.entries.push(createIniEntry(key, value, lineNumber));
      }
      continue;
    }

    // Non-empty, non-comment line without `=` inside a section
    if (currentSection !== null) {
      diagnostics.push(
        createDiagnostic(
          'warning',
          filePath,
          lineNumber,
          currentSection.name,
          null,
          `Line without '=' in section '${currentSection.name}': ${trimmed}`
        )
      );
    } else {
      // Non-empty, non-comment, non-section, no `=`, before any section
      diagnostics.push(
        createDiagnostic(
          'warning',
          filePath,
          lineNumber,
          null,
          null,
          `Unrecognized line before any section header: ${trimmed}`
        )
      );
    }
  }

  const document = createIniDocument(sections, orphans);
  return { document, diagnostics };
}


/**
 * Serialize an IniDocument back to an INI-format string.
 * @param {import('./data-models.js').IniDocument} document
 * @returns {string}
 */
export function printIni(document) {
  const parts = [];

  for (let i = 0; i < document.sections.length; i++) {
    const section = document.sections[i];

    // Add blank line separator between sections (not before the first)
    if (i > 0) {
      parts.push('');
    }

    parts.push(`[${section.name}]`);

    for (const entry of section.entries) {
      parts.push(`${entry.key}=${entry.value}`);
    }
  }

  // Trailing newline for well-formed file output
  return parts.join('\n') + '\n';
}
