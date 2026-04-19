/**
 * Property-based tests for INI Parser and Printer.
 *
 * Uses fast-check with Node.js built-in test runner.
 * Each property references its design document property via a tag comment.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fc from 'fast-check';
import { parseIni, printIni } from '../lib/ini-parser.js';
import { createIniDocument, createIniSection, createIniEntry } from '../lib/data-models.js';

// ---------------------------------------------------------------------------
// Custom Arbitraries
// ---------------------------------------------------------------------------

/**
 * arbUpperAlphaNum: Generates uppercase alphanumeric strings (1-20 chars).
 * Used for section names and keys.
 */
const arbUpperAlphaNum = fc
  .stringOf(
    fc.mapToConstant(
      { num: 26, build: (v) => String.fromCharCode(65 + v) }, // A-Z
      { num: 10, build: (v) => String.fromCharCode(48 + v) }  // 0-9
    ),
    { minLength: 1, maxLength: 20 }
  );

/**
 * arbPrintableValue: Generates printable ASCII strings excluding \n, \r.
 * Used for INI entry values.
 */
const arbPrintableValue = fc.stringOf(
  fc.integer({ min: 32, max: 126 }).map((c) => String.fromCharCode(c)),
  { minLength: 0, maxLength: 50 }
);

/**
 * arbIniEntry: Generates a single IniEntry with uppercase key and printable value.
 */
const arbIniEntry = fc.tuple(arbUpperAlphaNum, arbPrintableValue).map(
  ([key, value]) => createIniEntry(key, value, 0)
);

/**
 * arbIniSection: Generates an IniSection with a unique uppercase name and 0-10 entries
 * with unique keys.
 */
const arbIniSection = fc.tuple(
  arbUpperAlphaNum,
  fc.array(arbIniEntry, { minLength: 0, maxLength: 10 })
).map(([name, entries]) => {
  // Deduplicate keys within the section
  const seen = new Set();
  const uniqueEntries = [];
  for (const e of entries) {
    if (!seen.has(e.key)) {
      seen.add(e.key);
      uniqueEntries.push(e);
    }
  }
  return createIniSection(name, 0, uniqueEntries);
});


/**
 * arbIniDocument: Generates IniDocument with 1-5 sections, each with 0-10 entries.
 * Section names are uppercase alphanumeric (1-20 chars). Keys are uppercase alphanumeric.
 * Values are printable ASCII strings (excluding \n, \r). No duplicate section names or
 * keys within a section.
 */
const arbIniDocument = fc
  .array(arbIniSection, { minLength: 1, maxLength: 5 })
  .map((sections) => {
    // Deduplicate section names
    const seen = new Set();
    const uniqueSections = [];
    for (const s of sections) {
      if (!seen.has(s.name)) {
        seen.add(s.name);
        uniqueSections.push(s);
      }
    }
    return createIniDocument(uniqueSections, []);
  });

/**
 * arbMixedCaseName: Generates strings with mixed case (1-20 chars, alpha + digits).
 */
const arbMixedCaseName = fc.stringOf(
  fc.mapToConstant(
    { num: 26, build: (v) => String.fromCharCode(65 + v) }, // A-Z
    { num: 26, build: (v) => String.fromCharCode(97 + v) }, // a-z
    { num: 10, build: (v) => String.fromCharCode(48 + v) }  // 0-9
  ),
  { minLength: 1, maxLength: 20 }
);

/**
 * arbIniString: Generates raw INI text with valid sections, entries, comments,
 * and blank lines interspersed. Section and key names use mixed case.
 */
const arbIniString = fc
  .tuple(
    fc.array(
      fc.tuple(
        arbMixedCaseName,
        fc.array(
          fc.tuple(arbMixedCaseName, arbPrintableValue),
          { minLength: 0, maxLength: 10 }
        ),
        fc.array(
          fc.constantFrom('; a comment', '# another comment', "' vb comment", ''),
          { minLength: 0, maxLength: 3 }
        )
      ),
      { minLength: 1, maxLength: 5 }
    )
  )
  .map(([sections]) => {
    const lines = [];
    for (const [sectionName, entries, extras] of sections) {
      // Sprinkle some comments/blanks before section
      for (const e of extras) {
        lines.push(e);
      }
      lines.push(`[${sectionName}]`);
      for (const [key, value] of entries) {
        lines.push(`${key}=${value}`);
      }
    }
    return lines.join('\n');
  });

// ---------------------------------------------------------------------------
// Property Tests
// ---------------------------------------------------------------------------

describe('INI Parser Property Tests', () => {
  // Feature: config-file-validation, Property 1: INI round-trip
  // **Validates: Requirements 1.9, 1.10**
  it('Property 1: INI Parse/Print Round-Trip', () => {
    fc.assert(
      fc.property(arbIniDocument, (doc) => {
        const printed = printIni(doc);
        const { document: parsed } = parseIni(printed, 'test.ini');

        // Same number of sections
        assert.equal(parsed.sections.length, doc.sections.length,
          'Section count mismatch after round-trip');

        for (let i = 0; i < doc.sections.length; i++) {
          const original = doc.sections[i];
          const roundTripped = parsed.sections[i];

          assert.equal(roundTripped.name, original.name,
            `Section name mismatch at index ${i}`);
          assert.equal(roundTripped.entries.length, original.entries.length,
            `Entry count mismatch in section ${original.name}`);

          for (let j = 0; j < original.entries.length; j++) {
            assert.equal(roundTripped.entries[j].key, original.entries[j].key,
              `Key mismatch in section ${original.name} at entry ${j}`);
            assert.equal(roundTripped.entries[j].value, original.entries[j].value,
              `Value mismatch in section ${original.name} at entry ${j}`);
          }
        }

        // No orphans after round-trip
        assert.equal(parsed.orphans.length, 0, 'Round-trip should produce no orphans');
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 2: Case normalization
  // **Validates: Requirements 1.2, 1.3**
  it('Property 2: Case Normalization Invariant', () => {
    fc.assert(
      fc.property(arbIniString, (iniText) => {
        const { document } = parseIni(iniText, 'test.ini');

        for (const section of document.sections) {
          assert.equal(section.name, section.name.toUpperCase(),
            `Section name '${section.name}' is not uppercase`);

          for (const entry of section.entries) {
            assert.equal(entry.key, entry.key.toUpperCase(),
              `Key '${entry.key}' in section '${section.name}' is not uppercase`);
          }
        }
      }),
      { numRuns: 100 }
    );
  });


  // Feature: config-file-validation, Property 7: Duplicate detection
  // **Validates: Requirements 3.3, 3.4**
  it('Property 7: Duplicate Detection', () => {
    // Generate INI strings that intentionally contain duplicate sections and/or keys
    const arbDuplicateIni = fc.tuple(
      arbUpperAlphaNum, // section name to duplicate
      fc.array(
        fc.tuple(arbUpperAlphaNum, arbPrintableValue),
        { minLength: 1, maxLength: 5 }
      ),
      fc.boolean(), // whether to add duplicate section
      fc.boolean()  // whether to add duplicate key within section
    ).map(([sectionName, entries, dupSection, dupKey]) => {
      const lines = [];
      // Deduplicate entries for the base section
      const seenKeys = new Set();
      const uniqueEntries = [];
      for (const [k, v] of entries) {
        if (!seenKeys.has(k)) {
          seenKeys.add(k);
          uniqueEntries.push([k, v]);
        }
      }

      lines.push(`[${sectionName}]`);
      for (const [k, v] of uniqueEntries) {
        lines.push(`${k}=${v}`);
      }

      let expectedDupSections = 0;
      let expectedDupKeys = 0;

      // Add duplicate key within the same section
      if (dupKey && uniqueEntries.length > 0) {
        const [dupKeyName] = uniqueEntries[0];
        lines.push(`${dupKeyName}=duplicateValue`);
        expectedDupKeys = 1;
      }

      // Add duplicate section
      if (dupSection) {
        lines.push(`[${sectionName}]`);
        lines.push(`UNIQUEKEY${Date.now()}=val`);
        expectedDupSections = 1;
      }

      return { text: lines.join('\n'), expectedDupSections, expectedDupKeys };
    });

    fc.assert(
      fc.property(arbDuplicateIni, ({ text, expectedDupSections, expectedDupKeys }) => {
        const { diagnostics } = parseIni(text, 'test.ini');

        const dupSectionDiags = diagnostics.filter(
          (d) => d.message.toLowerCase().includes('duplicate section')
        );
        const dupKeyDiags = diagnostics.filter(
          (d) => d.message.toLowerCase().includes('duplicate key')
        );

        assert.equal(dupSectionDiags.length, expectedDupSections,
          `Expected ${expectedDupSections} duplicate section diagnostic(s), got ${dupSectionDiags.length}`);
        assert.equal(dupKeyDiags.length, expectedDupKeys,
          `Expected ${expectedDupKeys} duplicate key diagnostic(s), got ${dupKeyDiags.length}`);
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 12: Line ending normalization
  // **Validates: Requirements 8.4**
  it('Property 12: Line Ending Normalization', () => {
    fc.assert(
      fc.property(arbIniString, (iniText) => {
        // Ensure we start with pure \n line endings
        const unixText = iniText.replace(/\r\n/g, '\n');
        const windowsText = unixText.replace(/\n/g, '\r\n');

        const { document: unixDoc } = parseIni(unixText, 'test.ini');
        const { document: winDoc } = parseIni(windowsText, 'test.ini');

        // Same number of sections
        assert.equal(winDoc.sections.length, unixDoc.sections.length,
          'Section count differs between \\n and \\r\\n');

        for (let i = 0; i < unixDoc.sections.length; i++) {
          const uSec = unixDoc.sections[i];
          const wSec = winDoc.sections[i];

          assert.equal(wSec.name, uSec.name,
            `Section name mismatch at index ${i}`);
          assert.equal(wSec.entries.length, uSec.entries.length,
            `Entry count mismatch in section ${uSec.name}`);

          for (let j = 0; j < uSec.entries.length; j++) {
            assert.equal(wSec.entries[j].key, uSec.entries[j].key,
              `Key mismatch in section ${uSec.name} at entry ${j}`);
            assert.equal(wSec.entries[j].value, uSec.entries[j].value,
              `Value mismatch in section ${uSec.name} at entry ${j}`);
          }
        }

        // Orphans should also match
        assert.equal(winDoc.orphans.length, unixDoc.orphans.length,
          'Orphan count differs between \\n and \\r\\n');
      }),
      { numRuns: 100 }
    );
  });
});
