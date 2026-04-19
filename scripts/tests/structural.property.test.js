/**
 * Property-based tests for Structural Validator.
 *
 * Uses fast-check with Node.js built-in test runner.
 * Each property references its design document property via a tag comment.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fc from 'fast-check';
import { validateStructure } from '../lib/validators/structural.js';
import {
  createIniDocument,
  createIniSection,
  createIniEntry,
  createConfigSchema,
} from '../lib/data-models.js';

// ---------------------------------------------------------------------------
// Custom Arbitraries
// ---------------------------------------------------------------------------

/**
 * arbUpperAlphaNum: Generates uppercase alphanumeric strings (1-12 chars).
 * Used for section names and keys.
 */
const arbUpperAlphaNum = fc.stringOf(
  fc.mapToConstant(
    { num: 26, build: (v) => String.fromCharCode(65 + v) }, // A-Z
    { num: 10, build: (v) => String.fromCharCode(48 + v) }  // 0-9
  ),
  { minLength: 1, maxLength: 12 }
);

/**
 * Generate a set of N unique uppercase alphanumeric names.
 */
function arbUniqueNames(minCount, maxCount) {
  return fc
    .uniqueArray(arbUpperAlphaNum, {
      minLength: minCount,
      maxLength: maxCount,
      comparator: (a, b) => a === b,
    });
}


// ---------------------------------------------------------------------------
// Property Tests
// ---------------------------------------------------------------------------

describe('Structural Validator Property Tests', () => {
  // Feature: config-file-validation, Property 5: Missing required elements
  // **Validates: Requirements 3.5, 3.6**
  it('Property 5: Missing Required Elements Detection', () => {
    /**
     * Strategy:
     * 1. Generate a list of unique section names, each with unique required keys.
     * 2. Randomly select a subset of sections to INCLUDE in the document (the rest are missing).
     * 3. For included sections, randomly select a subset of keys to INCLUDE (the rest are missing).
     * 4. Count total missing required sections + missing required keys = expected error count.
     * 5. Assert validateStructure produces exactly that many error diagnostics.
     */
    const arbMissingElementsClean = fc.record({
      sectionNames: arbUniqueNames(1, 4),
      seed: fc.nat(),
    }).chain(({ sectionNames, seed }) => {
      return fc.tuple(
        fc.constant(sectionNames),
        // For each section, generate unique key names
        fc.tuple(...sectionNames.map(() => arbUniqueNames(1, 5))),
        // For each section, whether to include it in the doc
        fc.tuple(...sectionNames.map(() => fc.boolean())),
        // For each section, for each key, whether to include it
        // We'll generate arrays of booleans and trim later
        fc.tuple(...sectionNames.map(() => fc.array(fc.boolean(), { minLength: 5, maxLength: 5 }))),
      );
    }).map(([sectionNames, keysPerSection, includeSection, includeKeysPerSection]) => {
      const schemaSections = {};
      const docSections = [];
      let expectedMissingSections = 0;
      let expectedMissingKeys = 0;
      let lineCounter = 1;

      for (let i = 0; i < sectionNames.length; i++) {
        const secName = sectionNames[i];
        const keys = keysPerSection[i];
        const secIncluded = includeSection[i];
        const keyBools = includeKeysPerSection[i];

        // Build schema section — all keys required
        const schemaKeys = {};
        for (const k of keys) {
          schemaKeys[k] = { type: 'string', required: true };
        }
        schemaSections[secName] = { required: true, strict: false, keys: schemaKeys };

        if (!secIncluded) {
          expectedMissingSections++;
          // Validator does NOT report missing keys for absent sections
        } else {
          const entries = [];
          for (let j = 0; j < keys.length; j++) {
            if (keyBools[j]) {
              entries.push(createIniEntry(keys[j], 'val', lineCounter++));
            } else {
              expectedMissingKeys++;
            }
          }
          docSections.push(createIniSection(secName, lineCounter++, entries));
        }
      }

      const schema = createConfigSchema('test.ini', 'test', schemaSections, []);
      const document = createIniDocument(docSections, []);
      const totalExpectedErrors = expectedMissingSections + expectedMissingKeys;

      return { schema, document, totalExpectedErrors, expectedMissingSections, expectedMissingKeys };
    });

    fc.assert(
      fc.property(arbMissingElementsClean, ({ schema, document, totalExpectedErrors, expectedMissingSections, expectedMissingKeys }) => {
        const diagnostics = validateStructure(document, schema, 'test.ini');

        // Filter to only error-severity diagnostics about missing elements
        const missingDiags = diagnostics.filter(
          (d) => d.severity === 'error' && d.message.toLowerCase().includes('missing required')
        );

        const missingSectionDiags = missingDiags.filter(
          (d) => d.message.toLowerCase().includes('missing required section')
        );
        const missingKeyDiags = missingDiags.filter(
          (d) => d.message.toLowerCase().includes('missing required key')
        );

        assert.equal(
          missingSectionDiags.length,
          expectedMissingSections,
          `Expected ${expectedMissingSections} missing section error(s), got ${missingSectionDiags.length}`
        );
        assert.equal(
          missingKeyDiags.length,
          expectedMissingKeys,
          `Expected ${expectedMissingKeys} missing key error(s), got ${missingKeyDiags.length}`
        );
        assert.equal(
          missingDiags.length,
          totalExpectedErrors,
          `Expected ${totalExpectedErrors} total missing element error(s), got ${missingDiags.length}`
        );
      }),
      { numRuns: 100 }
    );
  });


  // Feature: config-file-validation, Property 6: Unknown elements with strict flag
  // **Validates: Requirements 3.7, 3.8, 9.12, 9.13**
  it('Property 6: Unknown Elements Detection with Strict Flag', () => {
    /**
     * Strategy:
     * 1. Generate a schema with known sections and keys, each section having a random strict flag.
     * 2. Generate an IniDocument that contains:
     *    - All schema-declared sections (so no missing-section errors)
     *    - All required keys (so no missing-key errors)
     *    - Some EXTRA sections not in the schema
     *    - Some EXTRA keys in each section not in the schema
     * 3. Assert:
     *    - One warning per unknown section
     *    - One warning per unknown key in strict sections
     *    - Zero warnings for unknown keys in non-strict sections
     */
    const arbUnknownElements = fc.record({
      // Schema section names (known)
      knownSectionNames: arbUniqueNames(1, 3),
    }).chain(({ knownSectionNames }) => {
      return fc.tuple(
        fc.constant(knownSectionNames),
        // For each known section: unique key names
        fc.tuple(...knownSectionNames.map(() => arbUniqueNames(1, 3))),
        // For each known section: strict flag
        fc.tuple(...knownSectionNames.map(() => fc.boolean())),
        // Extra section names (must not collide with known)
        arbUniqueNames(0, 3),
        // For each known section: extra key names to add (must not collide with schema keys)
        fc.tuple(...knownSectionNames.map(() => arbUniqueNames(0, 3))),
      );
    }).map(([knownSectionNames, keysPerSection, strictFlags, extraSectionCandidates, extraKeysPerSection]) => {
      const knownSet = new Set(knownSectionNames);
      // Filter extra sections to avoid collision with known sections
      const extraSections = extraSectionCandidates.filter((s) => !knownSet.has(s));

      const schemaSections = {};
      const docSections = [];
      let lineCounter = 1;
      let expectedUnknownSectionWarnings = extraSections.length;
      let expectedUnknownKeyWarnings = 0;

      // Build schema and document for known sections
      for (let i = 0; i < knownSectionNames.length; i++) {
        const secName = knownSectionNames[i];
        const schemaKeyNames = keysPerSection[i];
        const strict = strictFlags[i];
        const extraKeysCandidates = extraKeysPerSection[i];

        // Filter extra keys to avoid collision with schema keys
        const schemaKeySet = new Set(schemaKeyNames);
        const extraKeys = extraKeysCandidates.filter((k) => !schemaKeySet.has(k));

        // Schema definition
        const schemaKeys = {};
        for (const k of schemaKeyNames) {
          schemaKeys[k] = { type: 'string', required: true };
        }
        schemaSections[secName] = { required: false, strict, keys: schemaKeys };

        // Document section: include all schema keys + extra keys
        const entries = [];
        for (const k of schemaKeyNames) {
          entries.push(createIniEntry(k, 'val', lineCounter++));
        }
        for (const k of extraKeys) {
          entries.push(createIniEntry(k, 'extraval', lineCounter++));
        }
        docSections.push(createIniSection(secName, lineCounter++, entries));

        // Count expected unknown key warnings
        if (strict) {
          expectedUnknownKeyWarnings += extraKeys.length;
        }
        // If strict is false, no warnings for extra keys
      }

      // Add extra sections to document (not in schema)
      for (const secName of extraSections) {
        docSections.push(
          createIniSection(secName, lineCounter++, [
            createIniEntry('SOMEKEY', 'someval', lineCounter++),
          ])
        );
      }

      const schema = createConfigSchema('test.ini', 'test', schemaSections, []);
      const document = createIniDocument(docSections, []);

      return {
        schema,
        document,
        expectedUnknownSectionWarnings,
        expectedUnknownKeyWarnings,
      };
    });

    fc.assert(
      fc.property(arbUnknownElements, ({ schema, document, expectedUnknownSectionWarnings, expectedUnknownKeyWarnings }) => {
        const diagnostics = validateStructure(document, schema, 'test.ini');

        const unknownSectionDiags = diagnostics.filter(
          (d) => d.severity === 'warning' && d.message.toLowerCase().includes('unknown section')
        );
        const unknownKeyDiags = diagnostics.filter(
          (d) => d.severity === 'warning' && d.message.toLowerCase().includes('unknown key')
        );

        assert.equal(
          unknownSectionDiags.length,
          expectedUnknownSectionWarnings,
          `Expected ${expectedUnknownSectionWarnings} unknown section warning(s), got ${unknownSectionDiags.length}`
        );
        assert.equal(
          unknownKeyDiags.length,
          expectedUnknownKeyWarnings,
          `Expected ${expectedUnknownKeyWarnings} unknown key warning(s), got ${unknownKeyDiags.length}`
        );
      }),
      { numRuns: 100 }
    );
  });
});
