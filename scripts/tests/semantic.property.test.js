/**
 * Property-based tests for Semantic Validator.
 *
 * Uses fast-check with Node.js built-in test runner.
 * Each property references its design document property via a tag comment.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fc from 'fast-check';
import { validateSemantics } from '../lib/validators/semantic.js';
import {
  createIniDocument,
  createIniSection,
  createIniEntry,
  createConfigSchema,
} from '../lib/data-models.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Admin entry regex — mirrors the one in semantic.js */
const ADMIN_ENTRY_RE = /^[^|]+\|[^@]*@/;

// ---------------------------------------------------------------------------
// Property Tests
// ---------------------------------------------------------------------------

describe('Semantic Validator Property Tests', () => {

  // Feature: config-file-validation, Property 8: Min/max cross-field
  // **Validates: Requirements 5.1, 5.6**
  it('Property 8: Min/Max Cross-Field Validation', () => {
    /**
     * Strategy:
     * Generate two integer values (minVal, maxVal). Create a document with a
     * section containing both keys. Create a schema with a min_max rule.
     * Run validateSemantics. Assert error diagnostic iff minVal > maxVal.
     */
    fc.assert(
      fc.property(
        fc.integer({ min: -10000, max: 10000 }),
        fc.integer({ min: -10000, max: 10000 }),
        (minVal, maxVal) => {
          const doc = createIniDocument([
            createIniSection('INIT', 1, [
              createIniEntry('MINKEY', String(minVal), 2),
              createIniEntry('MAXKEY', String(maxVal), 3),
            ]),
          ]);

          const schema = createConfigSchema('test.ini', 'test', {
            INIT: { required: true, strict: false, keys: {} },
          }, [
            {
              type: 'min_max',
              section: 'INIT',
              minKey: 'MINKEY',
              maxKey: 'MAXKEY',
            },
          ]);

          const diagnostics = validateSemantics(doc, schema, 'test.ini');
          const errors = diagnostics.filter((d) => d.severity === 'error');

          if (minVal > maxVal) {
            assert.equal(errors.length, 1,
              `Expected 1 error when min (${minVal}) > max (${maxVal}), got ${errors.length}`);
            assert.ok(errors[0].message.includes('MINKEY'),
              'Error message should reference the min key');
            assert.ok(errors[0].message.includes('MAXKEY'),
              'Error message should reference the max key');
          } else {
            assert.equal(errors.length, 0,
              `Expected 0 errors when min (${minVal}) <= max (${maxVal}), got ${errors.length}`);
          }
        }
      ),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 9: Count-indexed sections
  // **Validates: Requirements 5.2, 5.4, 5.8**
  it('Property 9: Count-Indexed Section Validation', () => {
    /**
     * Strategy:
     * Generate a count N (1-10) and a number of actual indexed sections M (0-10).
     * Create a document with a count section containing COUNTKEY=N and M indexed
     * sections [PREFIX1]..[PREFIXM]. Create a schema with a count_indexed rule.
     * Assert error diagnostics for each missing section (when M < N).
     */
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 10 }),
        fc.integer({ min: 0, max: 10 }),
        (expectedCount, actualCount) => {
          let lineCounter = 1;

          // Build sections: count section + actual indexed sections
          const sections = [
            createIniSection('INIT', lineCounter++, [
              createIniEntry('COUNTKEY', String(expectedCount), lineCounter++),
            ]),
          ];

          for (let i = 1; i <= actualCount; i++) {
            sections.push(
              createIniSection(`ITEM${i}`, lineCounter++, [
                createIniEntry('NAME', `item_${i}`, lineCounter++),
              ])
            );
          }

          const doc = createIniDocument(sections);

          const schema = createConfigSchema('test.ini', 'test', {
            INIT: { required: true, strict: false, keys: {} },
          }, [
            {
              type: 'count_indexed',
              countSection: 'INIT',
              countKey: 'COUNTKEY',
              targetSectionPrefix: 'ITEM',
            },
          ]);

          const diagnostics = validateSemantics(doc, schema, 'test.ini');
          const errors = diagnostics.filter((d) => d.severity === 'error');

          // Count how many sections from 1..expectedCount are missing
          const missingSections = [];
          for (let i = 1; i <= expectedCount; i++) {
            if (i > actualCount) {
              missingSections.push(`ITEM${i}`);
            }
          }

          assert.equal(errors.length, missingSections.length,
            `Expected ${missingSections.length} error(s) for N=${expectedCount}, M=${actualCount}, got ${errors.length}`);

          // Each error should reference a missing indexed section
          for (const err of errors) {
            assert.ok(err.message.toLowerCase().includes('missing indexed section'),
              `Error message should mention missing indexed section: ${err.message}`);
          }
        }
      ),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 10: Admin entry format
  // **Validates: Requirements 5.3**
  it('Property 10: Admin Entry Format Validation', () => {
    /**
     * Strategy:
     * Generate random strings. Create a document with an admin section containing
     * the string as a value. Create a schema with an admin_entry rule. Assert
     * error diagnostic iff the string does NOT match the pattern /^[^|]+\|[^@]*@/.
     */
    fc.assert(
      fc.property(
        fc.stringOf(
          fc.oneof(
            fc.char(),                          // any single char
            fc.constant('|'),                   // pipe
            fc.constant('@'),                   // at sign
            fc.constant('.'),                   // dot
          ),
          { minLength: 1, maxLength: 30 }
        ),
        (adminValue) => {
          const doc = createIniDocument([
            createIniSection('ADMINES', 1, [
              createIniEntry('ADMIN1', adminValue, 2),
            ]),
          ]);

          const schema = createConfigSchema('test.ini', 'test', {
            ADMINES: { required: false, strict: false, keys: {} },
          }, [
            {
              type: 'admin_entry',
              section: 'ADMINES',
            },
          ]);

          const diagnostics = validateSemantics(doc, schema, 'test.ini');
          const errors = diagnostics.filter((d) => d.severity === 'error');
          const isValid = ADMIN_ENTRY_RE.test(adminValue);

          if (isValid) {
            assert.equal(errors.length, 0,
              `Expected 0 errors for valid admin entry '${adminValue}', got ${errors.length}`);
          } else {
            assert.equal(errors.length, 1,
              `Expected 1 error for invalid admin entry '${adminValue}', got ${errors.length}`);
            assert.ok(errors[0].message.toLowerCase().includes('invalid admin entry'),
              `Error message should mention invalid admin entry: ${errors[0].message}`);
          }
        }
      ),
      { numRuns: 100 }
    );
  });
});
