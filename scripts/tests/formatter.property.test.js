/**
 * Property-based tests for Diagnostic Formatter.
 *
 * Uses fast-check with Node.js built-in test runner.
 * Each property references its design document property via a tag comment.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fc from 'fast-check';
import { formatText, formatJson, computeExitCode } from '../lib/diagnostic-formatter.js';
import { createDiagnostic } from '../lib/data-models.js';

// ---------------------------------------------------------------------------
// Custom Arbitraries
// ---------------------------------------------------------------------------

/**
 * arbSeverity: Generates one of the three valid severity levels.
 */
const arbSeverity = fc.constantFrom('error', 'warning', 'info');

/**
 * arbNonEmptyString: Generates non-empty printable ASCII strings (1-50 chars).
 */
const arbNonEmptyString = fc.stringOf(
  fc.integer({ min: 32, max: 126 }).map((c) => String.fromCharCode(c)),
  { minLength: 1, maxLength: 50 }
);

/**
 * arbFilePath: Generates file-path-like strings (e.g., "Server.ini", "path/to/file.dat").
 */
const arbFilePath = fc.tuple(
  fc.stringOf(
    fc.mapToConstant(
      { num: 26, build: (v) => String.fromCharCode(97 + v) }, // a-z
      { num: 10, build: (v) => String.fromCharCode(48 + v) }, // 0-9
      { num: 1, build: () => '/' },
      { num: 1, build: () => '.' }
    ),
    { minLength: 1, maxLength: 30 }
  )
).map(([p]) => p);

/**
 * arbDiagnostic: Generates Diagnostic objects with random severity ('error'|'warning'|'info'),
 * file path, line number (>=0), section (string or null), key (string or null),
 * and message (non-empty string).
 */
const arbDiagnostic = fc.tuple(
  arbSeverity,
  arbFilePath,
  fc.nat({ max: 10000 }),
  fc.option(arbNonEmptyString, { nil: null }),
  fc.option(arbNonEmptyString, { nil: null }),
  arbNonEmptyString
).map(([severity, file, line, section, key, message]) =>
  createDiagnostic(severity, file, line, section, key, message)
);

/**
 * arbDiagnosticList: Generates arrays of 0-20 diagnostics.
 */
const arbDiagnosticList = fc.array(arbDiagnostic, { minLength: 0, maxLength: 20 });

// ---------------------------------------------------------------------------
// Property Tests
// ---------------------------------------------------------------------------

describe('Diagnostic Formatter Property Tests', () => {

  // Feature: config-file-validation, Property 16: JSON output validity
  // **Validates: Requirements 7.6**
  it('Property 16: JSON Diagnostic Output Validity', () => {
    fc.assert(
      fc.property(arbDiagnosticList, (diagnostics) => {
        const jsonStr = formatJson(diagnostics);

        // Must be valid JSON
        let parsed;
        assert.doesNotThrow(() => {
          parsed = JSON.parse(jsonStr);
        }, 'formatJson output must be valid JSON');

        // Must be an array of the same length
        assert.ok(Array.isArray(parsed), 'Parsed JSON must be an array');
        assert.equal(parsed.length, diagnostics.length,
          'Parsed array length must match original diagnostics length');

        // Each element must have matching fields
        for (let i = 0; i < diagnostics.length; i++) {
          const original = diagnostics[i];
          const restored = parsed[i];

          assert.equal(restored.severity, original.severity,
            `severity mismatch at index ${i}`);
          assert.equal(restored.file, original.file,
            `file mismatch at index ${i}`);
          assert.equal(restored.line, original.line,
            `line mismatch at index ${i}`);
          assert.equal(restored.message, original.message,
            `message mismatch at index ${i}`);
          assert.equal(restored.section, original.section,
            `section mismatch at index ${i}`);
          assert.equal(restored.key, original.key,
            `key mismatch at index ${i}`);
        }
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 17: Diagnostic completeness
  // **Validates: Requirements 6.1, 6.2**
  it('Property 17: Diagnostic Completeness', () => {
    fc.assert(
      fc.property(arbDiagnosticList, (diagnostics) => {
        for (let i = 0; i < diagnostics.length; i++) {
          const d = diagnostics[i];

          // Non-empty file path
          assert.ok(typeof d.file === 'string' && d.file.length > 0,
            `Diagnostic at index ${i} must have a non-empty file path`);

          // Line number >= 0
          assert.ok(typeof d.line === 'number' && d.line >= 0,
            `Diagnostic at index ${i} must have line >= 0`);

          // Non-empty message
          assert.ok(typeof d.message === 'string' && d.message.length > 0,
            `Diagnostic at index ${i} must have a non-empty message`);

          // Severity is one of the valid values
          assert.ok(['error', 'warning', 'info'].includes(d.severity),
            `Diagnostic at index ${i} must have severity in {error, warning, info}, got '${d.severity}'`);
        }
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 11: Exit code
  // **Validates: Requirements 6.4, 6.5**
  it('Property 11: Exit Code Reflects Error Presence', () => {
    fc.assert(
      fc.property(arbDiagnosticList, (diagnostics) => {
        const exitCode = computeExitCode(diagnostics);
        const hasError = diagnostics.some((d) => d.severity === 'error');

        if (hasError) {
          assert.equal(exitCode, 1,
            'Exit code must be 1 when at least one error diagnostic exists');
        } else {
          assert.equal(exitCode, 0,
            'Exit code must be 0 when no error diagnostics exist');
        }
      }),
      { numRuns: 100 }
    );
  });
});
