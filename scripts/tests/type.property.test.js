/**
 * Property-based tests for Type Validator.
 *
 * Uses fast-check with Node.js built-in test runner.
 * Each property references its design document property via a tag comment.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fc from 'fast-check';
import { validateTypes, vb6Val } from '../lib/validators/type.js';
import {
  createIniDocument,
  createIniSection,
  createIniEntry,
  createConfigSchema,
} from '../lib/data-models.js';

// ---------------------------------------------------------------------------
// Type rules (mirrors the validator logic for oracle comparison)
// ---------------------------------------------------------------------------

const TYPE_RANGES = {
  integer: { min: -32768, max: 32767 },
  long: { min: -2147483648, max: 2147483647 },
};

/**
 * Oracle: returns true if value V is valid for type T according to the spec.
 * For numeric types, applies vb6Val() stripping first (matching validator behavior).
 */
function isValidForType(type, value) {
  // For numeric types, strip trailing text like the validator does
  const effectiveValue = (type !== 'string' && type !== 'boolean')
    ? vb6Val(value)
    : value.trim();

  switch (type) {
    case 'integer': {
      if (!/^-?\d+$/.test(effectiveValue)) return false;
      const n = Number(effectiveValue);
      return n >= TYPE_RANGES.integer.min && n <= TYPE_RANGES.integer.max;
    }
    case 'long': {
      if (!/^-?\d+$/.test(effectiveValue)) return false;
      const n = Number(effectiveValue);
      return n >= TYPE_RANGES.long.min && n <= TYPE_RANGES.long.max;
    }
    case 'single':
    case 'double': {
      if (effectiveValue.trim() === '') return false;
      const n = parseFloat(effectiveValue);
      return !isNaN(n) && isFinite(n);
    }
    case 'boolean':
      return effectiveValue === '0' || effectiveValue === '1';
    case 'string':
      return true;
    default:
      return false;
  }
}

// ---------------------------------------------------------------------------
// Custom Arbitraries
// ---------------------------------------------------------------------------

/**
 * arbUpperAlphaNum: Generates uppercase alphanumeric strings (1-12 chars).
 */
const arbUpperAlphaNum = fc.stringOf(
  fc.mapToConstant(
    { num: 26, build: (v) => String.fromCharCode(65 + v) }, // A-Z
    { num: 10, build: (v) => String.fromCharCode(48 + v) }  // 0-9
  ),
  { minLength: 1, maxLength: 12 }
);

/**
 * arbType: Generates one of the six supported schema types.
 */
const arbType = fc.constantFrom('integer', 'long', 'single', 'double', 'boolean', 'string');

/**
 * For a given type, generate a value that is VALID for that type.
 */
function arbValidValueForType(type) {
  switch (type) {
    case 'integer':
      return fc.integer({ min: -32768, max: 32767 }).map(String);
    case 'long':
      return fc.integer({ min: -2147483648, max: 2147483647 }).map(String);
    case 'single':
    case 'double':
      return fc.oneof(
        fc.integer({ min: -100000, max: 100000 }).map(String),
        fc.double({ min: -1e10, max: 1e10, noNaN: true, noDefaultInfinity: true })
          .filter((n) => isFinite(n))
          .map((n) => String(n))
      );
    case 'boolean':
      return fc.constantFrom('0', '1');
    case 'string':
      return fc.string({ minLength: 0, maxLength: 20 });
    default:
      return fc.constant('');
  }
}

/**
 * For a given type, generate a value that is INVALID for that type.
 * String type always accepts, so we skip it.
 */
function arbInvalidValueForType(type) {
  switch (type) {
    case 'integer':
      return fc.oneof(
        // Float values (decimal notation, not scientific — vb6Val preserves these)
        fc.tuple(
          fc.integer({ min: -1000, max: 1000 }),
          fc.integer({ min: 1, max: 99 })
        ).map(([whole, frac]) => `${whole}.${frac}`),
        // Out-of-range integers
        fc.oneof(
          fc.integer({ min: 32768, max: 100000 }).map(String),
          fc.integer({ min: -100000, max: -32769 }).map(String)
        ),
        // Non-numeric strings (no leading digits)
        fc.stringOf(
          fc.mapToConstant({ num: 26, build: (v) => String.fromCharCode(65 + v) }),
          { minLength: 1, maxLength: 10 }
        )
      );
    case 'long':
      return fc.oneof(
        // Non-numeric strings (no leading digits)
        fc.stringOf(
          fc.mapToConstant({ num: 26, build: (v) => String.fromCharCode(65 + v) }),
          { minLength: 1, maxLength: 10 }
        ),
        // Empty string
        fc.constant(''),
        // Whitespace only
        fc.constant('   ')
      );
    case 'single':
    case 'double':
      return fc.oneof(
        // Non-numeric strings
        fc.stringOf(
          fc.mapToConstant({ num: 26, build: (v) => String.fromCharCode(65 + v) }),
          { minLength: 1, maxLength: 10 }
        ),
        // Empty string
        fc.constant(''),
        // Infinity / NaN as strings
        fc.constantFrom('Infinity', '-Infinity', 'NaN')
      );
    case 'boolean':
      return fc.oneof(
        // Integers other than 0 and 1
        fc.integer({ min: 2, max: 100 }).map(String),
        fc.integer({ min: -100, max: -1 }).map(String),
        // Non-numeric strings
        fc.stringOf(
          fc.mapToConstant({ num: 26, build: (v) => String.fromCharCode(65 + v) }),
          { minLength: 1, maxLength: 10 }
        ),
        fc.constantFrom('true', 'false', 'yes', 'no', '')
      );
    default:
      // string type always valid — should not be called
      return fc.constant('');
  }
}

// ---------------------------------------------------------------------------
// Helper: build a single-key document + schema for type validation
// ---------------------------------------------------------------------------

/**
 * Creates a minimal IniDocument and ConfigSchema with one section and one key,
 * then runs validateTypes and returns the diagnostics.
 */
function runTypeValidation(type, value, opts = {}) {
  const sectionName = 'TESTSECTION';
  const keyName = 'TESTKEY';

  const entry = createIniEntry(keyName, value, 1);
  const section = createIniSection(sectionName, 1, [entry]);
  const document = createIniDocument([section], []);

  const keyDef = { type, required: true };
  if (opts.min !== undefined) keyDef.min = opts.min;
  if (opts.max !== undefined) keyDef.max = opts.max;

  const schema = createConfigSchema('test.ini', 'test', {
    [sectionName]: {
      required: true,
      strict: true,
      keys: { [keyName]: keyDef },
    },
  }, []);

  return validateTypes(document, schema, 'test.ini');
}

// ---------------------------------------------------------------------------
// Property Tests
// ---------------------------------------------------------------------------

describe('Type Validator Property Tests', () => {
  // Feature: config-file-validation, Property 3: Type validation soundness
  // **Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5**
  it('Property 3: Type Validation Soundness — valid values produce no diagnostics', () => {
    // For each type, generate valid values and assert no type error diagnostics
    const arbTypeAndValidValue = arbType
      .filter((t) => t !== 'string') // string always valid, tested separately
      .chain((type) => fc.tuple(fc.constant(type), arbValidValueForType(type)));

    fc.assert(
      fc.property(arbTypeAndValidValue, ([type, value]) => {
        const diagnostics = runTypeValidation(type, value);
        const typeErrors = diagnostics.filter(
          (d) => d.severity === 'error' && d.message.includes('Expected type')
        );
        assert.equal(
          typeErrors.length,
          0,
          `Type '${type}' should accept '${value}' but got diagnostic: ${typeErrors.map((d) => d.message).join('; ')}`
        );
      }),
      { numRuns: 100 }
    );
  });

  it('Property 3: Type Validation Soundness — invalid values produce diagnostics', () => {
    // For each non-string type, generate invalid values and assert a type error diagnostic
    const arbTypeAndInvalidValue = fc
      .constantFrom('integer', 'long', 'single', 'double', 'boolean')
      .chain((type) => fc.tuple(fc.constant(type), arbInvalidValueForType(type)));

    fc.assert(
      fc.property(arbTypeAndInvalidValue, ([type, value]) => {
        const diagnostics = runTypeValidation(type, value);
        const typeErrors = diagnostics.filter(
          (d) => d.severity === 'error' && d.message.includes('Expected type')
        );
        assert.equal(
          typeErrors.length,
          1,
          `Type '${type}' should reject '${value}' but got ${typeErrors.length} diagnostic(s)`
        );
      }),
      { numRuns: 100 }
    );
  });

  it('Property 3: Type Validation Soundness — string type always accepts', () => {
    fc.assert(
      fc.property(fc.string({ minLength: 0, maxLength: 50 }), (value) => {
        const diagnostics = runTypeValidation('string', value);
        assert.equal(
          diagnostics.length,
          0,
          `String type should accept any value but got diagnostics for '${value}'`
        );
      }),
      { numRuns: 100 }
    );
  });

  it('Property 3: Type Validation Soundness — acceptance matches oracle', () => {
    // Combined property: for any (type, value) pair, the validator's acceptance
    // matches the oracle function
    const arbTypeAndAnyValue = arbType.chain((type) =>
      fc.tuple(
        fc.constant(type),
        fc.oneof(
          arbValidValueForType(type),
          // Also generate some random strings
          fc.stringOf(
            fc.integer({ min: 32, max: 126 }).map((c) => String.fromCharCode(c)),
            { minLength: 0, maxLength: 20 }
          )
        )
      )
    );

    fc.assert(
      fc.property(arbTypeAndAnyValue, ([type, value]) => {
        const diagnostics = runTypeValidation(type, value);
        const typeErrors = diagnostics.filter(
          (d) => d.severity === 'error' && d.message.includes('Expected type')
        );
        const validatorAccepts = typeErrors.length === 0;
        const oracleAccepts = isValidForType(type, value);

        assert.equal(
          validatorAccepts,
          oracleAccepts,
          `Type '${type}', value '${value}': validator ${validatorAccepts ? 'accepted' : 'rejected'} but oracle ${oracleAccepts ? 'accepted' : 'rejected'}`
        );
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 4: Bounds validation
  // **Validates: Requirements 4.7**
  it('Property 4: Bounds Validation — diagnostic iff value out of bounds', () => {
    // Generate a numeric type, a valid value for that type, and min/max bounds.
    // Assert a bounds diagnostic is produced iff the value is outside [min, max].
    const arbNumericType = fc.constantFrom('integer', 'long');

    const arbBoundsScenario = arbNumericType.chain((type) => {
      const range = TYPE_RANGES[type];
      // Generate a valid value within the type's range
      const arbValue = fc.integer({ min: range.min, max: range.max });
      // Generate min and max bounds within the type's range
      const arbMin = fc.integer({ min: range.min, max: range.max });
      const arbMax = fc.integer({ min: range.min, max: range.max });

      return fc.tuple(
        fc.constant(type),
        arbValue,
        arbMin,
        arbMax
      ).map(([t, val, rawMin, rawMax]) => {
        // Ensure min <= max for the bounds
        const min = Math.min(rawMin, rawMax);
        const max = Math.max(rawMin, rawMax);
        return { type: t, value: val, min, max };
      });
    });

    fc.assert(
      fc.property(arbBoundsScenario, ({ type, value, min, max }) => {
        const diagnostics = runTypeValidation(type, String(value), { min, max });

        // Filter for bounds-related diagnostics (below minimum / above maximum)
        const boundsDiags = diagnostics.filter(
          (d) =>
            d.severity === 'error' &&
            (d.message.includes('below minimum') || d.message.includes('above maximum'))
        );

        const outOfBounds = value < min || value > max;

        if (outOfBounds) {
          assert.equal(
            boundsDiags.length,
            1,
            `Value ${value} is outside [${min}, ${max}] for type '${type}' — expected 1 bounds diagnostic, got ${boundsDiags.length}`
          );
        } else {
          assert.equal(
            boundsDiags.length,
            0,
            `Value ${value} is within [${min}, ${max}] for type '${type}' — expected 0 bounds diagnostics, got ${boundsDiags.length}`
          );
        }
      }),
      { numRuns: 100 }
    );
  });

  it('Property 4: Bounds Validation — no bounds diagnostics when bounds not specified', () => {
    // When no min/max is in the schema, no bounds diagnostics should appear
    const arbNumericType = fc.constantFrom('integer', 'long');

    const arbNoBoundsScenario = arbNumericType.chain((type) => {
      const range = TYPE_RANGES[type];
      return fc.tuple(
        fc.constant(type),
        fc.integer({ min: range.min, max: range.max }).map(String)
      );
    });

    fc.assert(
      fc.property(arbNoBoundsScenario, ([type, value]) => {
        const diagnostics = runTypeValidation(type, value);
        const boundsDiags = diagnostics.filter(
          (d) =>
            d.severity === 'error' &&
            (d.message.includes('below minimum') || d.message.includes('above maximum'))
        );
        assert.equal(
          boundsDiags.length,
          0,
          `No bounds in schema, but got bounds diagnostic for type '${type}', value '${value}'`
        );
      }),
      { numRuns: 100 }
    );
  });
});
