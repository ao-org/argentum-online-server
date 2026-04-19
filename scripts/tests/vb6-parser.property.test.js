/**
 * Property-based tests for VB6 Source Parser and Schema Merger.
 *
 * Uses fast-check with Node.js built-in test runner.
 * Each property references its design document property via a tag comment.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import fc from 'fast-check';
import { extractSchemaEntries, mergeSchemas } from '../lib/vb6-parser.js';
import { createConfigSchema } from '../lib/data-models.js';

// ---------------------------------------------------------------------------
// Custom Arbitraries
// ---------------------------------------------------------------------------

/**
 * arbUpperAlphaNum: Generates uppercase alphanumeric strings (1-15 chars).
 * Used for section names and keys.
 */
const arbUpperAlphaNum = fc.stringOf(
  fc.mapToConstant(
    { num: 26, build: (v) => String.fromCharCode(65 + v) }, // A-Z
    { num: 10, build: (v) => String.fromCharCode(48 + v) }  // 0-9
  ),
  { minLength: 1, maxLength: 15 }
);

/**
 * VB6 type wrappers and their expected schema types.
 */
const VB6_TYPED_WRAPPERS = [
  { fn: 'CLng', schemaType: 'long' },
  { fn: 'CInt', schemaType: 'integer' },
  { fn: 'CDbl', schemaType: 'double' },
  { fn: 'CSng', schemaType: 'single' },
  { fn: 'CByte', schemaType: 'integer' },
];

/**
 * arbVb6TypeWrapper: Picks a random VB6 type conversion function.
 */
const arbVb6TypeWrapper = fc.constantFrom(...VB6_TYPED_WRAPPERS);

/**
 * arbDefaultValue: Generates a simple default value (numeric string).
 */
const arbDefaultValue = fc.integer({ min: 0, max: 9999 }).map(String);

/**
 * arbSettingName: Generates a setting name for mSettings.Add (e.g., "DP_SomeKey").
 */
const arbSettingName = arbUpperAlphaNum.map((s) => `SET_${s}`);


/**
 * arbVb6SettingsLine: Generates mSettings.Add lines with typed wrappers.
 *
 * Format: mSettings.Add "SETTINGNAME", CLng(val(reader.GetValue("SECTION", "KEY", DEFAULT)))
 * or without default: mSettings.Add "SETTINGNAME", CLng(val(reader.GetValue("SECTION", "KEY")))
 *
 * Returns { line, section, key, schemaType, defaultValue, hasDefault }
 */
const arbVb6TypedSettingsLine = fc.tuple(
  arbSettingName,
  arbVb6TypeWrapper,
  arbUpperAlphaNum,
  arbUpperAlphaNum,
  fc.option(arbDefaultValue, { nil: undefined })
).map(([settingName, wrapper, section, key, defaultVal]) => {
  const hasDefault = defaultVal !== undefined;
  const getValueArgs = hasDefault
    ? `"${section}", "${key}", ${defaultVal}`
    : `"${section}", "${key}"`;
  const line = `mSettings.Add "${settingName}", ${wrapper.fn}(val(reader.GetValue(${getValueArgs})))`;
  return {
    line,
    section: section.toUpperCase(),
    key: key.toUpperCase(),
    schemaType: wrapper.schemaType,
    defaultValue: hasDefault ? defaultVal : null,
    hasDefault,
  };
});

/**
 * arbVb6StringSettingsLine: Generates mSettings.Add lines WITHOUT type wrapper (string type).
 *
 * Format: mSettings.Add "SETTINGNAME", reader.GetValue("SECTION", "KEY")
 *
 * Returns { line, section, key, schemaType: 'string' }
 */
const arbVb6StringSettingsLine = fc.tuple(
  arbSettingName,
  arbUpperAlphaNum,
  arbUpperAlphaNum
).map(([settingName, section, key]) => {
  const line = `mSettings.Add "${settingName}", reader.GetValue("${section}", "${key}")`;
  return {
    line,
    section: section.toUpperCase(),
    key: key.toUpperCase(),
    schemaType: 'string',
    defaultValue: null,
    hasDefault: false,
  };
});

/**
 * arbVb6SettingsLine: Union of typed and string settings lines.
 */
const arbVb6SettingsLine = fc.oneof(arbVb6TypedSettingsLine, arbVb6StringSettingsLine);

/**
 * arbFileName: Generates a simple filename for GetVar paths.
 */
const arbFileName = fc.stringOf(
  fc.mapToConstant(
    { num: 26, build: (v) => String.fromCharCode(97 + v) } // a-z
  ),
  { minLength: 1, maxLength: 10 }
).map((s) => `${s}.ini`);

/**
 * arbVb6GetVarLine: Generates GetVar(App.Path & "\filename.ini", "SECTION", "KEY") lines.
 *
 * Returns { line, section, key, fileName }
 */
const arbVb6GetVarLine = fc.tuple(
  arbFileName,
  arbUpperAlphaNum,
  arbUpperAlphaNum
).map(([fileName, section, key]) => {
  const line = `GetVar(App.Path & "\\${fileName}", "${section}", "${key}")`;
  return {
    line,
    section: section.toUpperCase(),
    key: key.toUpperCase(),
    fileName,
  };
});

/**
 * arbAnnotations: Generates manual annotation fields for schema keys.
 */
const arbAnnotations = fc.record({
  min: fc.option(fc.integer({ min: -1000, max: 0 }), { nil: undefined }),
  max: fc.option(fc.integer({ min: 1, max: 10000 }), { nil: undefined }),
  pattern: fc.option(fc.constant('^[A-Z]+$'), { nil: undefined }),
  description: fc.option(fc.constant('A manual description'), { nil: undefined }),
}).map((rec) => {
  // Filter out undefined values
  const result = {};
  if (rec.min !== undefined) result.min = rec.min;
  if (rec.max !== undefined) result.max = rec.max;
  if (rec.pattern !== undefined) result.pattern = rec.pattern;
  if (rec.description !== undefined) result.description = rec.description;
  return result;
});

/**
 * arbSchemaWithAnnotations: Generates a ConfigSchema with manual annotations on keys.
 *
 * Returns { existing, generated, annotatedKeys }
 * - existing: schema with manual annotations
 * - generated: auto-generated schema (same structure, no annotations)
 * - annotatedKeys: map of sectionName -> keyName -> annotations
 */
const arbSchemaWithAnnotations = fc.tuple(
  arbUpperAlphaNum,
  fc.array(
    fc.tuple(
      arbUpperAlphaNum,
      fc.array(
        fc.tuple(
          arbUpperAlphaNum,
          fc.constantFrom('integer', 'long', 'string', 'double'),
          arbAnnotations
        ),
        { minLength: 1, maxLength: 4 }
      )
    ),
    { minLength: 1, maxLength: 3 }
  ),
  fc.option(
    fc.array(
      fc.record({
        type: fc.constant('min_max'),
        section: arbUpperAlphaNum,
        minKey: fc.constant('MINVAL'),
        maxKey: fc.constant('MAXVAL'),
      }),
      { minLength: 1, maxLength: 2 }
    ),
    { nil: undefined }
  )
).map(([fileName, sectionDefs, rules]) => {
  const existingSections = {};
  const generatedSections = {};
  const annotatedKeys = {};

  for (const [sectionName, keyDefs] of sectionDefs) {
    const existingKeys = {};
    const generatedKeys = {};
    const sectionAnnotations = {};
    const seenKeys = new Set();

    for (const [keyName, type, annotations] of keyDefs) {
      if (seenKeys.has(keyName)) continue;
      seenKeys.add(keyName);

      // Generated key: type + required, no annotations
      generatedKeys[keyName] = { type, required: true };

      // Existing key: type + required + manual annotations
      existingKeys[keyName] = { type, required: true, ...annotations };

      if (Object.keys(annotations).length > 0) {
        sectionAnnotations[keyName] = annotations;
      }
    }

    existingSections[sectionName] = { required: true, strict: true, keys: existingKeys };
    generatedSections[sectionName] = { required: true, strict: false, keys: generatedKeys };

    if (Object.keys(sectionAnnotations).length > 0) {
      annotatedKeys[sectionName] = sectionAnnotations;
    }
  }

  const existingRules = rules !== undefined ? rules : [];

  const existing = createConfigSchema(`${fileName}.ini`, 'Manual description', existingSections, existingRules);
  const generated = createConfigSchema(`${fileName}.ini`, '', generatedSections, []);

  return { existing, generated, annotatedKeys, existingRules };
});


// ---------------------------------------------------------------------------
// Property Tests
// ---------------------------------------------------------------------------

describe('VB6 Parser Property Tests', () => {
  // Feature: config-file-validation, Property 13: VB6 mSettings.Add extraction
  // **Validates: Requirements 9.2, 9.3, 9.6, 9.7**
  it('Property 13: VB6 mSettings.Add Extraction', () => {
    fc.assert(
      fc.property(arbVb6SettingsLine, ({ line, section, key, schemaType, defaultValue, hasDefault }) => {
        const entries = extractSchemaEntries(line, 'TestFile.cls');

        assert.equal(entries.length, 1,
          `Expected 1 entry from line: ${line}, got ${entries.length}`);

        const entry = entries[0];

        // Correct section name (uppercase)
        assert.equal(entry.section, section,
          `Section mismatch: expected '${section}', got '${entry.section}'`);

        // Correct key name (uppercase)
        assert.equal(entry.key, key,
          `Key mismatch: expected '${key}', got '${entry.key}'`);

        // Correct schema type
        assert.equal(entry.type, schemaType,
          `Type mismatch: expected '${schemaType}', got '${entry.type}'`);

        // Default value and required flag
        if (hasDefault) {
          assert.equal(entry.defaultValue, defaultValue,
            `Default mismatch: expected '${defaultValue}', got '${entry.defaultValue}'`);
          assert.equal(entry.required, false,
            'Entry with default should not be required');
        } else {
          assert.equal(entry.defaultValue, null,
            `Expected null default, got '${entry.defaultValue}'`);
          assert.equal(entry.required, true,
            'Entry without default should be required');
        }

        // Source should reference the file
        assert.ok(entry.source.startsWith('TestFile.cls:'),
          `Source should start with 'TestFile.cls:', got '${entry.source}'`);
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 14: VB6 GetVar extraction
  // **Validates: Requirements 9.4**
  it('Property 14: VB6 GetVar Extraction', () => {
    fc.assert(
      fc.property(arbVb6GetVarLine, ({ line, section, key, fileName }) => {
        const entries = extractSchemaEntries(line, 'TestFile.bas');

        assert.equal(entries.length, 1,
          `Expected 1 entry from line: ${line}, got ${entries.length}`);

        const entry = entries[0];

        // Correct section name (uppercase)
        assert.equal(entry.section, section,
          `Section mismatch: expected '${section}', got '${entry.section}'`);

        // Correct key name (uppercase)
        assert.equal(entry.key, key,
          `Key mismatch: expected '${key}', got '${entry.key}'`);

        // File should be resolved from the path expression
        assert.equal(entry.file, fileName,
          `File mismatch: expected '${fileName}', got '${entry.file}'`);

        // Source should reference the file
        assert.ok(entry.source.startsWith('TestFile.bas:'),
          `Source should start with 'TestFile.bas:', got '${entry.source}'`);
      }),
      { numRuns: 100 }
    );
  });

  // Feature: config-file-validation, Property 15: Schema merge preserves manual annotations
  // **Validates: Requirements 9.14**
  it('Property 15: Schema Merge Preserves Manual Annotations', () => {
    fc.assert(
      fc.property(arbSchemaWithAnnotations, ({ existing, generated, annotatedKeys, existingRules }) => {
        const merged = mergeSchemas(generated, existing);

        // All generated sections and keys should be present
        for (const [sectionName, genSection] of Object.entries(generated.sections)) {
          assert.ok(merged.sections[sectionName],
            `Merged schema missing generated section '${sectionName}'`);

          for (const keyName of Object.keys(genSection.keys || {})) {
            assert.ok(merged.sections[sectionName].keys[keyName],
              `Merged schema missing generated key '${sectionName}.${keyName}'`);
          }
        }

        // Manual annotations from existing should be preserved
        for (const [sectionName, keys] of Object.entries(annotatedKeys)) {
          for (const [keyName, annotations] of Object.entries(keys)) {
            const mergedKey = merged.sections[sectionName]?.keys[keyName];
            assert.ok(mergedKey,
              `Merged key '${sectionName}.${keyName}' not found`);

            for (const [annotationName, annotationValue] of Object.entries(annotations)) {
              assert.equal(mergedKey[annotationName], annotationValue,
                `Annotation '${annotationName}' not preserved on '${sectionName}.${keyName}': expected ${annotationValue}, got ${mergedKey[annotationName]}`);
            }
          }
        }

        // Existing description should be preserved
        assert.equal(merged.description, existing.description,
          `Description not preserved: expected '${existing.description}', got '${merged.description}'`);

        // Existing rules should be preserved
        if (existingRules.length > 0) {
          assert.deepStrictEqual(merged.rules, existingRules,
            'Rules from existing schema not preserved');
        }

        // Existing strict flags should be preserved
        for (const [sectionName, existSection] of Object.entries(existing.sections)) {
          if (merged.sections[sectionName]) {
            assert.equal(merged.sections[sectionName].strict, existSection.strict,
              `Strict flag not preserved on section '${sectionName}'`);
          }
        }
      }),
      { numRuns: 100 }
    );
  });
});
