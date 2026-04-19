/**
 * Unit tests for VB6 Source Parser.
 *
 * Uses Node.js built-in test runner (node:test) and node:assert/strict.
 * Tests extractSchemaEntries and mergeSchemas with real VB6 patterns
 * from ServerConfig.cls and FileIO.bas.
 *
 * Requirements: 9.1–9.8, 9.14
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { extractSchemaEntries, mergeSchemas } from '../lib/vb6-parser.js';
import { createConfigSchema } from '../lib/data-models.js';

describe('extractSchemaEntries', () => {

  // 1. CLng typed mSettings.Add with default (real line from ServerConfig.cls)
  it('extracts CLng typed mSettings.Add with default value', () => {
    const line = '    mSettings.Add "DP_BuffersPerThread", CLng(val(reader.GetValue("DIRECTPLAY", "BuffersPerThread", 5)))';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].section, 'DIRECTPLAY');
    assert.equal(entries[0].key, 'BUFFERSPERTHREAD');
    assert.equal(entries[0].type, 'long');
    assert.equal(entries[0].defaultValue, '5');
    assert.equal(entries[0].required, false);
  });

  // 2. CInt wrapper (real line from ServerConfig.cls)
  it('extracts CInt typed mSettings.Add', () => {
    const line = '    mSettings.Add "ExpMult", CInt(val(reader.GetValue("CONFIGURACIONES", "ExpMult")))';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].section, 'CONFIGURACIONES');
    assert.equal(entries[0].key, 'EXPMULT');
    assert.equal(entries[0].type, 'integer');
    assert.equal(entries[0].required, true);
    assert.equal(entries[0].defaultValue, null);
  });

  // 3. CSng wrapper (real line from ServerConfig.cls)
  it('extracts CSng typed mSettings.Add with string default', () => {
    const line = '    mSettings.Add "PenaltyExpUserPerLevel", CSng(val(reader.GetValue("CONFIGURACIONES", "PenaltyExpUserPerLevel", "0.05")))';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].section, 'CONFIGURACIONES');
    assert.equal(entries[0].key, 'PENALTYEXPUSERPERLEVEL');
    assert.equal(entries[0].type, 'single');
    assert.equal(entries[0].defaultValue, '"0.05"');
    assert.equal(entries[0].required, false);
  });

  // 4. CDbl wrapper (real line from ServerConfig.cls)
  it('extracts CDbl typed mSettings.Add', () => {
    const line = '    mSettings.Add "RecoleccionMult", CDbl(val(reader.GetValue("CONFIGURACIONES", "RecoleccionMult")))';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].section, 'CONFIGURACIONES');
    assert.equal(entries[0].key, 'RECOLECCIONMULT');
    assert.equal(entries[0].type, 'double');
    assert.equal(entries[0].required, true);
  });

  // 5. CByte wrapper
  it('extracts CByte typed mSettings.Add as integer', () => {
    const line = '    mSettings.Add "SomeByte", CByte(val(reader.GetValue("INIT", "ByteVal", 0)))';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].section, 'INIT');
    assert.equal(entries[0].key, 'BYTEVAL');
    assert.equal(entries[0].type, 'integer');
    assert.equal(entries[0].defaultValue, '0');
    assert.equal(entries[0].required, false);
  });

  // 6. String type — no wrapper (real line from ServerConfig.cls)
  it('extracts string mSettings.Add with no type wrapper', () => {
    const line = '    mSettings.Add "AUTOCAPTURETHEFLAG_Description", reader.GetValue("AUTOCAPTURETHEFLAG", "Description")';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].section, 'AUTOCAPTURETHEFLAG');
    assert.equal(entries[0].key, 'DESCRIPTION');
    assert.equal(entries[0].type, 'string');
    assert.equal(entries[0].required, true);
    assert.equal(entries[0].defaultValue, null);
  });

  // 7. No default ? required (real line from ServerConfig.cls)
  it('marks key as required when no default is provided', () => {
    const line = '    mSettings.Add "OroPorNivelBilletera", CLng(val(reader.GetValue("CONFIGURACIONES", "OroPorNivelBilletera")))';
    const entries = extractSchemaEntries(line, 'ServerConfig.cls');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].required, true);
    assert.equal(entries[0].defaultValue, null);
  });

  // 8. GetVar extraction (real line from FileIO.bas)
  it('extracts GetVar call with DatPath concatenation', () => {
    const line = '    n = val(GetVar(DatPath & "npcs.dat", "INIT", "NumNPCs"))';
    const entries = extractSchemaEntries(line, 'FileIO.bas');
    assert.equal(entries.length, 1);
    assert.equal(entries[0].file, 'npcs.dat');
    assert.equal(entries[0].section, 'INIT');
    assert.equal(entries[0].key, 'NUMNPCS');
  });

  // 9. Multiple lines ? multiple entries
  it('extracts multiple entries from multiple lines', () => {
    const content = [
      '    mSettings.Add "DP_BuffersPerThread", CLng(val(reader.GetValue("DIRECTPLAY", "BuffersPerThread", 5)))',
      '    mSettings.Add "ExpMult", CInt(val(reader.GetValue("CONFIGURACIONES", "ExpMult")))',
      '    mSettings.Add "AUTOCAPTURETHEFLAG_Description", reader.GetValue("AUTOCAPTURETHEFLAG", "Description")',
    ].join('\n');
    const entries = extractSchemaEntries(content, 'ServerConfig.cls');
    assert.equal(entries.length, 3);
    assert.equal(entries[0].key, 'BUFFERSPERTHREAD');
    assert.equal(entries[1].key, 'EXPMULT');
    assert.equal(entries[2].key, 'DESCRIPTION');
  });

  // 10. Non-matching lines ? empty array
  it('returns empty array for non-matching lines', () => {
    const content = [
      "' This is a comment",
      'Dim reader As clsIniManager',
      'Set reader = New clsIniManager',
      'Call reader.Initialize(Filename)',
      'If mSettings.Exists(key) Then',
    ].join('\n');
    const entries = extractSchemaEntries(content, 'ServerConfig.cls');
    assert.equal(entries.length, 0);
  });
});


describe('mergeSchemas', () => {

  // 11. Null existing ? returns generated as-is
  it('returns generated schema when existing is null', () => {
    const generated = createConfigSchema('Server.ini', 'Generated', {
      INIT: {
        required: true,
        strict: true,
        keys: {
          PORT: { type: 'long', required: true },
        },
      },
    }, []);
    const result = mergeSchemas(generated, null);
    assert.deepStrictEqual(result, generated);
  });

  // 12. Preserves min/max from existing
  it('preserves min and max bounds from existing schema', () => {
    const generated = createConfigSchema('Server.ini', 'Generated', {
      INIT: {
        required: true,
        strict: true,
        keys: {
          PORT: { type: 'long', required: true },
        },
      },
    }, []);
    const existing = createConfigSchema('Server.ini', 'Existing desc', {
      INIT: {
        required: true,
        strict: true,
        keys: {
          PORT: { type: 'long', required: true, min: 1, max: 65535 },
        },
      },
    }, []);
    const merged = mergeSchemas(generated, existing);
    assert.equal(merged.sections.INIT.keys.PORT.min, 1);
    assert.equal(merged.sections.INIT.keys.PORT.max, 65535);
    assert.equal(merged.sections.INIT.keys.PORT.type, 'long');
    assert.equal(merged.sections.INIT.keys.PORT.required, true);
  });

  // 13. Preserves rules from existing
  it('preserves rules array from existing schema', () => {
    const rules = [
      { type: 'min_max', section: 'INIT', minKey: 'MINDADOS', maxKey: 'MAXDADOS' },
    ];
    const generated = createConfigSchema('Server.ini', 'Generated', {
      INIT: { required: true, strict: true, keys: {} },
    }, []);
    const existing = createConfigSchema('Server.ini', 'Existing', {
      INIT: { required: true, strict: true, keys: {} },
    }, rules);
    const merged = mergeSchemas(generated, existing);
    assert.deepStrictEqual(merged.rules, rules);
  });

  // 14. Preserves strict flag from existing
  it('preserves strict flag from existing schema section', () => {
    const generated = createConfigSchema('Server.ini', '', {
      ADMINES: { required: false, strict: true, keys: {} },
    }, []);
    const existing = createConfigSchema('Server.ini', '', {
      ADMINES: { required: false, strict: false, keys: {} },
    }, []);
    const merged = mergeSchemas(generated, existing);
    assert.equal(merged.sections.ADMINES.strict, false);
  });

  // 15. Preserves description from existing
  it('preserves description from existing schema', () => {
    const generated = createConfigSchema('Server.ini', 'Auto-generated', {
      INIT: { required: true, strict: true, keys: {} },
    }, []);
    const existing = createConfigSchema('Server.ini', 'Main server configuration', {
      INIT: { required: true, strict: true, keys: {} },
    }, []);
    const merged = mergeSchemas(generated, existing);
    assert.equal(merged.description, 'Main server configuration');
  });
});
