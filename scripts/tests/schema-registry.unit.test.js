/**
 * Unit tests for schema-registry.js
 *
 * Tests: loadSchemas — directory reading, JSON parsing, key indexing, error handling.
 * Requirements: 2.7, 8.1, 8.2
 */

import { describe, it, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, mkdirSync, rmSync, chmodSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';
import { loadSchemas } from '../lib/schema-registry.js';

describe('loadSchemas', () => {
  let tempDir;

  beforeEach(() => {
    tempDir = mkdtempSync(join(tmpdir(), 'schema-test-'));
  });

  afterEach(() => {
    rmSync(tempDir, { recursive: true, force: true });
  });

  it('returns empty map and no diagnostics for empty directory', () => {
    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(schemas.size, 0);
    assert.equal(diagnostics.length, 0);
  });

  it('loads a valid schema file and indexes by uppercase file property', () => {
    const schema = { file: 'Server.ini', description: 'test', sections: {}, rules: [] };
    writeFileSync(join(tempDir, 'Server.ini.json'), JSON.stringify(schema));

    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(diagnostics.length, 0);
    assert.equal(schemas.size, 1);
    assert.ok(schemas.has('SERVER.INI'));
    assert.deepStrictEqual(schemas.get('SERVER.INI'), schema);
  });

  it('falls back to filename-derived key when schema has no file property', () => {
    const schema = { description: 'no file prop', sections: {}, rules: [] };
    writeFileSync(join(tempDir, 'Configuracion.ini.json'), JSON.stringify(schema));

    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(diagnostics.length, 0);
    assert.equal(schemas.size, 1);
    assert.ok(schemas.has('CONFIGURACION.INI'));
  });

  it('loads multiple schema files', () => {
    writeFileSync(join(tempDir, 'Server.ini.json'),
      JSON.stringify({ file: 'Server.ini', sections: {} }));
    writeFileSync(join(tempDir, 'feature_toggle.ini.json'),
      JSON.stringify({ file: 'feature_toggle.ini', sections: {} }));

    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(diagnostics.length, 0);
    assert.equal(schemas.size, 2);
    assert.ok(schemas.has('SERVER.INI'));
    assert.ok(schemas.has('FEATURE_TOGGLE.INI'));
  });

  it('ignores non-json files', () => {
    writeFileSync(join(tempDir, 'README.md'), '# Schemas');
    writeFileSync(join(tempDir, 'Server.ini.json'),
      JSON.stringify({ file: 'Server.ini', sections: {} }));

    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(diagnostics.length, 0);
    assert.equal(schemas.size, 1);
  });

  it('produces error diagnostic for malformed JSON', () => {
    writeFileSync(join(tempDir, 'bad.json'), '{ not valid json }');

    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(schemas.size, 0);
    assert.equal(diagnostics.length, 1);
    assert.equal(diagnostics[0].severity, 'error');
    assert.ok(diagnostics[0].message.includes('Malformed JSON'));
  });

  it('produces error diagnostic for non-existent directory', () => {
    const { schemas, diagnostics } = loadSchemas('/nonexistent/path/to/schemas');
    assert.equal(schemas.size, 0);
    assert.equal(diagnostics.length, 1);
    assert.equal(diagnostics[0].severity, 'error');
    assert.ok(diagnostics[0].message.includes('Cannot read schema directory'));
  });

  it('continues loading other files when one has malformed JSON', () => {
    writeFileSync(join(tempDir, 'bad.json'), '{ broken');
    writeFileSync(join(tempDir, 'Server.ini.json'),
      JSON.stringify({ file: 'Server.ini', sections: {} }));

    const { schemas, diagnostics } = loadSchemas(tempDir);
    assert.equal(schemas.size, 1);
    assert.ok(schemas.has('SERVER.INI'));
    assert.equal(diagnostics.length, 1);
    assert.equal(diagnostics[0].severity, 'error');
  });

  it('uses schema file property over filename for key', () => {
    // Schema file property differs from filename
    const schema = { file: 'Custom.dat', sections: {} };
    writeFileSync(join(tempDir, 'something-else.json'), JSON.stringify(schema));

    const { schemas } = loadSchemas(tempDir);
    assert.equal(schemas.size, 1);
    assert.ok(schemas.has('CUSTOM.DAT'));
    assert.ok(!schemas.has('SOMETHING-ELSE'));
  });
});
