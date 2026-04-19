/**
 * Unit tests for validation-runner.js
 *
 * Tests: runValidation — file reading, UTF-8/Latin-1 fallback, schema lookup,
 * validator orchestration, exit code determination, error handling.
 * Requirements: 6.4, 6.5, 8.1, 8.2, 8.3
 */

import { describe, it, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, mkdirSync, rmSync, chmodSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';
import { runValidation } from '../lib/validation-runner.js';

describe('runValidation', () => {
  let tempDir;
  let schemaDir;
  let configDir;

  beforeEach(() => {
    tempDir = mkdtempSync(join(tmpdir(), 'valrunner-test-'));
    schemaDir = join(tempDir, 'schemas');
    configDir = join(tempDir, 'configs');
    mkdirSync(schemaDir);
    mkdirSync(configDir);
  });

  afterEach(() => {
    rmSync(tempDir, { recursive: true, force: true });
  });

  it('returns empty diagnostics and exit code 0 for no config files', () => {
    const result = runValidation({ schemaDir, configFiles: [] });
    assert.deepStrictEqual(result.diagnostics, []);
    assert.equal(result.exitCode, 0);
  });

  it('parses a valid config file with a matching schema and returns exit code 0', () => {
    const schema = {
      file: 'Test.ini',
      sections: {
        INIT: { required: true, strict: true, keys: { PORT: { type: 'integer', required: true } } }
      },
      rules: []
    };
    writeFileSync(join(schemaDir, 'Test.ini.json'), JSON.stringify(schema));

    const configPath = join(configDir, 'Test.ini');
    writeFileSync(configPath, '[INIT]\nPORT=8080\n');

    const result = runValidation({ schemaDir, configFiles: [configPath] });
    assert.equal(result.exitCode, 0);
    // No errors expected for a valid file
    const errors = result.diagnostics.filter(d => d.severity === 'error');
    assert.equal(errors.length, 0);
  });

  it('produces error diagnostic for missing config file and continues', () => {
    const missingPath = join(configDir, 'DoesNotExist.ini');
    const result = runValidation({ schemaDir, configFiles: [missingPath] });

    assert.equal(result.exitCode, 1);
    assert.ok(result.diagnostics.length >= 1);
    const fileError = result.diagnostics.find(d => d.file === missingPath && d.severity === 'error');
    assert.ok(fileError);
    assert.ok(fileError.message.includes('not found'));
  });

  it('produces error diagnostic for permission error and continues', () => {
    // Create a file then make it unreadable
    const configPath = join(configDir, 'Locked.ini');
    writeFileSync(configPath, '[INIT]\nKEY=val\n');

    // Skip on Windows where chmod doesn't restrict reads the same way
    if (process.platform === 'win32') return;

    chmodSync(configPath, 0o000);
    const validPath = join(configDir, 'Other.ini');
    writeFileSync(validPath, '[INIT]\nKEY=val\n');

    const result = runValidation({ schemaDir, configFiles: [configPath, validPath] });
    // Should have an error for the locked file
    const lockError = result.diagnostics.find(d => d.file === configPath && d.severity === 'error');
    assert.ok(lockError);
    assert.ok(lockError.message.includes('Cannot read'));
    // Should still have processed the second file (no crash)
    assert.equal(result.exitCode, 1);

    // Restore permissions for cleanup
    chmodSync(configPath, 0o644);
  });

  it('runs structural, type, and semantic validators when schema exists', () => {
    const schema = {
      file: 'Server.ini',
      sections: {
        INIT: {
          required: true,
          strict: true,
          keys: {
            PORT: { type: 'integer', required: true },
            NAME: { type: 'string', required: true }
          }
        }
      },
      rules: []
    };
    writeFileSync(join(schemaDir, 'Server.ini.json'), JSON.stringify(schema));

    // Missing required key NAME, PORT has invalid type
    const configPath = join(configDir, 'Server.ini');
    writeFileSync(configPath, '[INIT]\nPORT=notanumber\n');

    const result = runValidation({ schemaDir, configFiles: [configPath] });
    assert.equal(result.exitCode, 1);

    // Should have structural error for missing NAME
    const missingKey = result.diagnostics.find(d =>
      d.severity === 'error' && d.key === 'NAME' && d.message.includes('Missing required key')
    );
    assert.ok(missingKey);

    // Should have type error for PORT
    const typeError = result.diagnostics.find(d =>
      d.severity === 'error' && d.key === 'PORT' && d.message.includes('type')
    );
    assert.ok(typeError);
  });

  it('skips validators when no schema matches the config file', () => {
    // No schema files at all
    const configPath = join(configDir, 'Unknown.ini');
    writeFileSync(configPath, '[SECTION]\nKEY=value\n');

    const result = runValidation({ schemaDir, configFiles: [configPath] });
    // No schema means no structural/type/semantic errors, just parse results
    assert.equal(result.exitCode, 0);
  });

  it('exit code is 0 when only warnings are produced', () => {
    const schema = {
      file: 'Test.ini',
      sections: {
        INIT: { required: false, strict: true, keys: {} }
      },
      rules: []
    };
    writeFileSync(join(schemaDir, 'Test.ini.json'), JSON.stringify(schema));

    // Unknown section produces a warning, not an error
    const configPath = join(configDir, 'Test.ini');
    writeFileSync(configPath, '[INIT]\n[UNKNOWN]\nFOO=bar\n');

    const result = runValidation({ schemaDir, configFiles: [configPath] });
    const warnings = result.diagnostics.filter(d => d.severity === 'warning');
    assert.ok(warnings.length > 0, 'Should have at least one warning');
    assert.equal(result.exitCode, 0);
  });

  it('handles Latin-1 encoded files via fallback', () => {
    const schema = {
      file: 'Latin.ini',
      sections: {
        INIT: { required: true, strict: false, keys: { NAME: { type: 'string', required: true } } }
      },
      rules: []
    };
    writeFileSync(join(schemaDir, 'Latin.ini.json'), JSON.stringify(schema));

    // Write a file with Latin-1 bytes that would produce replacement chars in UTF-8
    const configPath = join(configDir, 'Latin.ini');
    const latin1Content = Buffer.from('[INIT]\nNAME=caf\xe9\n', 'latin1');
    writeFileSync(configPath, latin1Content);

    const result = runValidation({ schemaDir, configFiles: [configPath] });
    // Should parse without crashing; the value should contain the Latin-1 character
    const errors = result.diagnostics.filter(d => d.severity === 'error');
    assert.equal(errors.length, 0);
    assert.equal(result.exitCode, 0);
  });

  it('collects diagnostics from multiple config files', () => {
    const schema = {
      file: 'Multi.ini',
      sections: {
        INIT: { required: true, strict: true, keys: { KEY: { type: 'integer', required: true } } }
      },
      rules: []
    };
    writeFileSync(join(schemaDir, 'Multi.ini.json'), JSON.stringify(schema));

    const path1 = join(configDir, 'Multi.ini');
    writeFileSync(path1, '[INIT]\nKEY=abc\n'); // type error

    const missingPath = join(configDir, 'Missing.ini'); // file not found

    const result = runValidation({ schemaDir, configFiles: [path1, missingPath] });
    assert.equal(result.exitCode, 1);
    // Should have diagnostics from both files
    const file1Diags = result.diagnostics.filter(d => d.file === path1);
    const file2Diags = result.diagnostics.filter(d => d.file === missingPath);
    assert.ok(file1Diags.length > 0);
    assert.ok(file2Diags.length > 0);
  });

  it('includes schema loading diagnostics in results', () => {
    const result = runValidation({
      schemaDir: '/nonexistent/schema/dir',
      configFiles: []
    });
    assert.ok(result.diagnostics.length >= 1);
    assert.equal(result.diagnostics[0].severity, 'error');
    assert.ok(result.diagnostics[0].message.includes('Cannot read schema directory'));
    assert.equal(result.exitCode, 1);
  });
});
