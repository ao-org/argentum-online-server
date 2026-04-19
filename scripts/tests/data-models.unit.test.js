import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  createIniDocument,
  createIniSection,
  createIniEntry,
  createDiagnostic,
  createConfigSchema,
  createSchemaEntry,
} from '../lib/data-models.js';

describe('data-models', () => {
  describe('createIniDocument', () => {
    it('creates an empty document by default', () => {
      const doc = createIniDocument();
      assert.deepStrictEqual(doc, { sections: [], orphans: [] });
    });

    it('creates a document with sections and orphans', () => {
      const entry = createIniEntry('KEY', 'val', 2);
      const section = createIniSection('INIT', 1, [entry]);
      const orphan = createIniEntry('STRAY', 'x', 0);
      const doc = createIniDocument([section], [orphan]);
      assert.equal(doc.sections.length, 1);
      assert.equal(doc.orphans.length, 1);
      assert.equal(doc.sections[0].name, 'INIT');
    });
  });

  describe('createIniSection', () => {
    it('creates a section with defaults', () => {
      const s = createIniSection('SETTINGS', 5);
      assert.equal(s.name, 'SETTINGS');
      assert.equal(s.line, 5);
      assert.deepStrictEqual(s.entries, []);
    });
  });

  describe('createIniEntry', () => {
    it('creates an entry', () => {
      const e = createIniEntry('PORT', '7667', 10);
      assert.equal(e.key, 'PORT');
      assert.equal(e.value, '7667');
      assert.equal(e.line, 10);
    });
  });

  describe('createDiagnostic', () => {
    it('creates a diagnostic with all fields', () => {
      const d = createDiagnostic('error', 'Server.ini', 42, 'INIT', 'PORT', 'Invalid value');
      assert.equal(d.severity, 'error');
      assert.equal(d.file, 'Server.ini');
      assert.equal(d.line, 42);
      assert.equal(d.section, 'INIT');
      assert.equal(d.key, 'PORT');
      assert.equal(d.message, 'Invalid value');
    });

    it('allows null section and key', () => {
      const d = createDiagnostic('warning', 'test.ini', 0, null, null, 'General warning');
      assert.equal(d.section, null);
      assert.equal(d.key, null);
    });
  });

  describe('createConfigSchema', () => {
    it('creates a schema with defaults', () => {
      const s = createConfigSchema('Server.ini');
      assert.equal(s.file, 'Server.ini');
      assert.equal(s.description, '');
      assert.deepStrictEqual(s.sections, {});
      assert.deepStrictEqual(s.rules, []);
    });

    it('creates a schema with sections and rules', () => {
      const sections = {
        INIT: { required: true, strict: true, keys: { PORT: { type: 'long', required: true } } },
      };
      const rules = [{ type: 'min_max', section: 'INIT', minKey: 'MIN', maxKey: 'MAX' }];
      const s = createConfigSchema('Server.ini', 'Main config', sections, rules);
      assert.equal(s.description, 'Main config');
      assert.equal(s.sections.INIT.required, true);
      assert.equal(s.rules.length, 1);
    });
  });

  describe('createSchemaEntry', () => {
    it('creates a schema entry', () => {
      const e = createSchemaEntry('Server.ini', 'INIT', 'PORT', 'long', true, null, 'ServerConfig.cls:42');
      assert.equal(e.file, 'Server.ini');
      assert.equal(e.section, 'INIT');
      assert.equal(e.key, 'PORT');
      assert.equal(e.type, 'long');
      assert.equal(e.required, true);
      assert.equal(e.defaultValue, null);
      assert.equal(e.source, 'ServerConfig.cls:42');
    });

    it('creates an optional entry with default', () => {
      const e = createSchemaEntry('Server.ini', 'INIT', 'HIDE', 'boolean', false, '0', 'ServerConfig.cls:50');
      assert.equal(e.required, false);
      assert.equal(e.defaultValue, '0');
    });
  });
});
