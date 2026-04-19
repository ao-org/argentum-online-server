/**
 * Unit tests for INI Parser and Printer.
 *
 * Uses Node.js built-in test runner (node:test) and node:assert/strict.
 * Covers edge cases for parseIni and printIni.
 *
 * Requirements: 1.1–1.10
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { parseIni, printIni } from '../lib/ini-parser.js';

describe('parseIni', () => {
  // 1. Empty file ? no sections, no orphans, no diagnostics
  it('parses an empty file with no sections, orphans, or diagnostics', () => {
    const { document, diagnostics } = parseIni('', 'test.ini');
    assert.equal(document.sections.length, 0);
    assert.equal(document.orphans.length, 0);
    assert.equal(diagnostics.length, 0);
  });

  // 2. File with only comments ? no sections, no orphans, no diagnostics
  it('parses a file with only comments (;, #, \')', () => {
    const content = "; semicolon comment\n# hash comment\n' apostrophe comment";
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections.length, 0);
    assert.equal(document.orphans.length, 0);
    assert.equal(diagnostics.length, 0);
  });

  // 3. File with only blank lines ? no sections, no orphans, no diagnostics
  it('parses a file with only blank lines', () => {
    const content = '\n\n   \n\t\n';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections.length, 0);
    assert.equal(document.orphans.length, 0);
    assert.equal(diagnostics.length, 0);
  });

  // 4. Section with no keys ? one section with empty entries
  it('parses a section with no keys', () => {
    const content = '[EmptySection]';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections.length, 1);
    assert.equal(document.sections[0].name, 'EMPTYSECTION');
    assert.equal(document.sections[0].entries.length, 0);
    assert.equal(diagnostics.length, 0);
  });

  // 5. Key with = in value ? value should be a=b=c
  it('preserves = characters in values by splitting on first = only', () => {
    const content = '[SECTION]\nKEY=a=b=c';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections[0].entries.length, 1);
    assert.equal(document.sections[0].entries[0].key, 'KEY');
    assert.equal(document.sections[0].entries[0].value, 'a=b=c');
    assert.equal(diagnostics.length, 0);
  });

  // 6. Orphaned keys before first section ? warning diagnostic, key in orphans
  it('produces warning for orphaned keys before any section', () => {
    const content = 'ORPHAN=value\n[SECTION]\nKEY=val';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.orphans.length, 1);
    assert.equal(document.orphans[0].key, 'ORPHAN');
    assert.equal(document.orphans[0].value, 'value');
    assert.equal(diagnostics.length, 1);
    assert.equal(diagnostics[0].severity, 'warning');
    assert.ok(diagnostics[0].message.includes('Orphaned'));
  });

  // 7. Malformed section header ([BROKEN) ? warning diagnostic
  it('produces warning for malformed section header missing ]', () => {
    const content = '[BROKEN\nKEY=val';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections.length, 0);
    const malformedDiags = diagnostics.filter(d => d.message.includes('Malformed'));
    assert.equal(malformedDiags.length, 1);
    assert.equal(malformedDiags[0].severity, 'warning');
  });

  // 8. Non-empty line without = inside a section ? warning diagnostic
  it('produces warning for line without = inside a section', () => {
    const content = '[SECTION]\nthis line has no equals';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    const noEqDiags = diagnostics.filter(d => d.message.includes("without '='"));
    assert.equal(noEqDiags.length, 1);
    assert.equal(noEqDiags[0].severity, 'warning');
    assert.equal(noEqDiags[0].section, 'SECTION');
  });

  // 9. Case normalization: mixed-case section and key names ? all uppercase
  it('normalizes section and key names to uppercase', () => {
    const content = '[MySection]\nmyKey=someValue';
    const { document } = parseIni(content, 'test.ini');
    assert.equal(document.sections[0].name, 'MYSECTION');
    assert.equal(document.sections[0].entries[0].key, 'MYKEY');
    assert.equal(document.sections[0].entries[0].value, 'someValue');
  });

  // 10. Duplicate section names ? warning diagnostic
  it('produces warning for duplicate section names', () => {
    const content = '[SECTION]\nA=1\n[SECTION]\nB=2';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections.length, 2);
    const dupDiags = diagnostics.filter(d => d.message.includes('Duplicate section'));
    assert.equal(dupDiags.length, 1);
    assert.equal(dupDiags[0].severity, 'warning');
  });

  // 11. Duplicate key names within section ? warning diagnostic
  it('produces warning for duplicate key names within a section', () => {
    const content = '[SECTION]\nKEY=first\nKEY=second';
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections[0].entries.length, 2);
    const dupDiags = diagnostics.filter(d => d.message.includes('Duplicate key'));
    assert.equal(dupDiags.length, 1);
    assert.equal(dupDiags[0].severity, 'warning');
  });

  // 12. Windows line endings (\r\n) ? same result as Unix (\n)
  it('handles Windows line endings identically to Unix', () => {
    const unixContent = '[SECTION]\nKEY1=val1\nKEY2=val2';
    const winContent = '[SECTION]\r\nKEY1=val1\r\nKEY2=val2';
    const unix = parseIni(unixContent, 'test.ini');
    const win = parseIni(winContent, 'test.ini');
    assert.equal(win.document.sections.length, unix.document.sections.length);
    assert.equal(win.document.sections[0].name, unix.document.sections[0].name);
    assert.equal(win.document.sections[0].entries.length, unix.document.sections[0].entries.length);
    for (let i = 0; i < unix.document.sections[0].entries.length; i++) {
      assert.equal(win.document.sections[0].entries[i].key, unix.document.sections[0].entries[i].key);
      assert.equal(win.document.sections[0].entries[i].value, unix.document.sections[0].entries[i].value);
    }
    assert.equal(win.diagnostics.length, unix.diagnostics.length);
  });

  // 14. Multiple sections with entries ? correct parsing
  it('parses multiple sections with entries correctly', () => {
    const content = [
      '[INIT]',
      'PORT=7667',
      'NAME=TestServer',
      '',
      '[DATABASE]',
      'HOST=localhost',
      'PORT=3306',
    ].join('\n');
    const { document, diagnostics } = parseIni(content, 'test.ini');
    assert.equal(document.sections.length, 2);
    assert.equal(document.sections[0].name, 'INIT');
    assert.equal(document.sections[0].entries.length, 2);
    assert.equal(document.sections[0].entries[0].key, 'PORT');
    assert.equal(document.sections[0].entries[0].value, '7667');
    assert.equal(document.sections[0].entries[1].key, 'NAME');
    assert.equal(document.sections[0].entries[1].value, 'TestServer');
    assert.equal(document.sections[1].name, 'DATABASE');
    assert.equal(document.sections[1].entries.length, 2);
    assert.equal(document.sections[1].entries[0].key, 'HOST');
    assert.equal(document.sections[1].entries[0].value, 'localhost');
    assert.equal(document.sections[1].entries[1].key, 'PORT');
    assert.equal(document.sections[1].entries[1].value, '3306');
    assert.equal(diagnostics.length, 0);
  });
});

describe('printIni', () => {
  // 13. printIni round-trip with a concrete example
  it('round-trips a concrete INI document', () => {
    const content = [
      '[INIT]',
      'PORT=7667',
      'NAME=TestServer',
      '',
      '[DATABASE]',
      'HOST=localhost',
    ].join('\n');
    const { document } = parseIni(content, 'test.ini');
    const printed = printIni(document);
    const { document: reparsed, diagnostics } = parseIni(printed, 'test.ini');

    assert.equal(reparsed.sections.length, document.sections.length);
    for (let i = 0; i < document.sections.length; i++) {
      assert.equal(reparsed.sections[i].name, document.sections[i].name);
      assert.equal(reparsed.sections[i].entries.length, document.sections[i].entries.length);
      for (let j = 0; j < document.sections[i].entries.length; j++) {
        assert.equal(reparsed.sections[i].entries[j].key, document.sections[i].entries[j].key);
        assert.equal(reparsed.sections[i].entries[j].value, document.sections[i].entries[j].value);
      }
    }
    assert.equal(reparsed.orphans.length, 0);
    assert.equal(diagnostics.length, 0);
  });
});
