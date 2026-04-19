#!/usr/bin/env node

/**
 * Schema Generator CLI — parses VB6 source files and generates JSON schema
 * definitions for config file validation.
 *
 * Usage:
 *   node scripts/generate-schema.js
 *     --source-dir <path>   VB6 source directory (default: Codigo/)
 *     --output-dir <path>   Schema output directory (default: scripts/schemas/)
 *     --check               Compare only, exit non-zero on diff
 *
 * Requirements: 9.8, 9.9, 9.10, 9.11
 */

import { readFileSync, writeFileSync, readdirSync, existsSync, mkdirSync } from 'fs';
import { join, resolve, extname } from 'path';
import { extractSchemaEntries, mergeSchemas } from './lib/vb6-parser.js';
import { loadSchemas } from './lib/schema-registry.js';
import { createConfigSchema } from './lib/data-models.js';

/**
 * Parse CLI arguments from process.argv.
 * @returns {{ sourceDir: string, outputDir: string, check: boolean }}
 */
function parseArgs(argv) {
  const args = argv.slice(2);
  let sourceDir = 'Codigo/';
  let outputDir = 'scripts/schemas/';
  let check = false;

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--source-dir':
        sourceDir = args[++i];
        break;
      case '--output-dir':
        outputDir = args[++i];
        break;
      case '--check':
        check = true;
        break;
    }
  }

  return { sourceDir, outputDir, check };
}

/**
 * Recursively scan a directory for .cls and .bas files.
 * @param {string} dir - Directory to scan
 * @returns {string[]} Array of file paths
 */
function scanVb6Files(dir) {
  const results = [];
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true });
  } catch {
    return results;
  }
  for (const entry of entries) {
    const fullPath = join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...scanVb6Files(fullPath));
    } else {
      const ext = extname(entry.name).toLowerCase();
      if (ext === '.cls' || ext === '.bas') {
        results.push(fullPath);
      }
    }
  }
  return results;
}

/**
 * Group SchemaEntry[] by config file name (entry.file).
 * @param {import('./lib/data-models.js').SchemaEntry[]} entries
 * @returns {Map<string, import('./lib/data-models.js').SchemaEntry[]>}
 */
function groupByConfigFile(entries) {
  const groups = new Map();
  for (const entry of entries) {
    const key = entry.file;
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(entry);
  }
  return groups;
}

/**
 * Build a ConfigSchema from grouped SchemaEntry objects.
 * @param {string} configFile - Config file name
 * @param {import('./lib/data-models.js').SchemaEntry[]} entries
 * @returns {import('./lib/data-models.js').ConfigSchema}
 */
function buildSchema(configFile, entries) {
  const sections = {};

  for (const entry of entries) {
    if (!sections[entry.section]) {
      sections[entry.section] = {
        required: true,
        strict: true,
        keys: {},
      };
    }

    const keyDef = {
      type: entry.type,
      required: entry.required,
    };
    if (entry.defaultValue !== null) {
      keyDef.default = entry.defaultValue;
    }

    // Only add if not already present (first occurrence wins)
    if (!sections[entry.section].keys[entry.key]) {
      sections[entry.section].keys[entry.key] = keyDef;
    }
  }

  return createConfigSchema(configFile, '', sections, []);
}

/**
 * Derive the schema filename from a config file name.
 * E.g., "Server.ini" -> "Server.ini.json"
 * @param {string} configFile
 * @returns {string}
 */
function schemaFilename(configFile) {
  return `${configFile}.json`;
}

/**
 * Main entry point.
 */
function main() {
  const { sourceDir, outputDir, check } = parseArgs(process.argv);
  const resolvedSourceDir = resolve(sourceDir);
  const resolvedOutputDir = resolve(outputDir);

  // 1. Scan VB6 source files
  const vb6Files = scanVb6Files(resolvedSourceDir);
  if (vb6Files.length === 0) {
    console.error(`No .cls or .bas files found in ${resolvedSourceDir}`);
    process.exit(1);
  }

  // 2. Extract schema entries from all files
  const allEntries = [];
  for (const filePath of vb6Files) {
    let content;
    try {
      content = readFileSync(filePath, 'utf-8');
    } catch (err) {
      console.error(`Warning: Cannot read ${filePath}: ${err.message}`);
      continue;
    }
    const entries = extractSchemaEntries(content, filePath);
    allEntries.push(...entries);
  }

  // 3. Group entries by config file
  const grouped = groupByConfigFile(allEntries);

  // 4. Build generated schemas
  const generatedSchemas = new Map();
  for (const [configFile, entries] of grouped) {
    generatedSchemas.set(configFile, buildSchema(configFile, entries));
  }

  // 5. Load existing schemas if output dir exists
  let existingSchemas = new Map();
  if (existsSync(resolvedOutputDir)) {
    const loaded = loadSchemas(resolvedOutputDir);
    existingSchemas = loaded.schemas;
  }

  // 6. Merge generated with existing
  const mergedSchemas = new Map();
  for (const [configFile, generated] of generatedSchemas) {
    const existingKey = configFile.toUpperCase();
    const existing = existingSchemas.get(existingKey) || null;
    mergedSchemas.set(configFile, mergeSchemas(generated, existing));
  }

  // 7. Check mode or write mode
  if (check) {
    const diffs = [];
    for (const [configFile, merged] of mergedSchemas) {
      const filename = schemaFilename(configFile);
      const filePath = join(resolvedOutputDir, filename);
      const newContent = JSON.stringify(merged, null, 2) + '\n';

      if (!existsSync(filePath)) {
        diffs.push(`${filename} (new file, not yet committed)`);
        continue;
      }

      let existingContent;
      try {
        existingContent = readFileSync(filePath, 'utf-8');
      } catch (err) {
        diffs.push(`${filename} (cannot read: ${err.message})`);
        continue;
      }

      if (existingContent !== newContent) {
        diffs.push(filename);
      }
    }

    if (diffs.length > 0) {
      console.error('Schema files differ from generated output:');
      for (const diff of diffs) {
        console.error(`  - ${diff}`);
      }
      process.exit(1);
    } else {
      console.log('All schema files are up to date.');
      process.exit(0);
    }
  } else {
    // Write mode
    if (!existsSync(resolvedOutputDir)) {
      mkdirSync(resolvedOutputDir, { recursive: true });
    }

    const written = [];
    for (const [configFile, merged] of mergedSchemas) {
      const filename = schemaFilename(configFile);
      const filePath = join(resolvedOutputDir, filename);
      const content = JSON.stringify(merged, null, 2) + '\n';
      writeFileSync(filePath, content, 'utf-8');
      written.push(filename);
    }

    console.log(`Generated ${written.length} schema file(s):`);
    for (const f of written) {
      console.log(`  - ${f}`);
    }
  }
}

main();
