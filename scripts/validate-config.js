#!/usr/bin/env node

/**
 * Config File Validator CLI — validates Argentum Online 20 server
 * configuration files against JSON schemas.
 *
 * Usage:
 *   node scripts/validate-config.js
 *     --config-dir <path>   Root directory for config files (default: server root + Recursos/Dat/)
 *     --schema-dir <path>   Directory containing schema JSON files (default: scripts/schemas/)
 *     --file <path...>      Validate only specific files
 *     --format text|json    Output format (default: text)
 *
 * Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7
 */

import { existsSync, readdirSync } from 'fs';
import { resolve, join, extname } from 'path';
import { runValidation } from './lib/validation-runner.js';
import { formatText, formatJson } from './lib/diagnostic-formatter.js';

/**
 * Parse CLI arguments from process.argv.
 * @param {string[]} argv
 * @returns {{ configDir: string|null, schemaDir: string, files: string[], format: string }}
 */
function parseArgs(argv) {
  const args = argv.slice(2);
  let configDir = null;
  let schemaDir = 'scripts/schemas/';
  const files = [];
  let format = 'text';

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--config-dir':
        configDir = args[++i];
        break;
      case '--schema-dir':
        schemaDir = args[++i];
        break;
      case '--file':
        // Collect all following args until the next flag or end
        while (i + 1 < args.length && !args[i + 1].startsWith('--')) {
          files.push(args[++i]);
        }
        break;
      case '--format':
        format = args[++i];
        break;
    }
  }

  return { configDir, schemaDir, files, format };
}

/**
 * Scan a directory for config files with the given extensions.
 * @param {string} dir - Directory to scan
 * @param {string[]} extensions - File extensions to include (e.g. ['.ini', '.dat'])
 * @returns {string[]} Array of absolute file paths
 */
function scanConfigFiles(dir, extensions) {
  if (!existsSync(dir)) {
    return [];
  }
  const results = [];
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true });
  } catch {
    return results;
  }
  for (const entry of entries) {
    if (entry.isFile()) {
      const ext = extname(entry.name).toLowerCase();
      if (extensions.includes(ext)) {
        results.push(join(dir, entry.name));
      }
    }
  }
  return results;
}

/**
 * Discover config files using default locations when no --file is specified.
 * Scans the server root for .ini files and Recursos/Dat/ for .dat files.
 * If --config-dir is provided, scans that single directory for both .ini and .dat files.
 * @param {string|null} configDir - Explicit config directory, or null for defaults
 * @returns {string[]}
 */
function discoverConfigFiles(configDir) {
  if (configDir) {
    const resolved = resolve(configDir);
    return scanConfigFiles(resolved, ['.ini', '.dat']);
  }

  // Default: server root for .ini, Recursos/Dat/ for .dat
  const serverRoot = resolve('.');
  const datDir = resolve('Recursos', 'Dat');

  const iniFiles = scanConfigFiles(serverRoot, ['.ini']);
  const datFiles = scanConfigFiles(datDir, ['.dat']);

  return [...iniFiles, ...datFiles];
}

/**
 * Main entry point.
 */
function main() {
  const { configDir, schemaDir, files, format } = parseArgs(process.argv);
  const resolvedSchemaDir = resolve(schemaDir);

  // Determine which config files to validate
  let configFiles;
  if (files.length > 0) {
    configFiles = files.map(f => resolve(f));
  } else {
    configFiles = discoverConfigFiles(configDir);
  }

  if (configFiles.length === 0) {
    console.log('No config files found to validate.');
    process.exit(0);
  }

  // Run validation
  const { diagnostics, exitCode } = runValidation({
    schemaDir: resolvedSchemaDir,
    configFiles,
  });

  // Format and print output
  const output = format === 'json'
    ? formatJson(diagnostics)
    : formatText(diagnostics);

  console.log(output);

  // Exit with appropriate code
  process.exit(exitCode);
}

main();
