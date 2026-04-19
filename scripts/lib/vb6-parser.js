/**
 * VB6 Source Parser for schema auto-generation.
 *
 * Extracts config-reading patterns from VB6 source files (.cls, .bas)
 * and produces SchemaEntry objects. Also provides schema merging to
 * preserve manual annotations when regenerating.
 *
 * Recognized patterns:
 * - mSettings.Add "KEY", TypeFn(val(reader.GetValue("SECTION", "KEY"[, DEFAULT])))
 * - mSettings.Add "KEY", reader.GetValue("SECTION", "KEY")
 * - GetVar(filePath, "SECTION", "KEY")
 */

import { createSchemaEntry, createConfigSchema } from './data-models.js';

/**
 * Map VB6 type conversion functions to schema types.
 * @type {Object<string, string>}
 */
const VB6_TYPE_MAP = {
  CINT: 'integer',
  CLNG: 'long',
  CSNG: 'single',
  CDBL: 'double',
  CBYTE: 'integer',
};

/**
 * Regex for mSettings.Add with a type wrapper around val(reader.GetValue(...)).
 * Captures: 1=type function, 2=section, 3=key, 4=optional default value
 */
const RE_SETTINGS_TYPED = /mSettings\.Add\s+"[^"]+"\s*,\s*(CLng|CInt|CDbl|CSng|CByte)\s*\(\s*val\s*\(\s*reader\.GetValue\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"(?:\s*,\s*([^)]+))?\s*\)/i;

/**
 * Regex for mSettings.Add with reader.GetValue but NO type wrapper (string type).
 * Captures: 1=section, 2=key
 */
const RE_SETTINGS_STRING = /mSettings\.Add\s+"[^"]+"\s*,\s*reader\.GetValue\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"/i;

/**
 * Regex for GetVar(filePath, "SECTION", "KEY") calls.
 * Captures: 1=filePath expression, 2=section, 3=key
 */
const RE_GETVAR = /GetVar\s*\(\s*([^,]+)\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)/i;


/**
 * Try to resolve a VB6 file path expression to a config file name.
 * Common patterns:
 *   - App.Path & "\Server.ini"  ? "Server.ini"
 *   - DatPath & "NPCs.dat"     ? "NPCs.dat"
 *   - IniPath & "Configuracion.ini" ? "Configuracion.ini"
 *   - String literal "C:\path\file.ini" ? "file.ini"
 *   - Variable like `npcfile` ? null (unresolvable)
 *
 * @param {string} expr - VB6 expression for the file path
 * @returns {string|null} Resolved file name or null
 */
function resolveFilePath(expr) {
  const trimmed = expr.trim();

  // Pattern: something & "...\filename" or something & "filename"
  const concatMatch = trimmed.match(/&\s*"([^"]+)"\s*$/i);
  if (concatMatch) {
    const pathPart = concatMatch[1];
    // Extract just the filename from the path (after last \ or /)
    const parts = pathPart.split(/[\\/]/);
    return parts[parts.length - 1] || null;
  }

  // Pattern: string literal "path\to\file.ini"
  const literalMatch = trimmed.match(/^"([^"]+)"$/);
  if (literalMatch) {
    const pathPart = literalMatch[1];
    const parts = pathPart.split(/[\\/]/);
    return parts[parts.length - 1] || null;
  }

  // Unresolvable expression (bare variable, function call, etc.)
  return null;
}

/**
 * Determine the schema type from a surrounding type conversion function
 * found on the same line as a GetVar call.
 *
 * @param {string} line - Full VB6 source line
 * @param {number} getVarStart - Character index where GetVar starts
 * @returns {string} Schema type
 */
function inferTypeFromContext(line, getVarStart) {
  // Look for a type conversion function wrapping or preceding the GetVar call
  const prefix = line.substring(0, getVarStart);
  // Check for pattern like CLng(val(GetVar... or CInt(GetVar... or val(GetVar...
  const typeMatch = prefix.match(/(CLng|CInt|CDbl|CSng|CByte)\s*\(\s*(val\s*\()?\s*$/i);
  if (typeMatch) {
    return VB6_TYPE_MAP[typeMatch[1].toUpperCase()] || 'string';
  }
  // Check for val( without type wrapper — VB6 val() returns Double
  const valMatch = prefix.match(/val\s*\(\s*$/i);
  if (valMatch) {
    return 'string';
  }
  return 'string';
}

/**
 * Parse a VB6 source file and extract config-reading patterns as SchemaEntry objects.
 *
 * @param {string} content - VB6 source file content
 * @param {string} filePath - Source file path (for the `source` field in entries)
 * @returns {import('./data-models.js').SchemaEntry[]}
 */
export function extractSchemaEntries(content, filePath) {
  const entries = [];
  const lines = content.split(/\r?\n/);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineNum = i + 1;
    const source = `${filePath}:${lineNum}`;

    // Pattern 1: mSettings.Add with type wrapper
    const typedMatch = line.match(RE_SETTINGS_TYPED);
    if (typedMatch) {
      const typeFn = typedMatch[1].toUpperCase();
      const section = typedMatch[2].toUpperCase();
      const key = typedMatch[3].toUpperCase();
      const rawDefault = typedMatch[4];
      const schemaType = VB6_TYPE_MAP[typeFn] || 'string';
      const hasDefault = rawDefault !== undefined && rawDefault !== null;
      const defaultValue = hasDefault ? rawDefault.trim() : null;
      const required = !hasDefault;

      entries.push(createSchemaEntry(
        'Configuracion.ini', // mSettings.Add lines read from the config file loaded by reader
        section,
        key,
        schemaType,
        required,
        defaultValue,
        source
      ));
      continue;
    }

    // Pattern 2: mSettings.Add string (no type wrapper)
    const stringMatch = line.match(RE_SETTINGS_STRING);
    if (stringMatch) {
      const section = stringMatch[1].toUpperCase();
      const key = stringMatch[2].toUpperCase();

      entries.push(createSchemaEntry(
        'Configuracion.ini',
        section,
        key,
        'string',
        true, // no default ? required
        null,
        source
      ));
      continue;
    }

    // Pattern 3: GetVar calls
    const getVarMatch = line.match(RE_GETVAR);
    if (getVarMatch) {
      const fileExpr = getVarMatch[1];
      const section = getVarMatch[2].toUpperCase();
      const key = getVarMatch[3].toUpperCase();
      const resolvedFile = resolveFilePath(fileExpr);
      const schemaType = inferTypeFromContext(line, line.search(RE_GETVAR));

      if (resolvedFile) {
        entries.push(createSchemaEntry(
          resolvedFile,
          section,
          key,
          schemaType,
          true, // GetVar has no default mechanism in the call itself
          null,
          source
        ));
      }
    }
  }

  return entries;
}


/**
 * Merge an auto-generated ConfigSchema with an existing one, preserving
 * manual annotations from the existing schema.
 *
 * Merge strategy:
 * - All sections and keys from `generated` are included.
 * - For keys that exist in both, manual annotations from `existing` are preserved:
 *   min, max, pattern, description, rules (on keys).
 * - Section-level properties like `strict` are preserved from `existing`.
 * - The `rules` array from `existing` is preserved.
 * - The `description` from `existing` is preserved if present.
 *
 * @param {import('./data-models.js').ConfigSchema} generated - Auto-generated schema
 * @param {import('./data-models.js').ConfigSchema|null} existing - Existing schema with manual annotations, or null
 * @returns {import('./data-models.js').ConfigSchema}
 */
export function mergeSchemas(generated, existing) {
  if (!existing) {
    return generated;
  }

  const merged = createConfigSchema(
    generated.file,
    existing.description || generated.description,
    {},
    existing.rules && existing.rules.length > 0 ? existing.rules : generated.rules
  );

  // Merge sections: start with all generated sections
  for (const [sectionName, genSection] of Object.entries(generated.sections)) {
    const existSection = existing.sections[sectionName];

    if (!existSection) {
      // New section from generated — use as-is
      merged.sections[sectionName] = { ...genSection };
      continue;
    }

    // Merge section-level properties: preserve strict, required from existing if set
    const mergedSection = {
      required: genSection.required !== undefined ? genSection.required : existSection.required,
      strict: existSection.strict !== undefined ? existSection.strict : genSection.strict,
      keys: {},
    };

    // Merge keys
    for (const [keyName, genKey] of Object.entries(genSection.keys || {})) {
      const existKey = existSection.keys ? existSection.keys[keyName] : undefined;

      if (!existKey) {
        // New key from generated
        mergedSection.keys[keyName] = { ...genKey };
        continue;
      }

      // Merge: generated provides type/required/default, existing provides annotations
      mergedSection.keys[keyName] = {
        type: genKey.type,
        required: genKey.required,
        ...(genKey.default !== undefined && { default: genKey.default }),
        // Preserve manual annotations from existing
        ...(existKey.min !== undefined && { min: existKey.min }),
        ...(existKey.max !== undefined && { max: existKey.max }),
        ...(existKey.pattern !== undefined && { pattern: existKey.pattern }),
        ...(existKey.description !== undefined && { description: existKey.description }),
        ...(existKey.rules !== undefined && { rules: existKey.rules }),
      };
    }

    merged.sections[sectionName] = mergedSection;
  }

  return merged;
}
