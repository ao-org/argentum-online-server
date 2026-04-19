/**
 * Diagnostic Formatter module.
 *
 * Formats validation diagnostics as human-readable text or JSON,
 * and provides an exit code helper.
 *
 * Text format per diagnostic:
 *   SEVERITY file:line [SECTION] KEY: message
 *
 * When section is null, the [SECTION] part is omitted.
 * When key is null, the KEY part is omitted.
 *
 * A summary line is appended:
 *   INFO: Validation complete. N error(s), N warning(s), N info(s).
 *
 * JSON format: JSON array of diagnostic objects (pretty-printed).
 */

/**
 * Compute summary counts from a list of diagnostics.
 * @param {import('./data-models.js').Diagnostic[]} diagnostics
 * @returns {{ errors: number, warnings: number, infos: number }}
 */
function computeSummary(diagnostics) {
  let errors = 0;
  let warnings = 0;
  let infos = 0;
  for (const d of diagnostics) {
    switch (d.severity) {
      case 'error': errors++; break;
      case 'warning': warnings++; break;
      case 'info': infos++; break;
    }
  }
  return { errors, warnings, infos };
}

/**
 * Format a single diagnostic as a text line.
 * Format: SEVERITY file:line [SECTION] KEY: message
 * @param {import('./data-models.js').Diagnostic} d
 * @returns {string}
 */
function formatDiagnosticLine(d) {
  const severity = d.severity.toUpperCase();
  let parts = [`${severity} ${d.file}:${d.line}`];

  if (d.section != null) {
    parts.push(`[${d.section}]`);
  }

  if (d.key != null) {
    parts.push(`${d.key}:`);
  } else if (d.section != null) {
    // When section present but no key, append colon to section bracket
    // e.g. "WARNING Server.ini:10 [UNKNOWN_SECTION]: message"
    parts[parts.length - 1] += ':';
  } else {
    // No section, no key — just append colon to file:line
    parts[parts.length - 1] += ':';
  }

  parts.push(d.message);
  return parts.join(' ');
}

/**
 * Format diagnostics as human-readable text.
 * Each diagnostic is one line, followed by a summary line.
 *
 * @param {import('./data-models.js').Diagnostic[]} diagnostics
 * @returns {string}
 */
export function formatText(diagnostics) {
  const lines = diagnostics.map(formatDiagnosticLine);
  const { errors, warnings, infos } = computeSummary(diagnostics);
  lines.push(`INFO: Validation complete. ${errors} error(s), ${warnings} warning(s), ${infos} info(s).`);
  return lines.join('\n');
}

/**
 * Format diagnostics as a JSON array (pretty-printed).
 *
 * @param {import('./data-models.js').Diagnostic[]} diagnostics
 * @returns {string}
 */
export function formatJson(diagnostics) {
  return JSON.stringify(diagnostics, null, 2);
}

/**
 * Compute the process exit code based on diagnostics.
 * Returns 1 if any diagnostic has severity 'error', 0 otherwise.
 *
 * @param {import('./data-models.js').Diagnostic[]} diagnostics
 * @returns {number}
 */
export function computeExitCode(diagnostics) {
  return diagnostics.some(d => d.severity === 'error') ? 1 : 0;
}
