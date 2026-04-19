# Config File Schemas

This directory contains JSON schema files used by the config file validation tool (`scripts/validate-config.js`).

Each `.json` file declares the expected structure for a specific server configuration file, including:

- Section names and whether they are required or optional
- Key names, data types (`string`, `integer`, `long`, `single`, `double`, `boolean`), and required/optional status
- Default values, numeric bounds (`min`/`max`), and format patterns
- Cross-field validation rules (min/max relationships, count-indexed sections, admin entry formats)

## Schema Generation

Schemas are auto-generated from VB6 source files using:

```sh
node scripts/generate-schema.js --source-dir Codigo/ --output-dir scripts/schemas/
```

Manual annotations (bounds, patterns, rules) are preserved across regeneration via schema merging.

## Usage

The validation tool loads schemas from this directory:

```sh
node scripts/validate-config.js --schema-dir scripts/schemas/
```
