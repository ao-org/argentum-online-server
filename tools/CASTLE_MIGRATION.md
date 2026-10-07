# Offline castle schema migration

`castle_migration.py` implements version 1 of the castle migration in [the shared plan](../Documentation/TriggerMigration.md). It accepts an explicit SQLite snapshot and publishes a separate validated database. It never edits the source or overwrites an existing output. Use Python 3.10 or later with SQLite and no additional packages.

Stop castle writes and take a consistent database backup before the release window. Rehearse against that backup, with the exact release maps available for bounds checks:

```powershell
python tools/castle_migration.py --source backup.db --maps ../Recursos/Mapas --dry-run
python tools/castle_migration.py --source backup.db --maps ../Recursos/Mapas --output migrated.db
```

The dry run performs the same transactional migration and validation on a temporary copy. JSON output includes counts, the legacy trigger-to-castle mapping, placements, owned/unplaced castles, and unconfigured slots. Capture this report with the release manifest. Publication uses an atomic hard link on the output filesystem; if that filesystem does not support links, use a local filesystem that does. No output is published before validation succeeds.

The tool validates the documented schema, foreign keys, unique references, all assigned interior positions, full outside footprints, and duplicate entrances. It fails on unexpected constraints or dependencies on the removed trigger column so they can be reviewed explicitly. It rebuilds `castle` without `trigger`, normalizes only all-zero outside triples to all NULL, and preserves every castle and coordinate ID, owner, item reference, name, stored date value, activity value, whitelist row, supported index/trigger/view, and autoincrement sequence. Missing coordinate rows remain missing.

`castle_schema_migrations` records version 1. A repeated run validates that version without converting again. `castle_legacy_trigger_map` retains the original mapping for legacy map conversion; migrated gameplay references stable `castle.id`. Existing general `migrations` history is preserved and the tool is not run automatically by the server's SQL-file migration runner.

For the supplied production shape, expect 20 castles, 15 coordinate rows with IDs 16–30, eight entrance bindings on maps 27 and 546, owned/unplaced castle 2, and unconfigured castles 16–20. These counts are validated in synthetic tests. The tool inventories its actual input rather than assuming production has not changed. The clipped name in the screenshot is never reconstructed.

After rehearsal, release the validated database with the matching server, HOO, and CSM3 asset revisions. Keep the original snapshot and old binaries/assets together for rollback; new masks and subsequent ownership changes cannot be reversed by recreating the old trigger column. Production deployment is a separate operation.

Run the isolated tests:

```powershell
python -m unittest discover -s tools -p test_castle_migration.py -v
```

Tests build their own databases from the historical table definitions and the supplied row shape. They cover stable/non-contiguous IDs, fractional timestamps, missing interiors, read-only source preservation, nullable placement constraints, incoming foreign keys, views/indexes/triggers, sequence continuity, rollback, idempotence, and malformed input. No production database is read by the tests.
