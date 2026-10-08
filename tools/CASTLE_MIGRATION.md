# Castle SQL migration rehearsal

The canonical migration is [20261007-01-migrate castle identities.sql](../ScriptsDB/20261007-01-migrate%20castle%20identities.sql), following the repository's `YYYYMMDD-XX-description.sql` convention. The existing server migration runner applies it and records `20261007-01` in `migrations`. SQLite 3.35 or newer is required for `ALTER TABLE DROP COLUMN`.

`castle_migration.py` rehearses that same SQL against an explicit SQLite snapshot and optionally publishes a separate validated database. It never edits the source or overwrites an existing output. Use Python 3.10 or later with SQLite 3.35 or newer and no additional packages.

Stop castle writes and take a consistent database backup before the release window. Rehearse against that backup, with the exact release maps available for bounds checks:

```powershell
python tools/castle_migration.py --source backup.db --maps ../Recursos/Mapas --dry-run
python tools/castle_migration.py --source backup.db --maps ../Recursos/Mapas --output migrated.db
```

The dry run performs the same transactional migration and validation on a temporary copy. JSON output includes counts, the legacy trigger-to-castle mapping, placements, owned/unplaced castles, and unconfigured slots. Capture this report with the release manifest. Publication uses an atomic hard link on the output filesystem; if that filesystem does not support links, use a local filesystem that does. No output is published before validation succeeds.

The tool validates the documented schema, foreign keys, unique references, all assigned interior positions, full outside footprints, and duplicate entrances. It fails on unexpected constraints or dependencies on the removed columns so they can be reviewed explicitly. The SQL captures the legacy mapping, drops only `castle.trigger`, and replaces the three outside placement columns with nullable columns and a tuple constraint. Only all-zero outside triples become all NULL. Tables, every castle and coordinate ID, owners, item references, names, stored dates, activity values, whitelists, unaffected indexes/triggers/views, and autoincrement sequences remain intact. Missing coordinate rows remain missing.

`castle_legacy_trigger_map` retains the original mapping for legacy map conversion; migrated gameplay references stable `castle.id`. The rehearsal records the same ordinary `migrations` key as the server, so a prepared output is not migrated again at startup. Existing migration history is preserved. The tool adds preservation checks before committing the SQL's transaction. The server runner rolls back a failed script and startup requires the normal migration marker. There is no separate castle migration registry.

For the supplied production shape, expect 20 castles, 15 coordinate rows with IDs 16–30, eight entrance bindings on maps 27 and 546, owned/unplaced castle 2, and unconfigured castles 16–20. These counts are validated in synthetic tests. The tool inventories its actual input rather than assuming production has not changed. The clipped name in the screenshot is never reconstructed.

After rehearsal, release the dated SQL with the matching server, HOO, and CSM3 asset revisions. The normal startup runner applies it to an unmigrated database. Alternatively, a validated output already contains the normal migration marker. Keep the original snapshot and old binaries/assets together for rollback; new masks and subsequent ownership changes cannot be reversed by recreating the old trigger column. Production deployment is a separate operation.

Run the isolated tests:

```powershell
python -m unittest discover -s tools -p test_castle_migration.py -v

# After building the native harness with tools/test_map_schema.ps1:
& "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -File tools/test_castle_sql.ps1
```

Tests build their own databases from the historical table definitions and the supplied row shape. They cover stable/non-contiguous IDs, fractional timestamps, missing interiors, read-only source preservation, nullable placement constraints, incoming foreign keys, views/indexes/triggers, sequence continuity, rollback, idempotence, and malformed input. The PowerShell test takes statements from the production VB6 splitter and executes them through prepared commands in the actual 32-bit ADO/SQLite driver. It verifies both successful migration and rollback after a malformed tuple. No production database is read by the tests.
