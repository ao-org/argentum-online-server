# GM trigger command regression tests

Run `./Codigo/Tests/TriggerCommand/run.ps1` in PowerShell with VB6 and the server's registered Aurora.Network dependency available. Pass `-Compiler` to override the compiler location. The compiler runs with process-local `RunAsInvoker`.

The script extracts the actual command handlers, `SetTileTriggerFlags`, `TilePropertyKey`, `EsGM`, known mask and enums into an ignored native test project. Logging and publication are captured in memory. It does not start the server, open sockets, read maps or access a database.

Coverage includes accepted masks 0, 255, 256 and 511; Dios/Admin authorization; restricted, missing, unknown and conflicting role bits; rejection of negative/unknown masks through the real migrated helper; incomplete one-, two- and three-byte payloads; query and confirmation text; audits and publication; and preservation of the next packet after the four-byte set payload. An injected query value 65536 verifies the full Long response independently of set validation. Results are written under `build/trigger-command-tests`; failures return a nonzero exit code.

This command PR targets `codex/trigger-map-schema-migration` and changes the set transport from one byte to Int32. The migration base supplies flags and live tile updates; this PR integrates its validation helper while keeping the command change separately reviewable. Release it with matching HOO `/trigger` support and the migration.
