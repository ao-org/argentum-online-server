# External production configuration

Start the server with an operator-managed directory outside deployment checkouts:

```text
server.exe /config-dir="C:\Ao20\ServerConfig"
server.exe /config-dir="C:\Ao20\Server Config"
```

Only these six filenames are managed:

| File | Without `/config-dir` | With `/config-dir` |
| --- | --- | --- |
| `Server.ini` | `IniPath` (application directory) | selected directory |
| `Configuracion.ini` | `IniPath` | selected directory |
| `feature_toggle.ini` | `IniPath` | selected directory |
| `Balance.dat` | configured `DatPath` | selected directory |
| `Eventos.dat` | configured `DatPath` | selected directory |
| `intervalos.ini` | `IniPath` (application directory) | selected directory |

`IniPath` and `DatPath` are not redirected. Other resource paths, databases,
keys, maps and DAT files retain their existing behavior. Relative directory
arguments are resolved once against the launch working directory, before the
server changes it. Use an absolute path in service configuration.

The argument name is case-insensitive. Both `/config-dir="path with spaces"`
and `/config-dir "path with spaces"` work. Empty, duplicate or unterminated
arguments are rejected. Option-like text inside a quoted path is not another option. Unknown arguments
and the standalone `--benchmark-string-builders` entry point remain supported.

## Startup and errors

Before configuration loaders, database startup or listeners,
external mode checks that the directory exists and **all six** files can be
opened and read. A required file missing, locked against reads or inaccessible
is fatal. No legacy path, example file or default configuration replaces it.
Each subsequent managed read checks its selected file again and propagates
read/parse failures. Empty files and files without INI sections are rejected.

Startup failures log a fatal error with the path and terminate with exit code 1,
without a modal dialog. Successful startup logs the mode and all six resolved
paths to the server console and Windows Event Log (`Argentum20`). File contents
and secrets are not added to these path diagnostics.

Without the argument, the existing paths apply. A missing feature-toggle file is
still optional on the first legacy load. Once loaded, a missing file during
reload is an error and retains the active toggles. The former current-directory
existence check now uses the same path as the actual feature loader.

This is a location and reload-safety change, not a new configuration schema.
Existing field defaults, `Val` conversions, numeric types, Win32 quoted-value
semantics for scheduled events/the spell allowlist, and gameplay formulas
remain. A missing optional **key** may still use its existing loader default;
that is different from replacing a missing **file**. This is not comprehensive
semantic validation of every field. Invasion entries with nonpositive `MaxNPCs`
or `SpawnBoxes` now report failure instead of returning with a partially replaced
invasion array. A configured invasion count of zero still disables invasions.

## Source audit and reload boundaries

Audit base: `16546f49f55c72a298e81ce8b2a2a0b15b00d85c`.

| Source/function | Access and treatment |
| --- | --- |
| [General.bas](../Codigo/General.bas): `Main`, `InicializarConstantes`, `Restart` | Early external preflight; legacy base initialization retained; restart uses the shared loaders. |
| [FileIO.bas](../Codigo/FileIO.bas): `loadAdministrativeUsers` | Reads `Server.ini`; builds both privilege containers before replacing either. |
| `FileIO.bas`: `LoadSini`, `LoadFeatureToggles` | Stages startup scalars, paths and feature values before applying them. Preserves the feature pursuit-toggle side effect after parsing. |
| `FileIO.bas`: `LoadIntervalos`, `SaveIntervalos` and [FrmInterv.frm](../Codigo/FrmInterv.frm): `Command2_Click` | Stages every timing value before publishing globals/UI or resetting NPC AI scheduling. Startup, restart and existing reload controls share this path. Save writes applied values to the selected `intervalos.ini`, correcting the former `intervalo.ini` filename mismatch; hunger/thirst values are converted back to file units. |
| `FileIO.bas`: `LoadBalance` | Stages scalars, class/race modifiers, experience, home tiers and elemental matrix before publishing. Existing home-tier validation/fallback retained. |
| `FileIO.bas`: `LoadMainConfigFile` and [ServerConfig.cls](../Codigo/ServerConfig.cls): `LoadSettings` | Builds a new settings object plus scheduled events and invasions before publishing managed state. |
| [ModEventos.bas](../Codigo/ModEventos.bas): `CargarEventos` | Stages the 24 scheduled events from the selected `Configuracion.ini`. |
| [ModInvasion.bas](../Codigo/ModInvasion.bas): `CargarInfoEventos` | Stages nested invasion arrays from the selected `Eventos.dat`; retains active state and timer setting on read/conversion failure. |
| [Protocol.bas](../Codigo/Protocol.bas): `HandleServerOpenToUsersToggle` | Persists `ServerSoloGMs` in the selected `Server.ini`; only changes the live flag after a successful write. |
| `Protocol.bas`: `IsUnassistedSpellAllowed` | Reads the selected `Configuracion.ini`; retains its existing process-lifetime cache. |
| [frmMain.frm](../Codigo/frmMain.frm), [frmServidor.frm](../Codigo/frmServidor.frm) | Existing reload buttons and restart paths call the shared loaders; no separate filename access. |
| [clsIniManager.cls](../Codigo/clsIniManager.cls): `Initialize` | Opt-in strict reads propagate failures; unrelated legacy readers retain their existing contract. |
| [modManagedConfig.bas](../Codigo/modManagedConfig.bas) | Central six-file registry, argument parsing, preflight, path resolution, diagnostics and checked write. |

The old loaders cleared dictionaries/arrays or assigned globals before completing
reads and conversions. Failed managed reads/conversions now report an error to
the console/Event Log and leave their previously loaded managed values intact.
There is no fallback, automatic retry, disk rollback or automatic migration.
A successful reload retains the existing reset behavior, including invasion
runtime-array/timer reset; avoid reloading during a live invasion.

**Limits:** these controls are not one atomic reload of the entire server.
`LoadSini` also loads `Ciudades.dat`; the pursuit feature may reload topology.
`LoadMainConfigFile` also reloads `Retos.dat` and `MapasEspeciales.dat` at their
original locations. Those unrelated loaders and their side effects are not
transactional here, and some still swallow their own errors. Their state can
partially change on failure; no whole-world rollback is claimed. Managed parsing
finishes before these auxiliary reloads; main-config managed publication follows
them. Restart already disconnects players/resets world state before `LoadSini`
and is not made transactional. Some existing server reload buttons also reload the
private key afterward; that separate operation is unchanged. The cached spell allowlist still requires a
restart to pick up edits. Out-of-memory failures during final publication are
not covered by the parse-before-publish guarantee. Interval reload retains the
original integer widths, `Val` conversions, hunger/thirst `/25` scaling and
NPC-AI/world-save defaults. The interval editor still saves applied globals;
its Apply behavior is unchanged. Save is a sequence of checked INI-key writes,
not an atomic file replacement: a mid-save I/O failure is reported but earlier
disk writes may remain. It does not change live globals. Failed read/parse reloads
leave active interval globals, UI fields and NPC AI scheduling untouched.

## Additional configuration audit: deliberately out of scope

| File | Current access | Recommendation |
| --- | --- | --- |
| `PacketRatePolicy.ini` | `FileIO.LoadPacketRatePolicy`, `IniPath`; startup | Include in a separately approved extension: operational anti-macro/rate policy. Its arrays and derived intervals need staged application. Not included now. |

`intervalos.ini` was added to the managed set after explicit scope approval.
Putting `PacketRatePolicy.ini` in `/config-dir` has no effect in this change. Do not assume
all operational configuration has been removed from deployment-controlled paths.

The obsolete `modPartition`, `Partitions.ini` and its three parser-only tests
have been removed. The module's ownership/endpoint loader had no runtime callers;
normal map loading is unchanged. Startup now always uses the existing single-world
process-count check and version caption. `/shard` is no longer interpreted and,
like other unknown arguments, has no effect. Remove it from launch commands;
no partition configuration is read or required.

## Staging rollout, migration and rollback

1. In staging, record the executable revision and current effective configuration
   locations. Back up the six files through the existing private operations
   process. Record the out-of-scope `PacketRatePolicy.ini` separately.
2. An operator creates the external directory and copies the **current active**
   six files there. Preserve values/encoding; do not use examples as production
   replacements. No automated migration is provided.
3. Give the service identity read/traverse access to the directory and all six
   files. The existing GM server-open toggle additionally needs write permission
   on `Server.ini`; the interval editor save control needs write permission on
   `intervalos.ini`. Test using that identity. Keep the directory outside Jenkins
   workspaces and deployment cleanup paths; Jenkins is not modified by this PR.
4. Stop the staging server and add the absolute `/config-dir` argument to its
   existing launch command. Remove obsolete `/shard` arguments. Verify logged
   paths, normal startup/login, and unchanged gameplay/configuration values.
5. On a disposable staging copy, remove each required file or deny reads and
   verify fatal startup with exit 1 and the exact path in the error. Restore it.
   Ensure no listener/database-dependent startup occurred on preflight failure.
6. While staging is running, test the existing server/admin, balance, interval and main
   reload controls. Change only harmless staging values, then test missing,
   locked and overflowing input. Check errors and retained managed values;
   correct the file and verify a subsequent reload succeeds. Test the GM toggle
   writes only the external `Server.ini`; deny writes and verify no live change.
   Test interval-editor save followed by reload: only external `intervalos.ini`
   should change, and hunger/thirst timing must stay unchanged. A late numeric
   overflow in interval reload must leave earlier timing values/UI intact.
7. Deploy through the normal release process without changing Jenkins. Confirm a
   subsequent deployment leaves the external files untouched. Avoid simultaneous
   edits: publish complete files by an operator-controlled replace and reload
   only after the full intended set is ready; no cross-file snapshot is promised.
8. Roll back by stopping the process and restoring the prior launch command
   (omit `/config-dir`) and binary as needed. Verify legacy files still contain
   the intended values before restart. External edits are **not** copied back;
   reconcile them manually in the private operations process if required.

Never commit production files, credentials or configuration dumps. Existing
`Example.*.ini` files remain unchanged. The original five live filenames are ignored;
the tracked legacy `intervalos.ini` remains unchanged as a development default.
Do not overwrite that tracked file with production values. External production
files stay outside the repository.

## Reproducible validation

From a development checkout with the registered Windows VB6 dependencies:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/test-managed-config.ps1
```

The script builds the real server with `UNIT_TEST=1`, runs
`server.exe --test-managed-config`, and checks the generated
`managed_config_test_results.txt`. The entry point executes before normal
startup, uses only temporary synthetic files, and does not open listeners,
connect to a database or read private configuration. The script also exercises
the real `Main` with a nonexistent external directory and requires exit code 1.
It leaves a test executable in the development checkout; rebuild with
`UNIT_TEST=0` before deploying. Do not run this build script in a deployment
folder.

Tests cover legacy/external selection, quoted/relative/trailing-slash paths,
argument coexistence and rejection, all six missing/locked files, no fallback
or file creation, strict empty-file handling, read-only/hidden files, managed
writes, and successful/failed reloads through the actual loaders.

Full production-equivalent startup, service-account ACLs, real resource reloads,
active players/invasions and the operator UI need staging validation. The
standalone configuration suite does not replace the initialized game-server
unit suite or that rollout checklist.
