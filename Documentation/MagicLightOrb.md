# Following Orb of Light (spell 352)

Deploy the matching HOO client and resource update before this server. No new
packet IDs, capability flags, database fields or objects are introduced. Restart
the server during deployment; live floor-orb instances are not migrated.

Spell data now uses Target=1 and AutoLanzar=1. Tipo=13, Particle=354, duration 60,
mana 20, stamina 5, cooldown 2 and bindability remain unchanged. Mage, Druid and
Bard restrictions and scroll learning remain authoritative. Obsolete floor-light
colour/range/cast-range fields are removed. No terrain target is requested.

The existing UseSpellSlot -> PuedeLanzar -> user-spell handler validates and
consumes costs once; successful casting starts the normal cooldown. The raw-tick
sentinel accepts negative timestamps so cooldowns work beyond 24.9 days uptime.

One entry per user slot holds a versioned owner reference and original deadline.
Recasting refreshes it without a removal packet. Movement, map changes and death
preserve the remaining lifetime. Disconnect/expiry remove it; recycled owner
references are discarded without touching the replacement occupant.

Existing eParticleFX (113), 17 bytes, little-endian:

| Offset | Type | Meaning |
| --- | --- | --- |
| 0 | Int16 | 113 |
| 2 | Int16 | caster character index |
| 4 | Int16 | 354, dedicated orb identity |
| 6 | Int32 | -1 upsert, 0 removal |
| 10 | Boolean byte | false upsert, true removal |
| 11 | Int32 | GRH override 0 |
| 15 | UInt8 | X 0 |
| 16 | UInt8 | Y 0 |

Upserts use the normal living-character area visibility rules. Snapshots are sent
after CharacterCreate in MakeUserChar, including the caster after map changes.
Repeated snapshots are idempotent in HOO and do not refresh server time. Admin
invisibility remains filtered; HOO suppresses the orb and light when invisible.
An invisible caster's snapshot sends SetInvisible before ParticleFX, even though
NotifyUser also sends invisibility later, so split reads cannot reveal the light.
Removal uses the ordinary area route so an observer who died after seeing the
orb can still remove it. Only particle 354 is removed, never all character FX.
The spell no longer emits floor packets 114/121 or changes MapData lights.

The orb regression suite runs through the existing full server UnitTesting
framework. Tests intercept outbound packets only under UNIT_TEST=1, DIRECT_PLAY=0.
The ordinary suite uses a saved/restored spell fixture and does not require an
unmerged resources branch. Before coordinated deployment, run the full test
server as `server.exe --validate-light-orb-resources` with the matching Dat path.
This additionally validates the loaded spell's self-cast target, type, duration,
costs, cooldown, particle and bindability. Old ground-targeted data must fail
this explicit integration check; passing unit tests alone does not validate a
deployment. Do not deploy this server with the old resource definition.
See HOO's `docs/magic-light-orb-spell.md` for visual design and manual acceptance.

Validation (2026-10-14): native Windows VB6 builds passed with DIRECT_PLAY=0,
PYMMO=0/1 and UNIT_TEST=0/1. Focused runs passed 25/25 in both test builds;
the full isolated server suite passed 421/421 in both PYMMO configurations,
including the loaded spell-data contract. No live client/server gameplay test
or production deployment was performed.

Review validation (2026-10-15): all four native build configurations passed
again. Focused suites passed 25/25, ordinary full suites 420/420, and explicit
`--validate-light-orb-resources` full suites 421/421 in both PYMMO modes. A
negative run with AutoLanzar=0 / Target=4 failed only the resource contract,
as intended. The former unconditional assertion reproduced a 420/421 result
with assets master, even though the fixture-based gameplay tests passed.

## Unit-test startup diagnostics

Jenkins build 527 compiled successfully but stayed alive without a final report.
The suite-39 failure was subsequently reproduced as described below.
UNIT_TEST=1 writes `test_progress.txt` beside the
executable, beginning at entry to Main, before configuration or test Init.
Each timestamped checkpoint opens, appends and closes the file, so its latest
line can be read while startup is blocked. Each run starts a fresh progress log.

BEFORE/AFTER checkpoints cover startup calls, including configuration, database
connections/migrations, resource loading, maps, networking, test initialization
and shutdown. Suite indices are logged before and after execution and correspond
to the existing RunSuite registration. Completion includes the new failure count.
The last unmatched BEFORE identifies the operation that has not returned.

Existing modal dialogs are not bypassed: spell, object, EOT, map, quest and
migration errors are logged before their dialogs, including Err.Number,
Err.Description and the preceding checkpoint. Duplicate-instance and anti-cheat
SDK dialogs are also identified. A MODAL entry explains an unattended wait;
it does not claim that the affected operation succeeded.

Unhandled Main errors publish a final FAIL with error details and last checkpoint.
Progress never writes a result. Final reports are written to a temporary file,
closed, then renamed to `test_results.txt`; no partial PASS is published.
Keep `test_progress.txt` when collecting a failed run, before workspace cleanup.
If even the BEGIN entry is absent, investigate execution/file permissions or a
failure before Main rather than assuming a test-suite failure.

All tracing is compiled out with UNIT_TEST=0. Initialization, anti-cheat policy,
test selection and orb gameplay are unchanged. No Jenkins changes are included.
Retain this instrumentation until the Jenkins startup failure is understood.

Diagnostics validation (2026-10-15): native builds passed with PYMMO=0/1 and
UNIT_TEST=0/1. The exact Jenkins override `DEBUGGING=1:PYMMO=0:UNIT_TEST=1`
passed 420/420 using the updated local assets at d7d21868, example configs,
isolated paths/binding and legacy anti-cheat disabled. All 41 suite boundaries
were verified. Missing configuration produced a final FAIL with error 76 and
the previous checkpoint. A deliberately invalid spell value logged error 6
(Overflow) before its existing modal dialog, with no premature final report.
That initial run did not reproduce the remote hang; the shorter spell table
reproduction below subsequently reproduced its exact exception sequence.

## Suite 39 overflow investigation

With commit `0ec619b4` and an isolated `Hechizos.dat` declaring
`NumeroHechizos=351`, the native Jenkins configuration reproduced `ERROR suite
39: Err.Number=6`, `ERROR Unit test runner: Err.Number=6`, and a final
`BEFORE UnitTesting.RunTestError` checkpoint with no results file.

The first invalid operation was `savedSpell = Hechizos(352)`, raising error 9.
The fixture's error handler resumed cleanup, whose unconditional
`Hechizos(352) = savedSpell` raised error 9 again. Each pass recorded another
failure and resumed the same cleanup. Eventually `TotalTests = TotalTests + 1`
overflowed its 16-bit Integer at 32767. Outer handlers retried the same failing
recorder, so final publication was never reached. This is not an orb timestamp
or packet-width overflow. The next CI trace records the loaded spell upper bound
to confirm its actual resource table; Jenkins' asset contents were not available
for direct inspection.

Ordinary lifecycle tests now extend the saved spell table only for their fixture,
run the original assertions and restore the original bounds. Cleanup runs once,
restores only initialized snapshots and reports restoration errors without
retrying cleanup. The explicit deployment-resource check still fails for missing
spell 352 or incompatible authored data. No production resource loading changed.

The separate Init error 9 came from encrypting the test token with an empty key:
PYMMO=0 does not call LoadPrivateKey, cipherEncryptBytes returns an empty array,
and cnvB64StrFromBytes dereferences abData(0). Only the PYMMO=0 test fixture now
uses a fixed test key, without changing the production PrivateKey. A regression
checks the initialized token's encryption/decryption roundtrip. Genuine Init
failures are recorded after runner initialization rather than erased by a reset.

Runner counters and report indices are Long. A regression injects an exception
with counters at 32767, exercises failure/success recording and verifies a closed
FAIL report before restoring the real suite counters. Main unwinds its active
test exception handler before recording and publishing the failure. Startup and
suite diagnostics remain enabled in test builds; none of these fixes change
UNIT_TEST=0 orb gameplay, protocol, costs, cooldowns or bindability.

Native validation of the fix (2026-10-15):

- Exact Jenkins override: `VB6.EXE /make Server.VBP /d
  DEBUGGING=1:PYMMO=0:UNIT_TEST=1 /out build.log /outdir <isolated-build>`:
  build succeeded and the complete suite passed 424/424.
- Explicit `DEBUGGING=1:UNLOCK_CPU=1:UNIT_TEST=0/1:PYMMO=0/1:
  DIRECT_PLAY=0:BATTLESERVER=0`: all four native builds succeeded; full test
  runs passed 424/424 in each PYMMO configuration.
- Short-table reproduction (351 spells): all 424 tests passed, including all
  original orb assertions and the setup-exception/restoration regressions.
- Strict resource validation with that short table: 424 passed, 1 expected
  failure, final `RESULT: FAIL`, remaining suites completed and process exited.
- Strict resource validation with the installed ground-targeted definition:
  424 passed, 1 expected contract failure in each PYMMO configuration.
- With spell data from matching assets PR #669 (`faa2db4cc`), strict resource
  validation passed 425/425 in each PYMMO configuration.

All runs used private copied resources, configuration, database and loopback test
port, with legacy anti-cheat disabled to match Jenkins 527. No live multiplayer
validation or deployment was performed. A fresh Jenkins run is still required;
retain the progress log, especially the loaded spell upper bound and suite 39.
