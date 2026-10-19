# Server network transport

Aurora.Network is the server's only supported transport. The retired DirectPlay
implementation, its compiler switch `DIRECT_PLAY`, settings, classes, tests,
DirectX type-library reference and bundled DX8VB.DLL have been removed.
Historical dated validation records and the retired-client command in Agents.md
are not supported server build instructions.

## Dependencies and configuration

Keep Aurora.Network.dll, its existing COM reference and RegistrarAurora.bat.
All unrelated COM references are unchanged. There is no replacement transport,
new dependency or protocol migration. Existing configuration files may still
contain a `[DIRECTPLAY]` section, but the server no longer reads its keys; the
section is absent from the shipped template and can be removed by operators.
No user configuration or database is edited by this change.

## Preserved behavior

The former Aurora branch is retained without changing admission, connection
mapping, callbacks, packet accounting, throttling, authentication, dispatch,
fan-out, visibility, buffer clearing, send ordering, disconnect diagnostics or
timeouts. Auxiliary writer initialization remains before castle initialization
and network startup; the main loop still ticks the same network server.
PYMMO, UNIT_TEST, DEBUGGING, LOGIN_STRESS_TEST and UNLOCK_CPU remain independent
compiler options. No packet IDs, fields or encoding have changed.

The existing runner has 41 contiguous suites: Aurora is 33, orbs 39, taming 40
and traps 41. All modern test bodies and startup/failure reporting remain.
Packet capture and test helpers remain under UNIT_TEST=1 and are absent from
production. Binary-buffer helpers used elsewhere are retained.

## Native validation commands

Use Windows VB6, registered existing dependencies, and separate output paths.
Substitute each PYMMO and UNIT_TEST combination from 0/0, 1/0, 0/1, 1/1:

```text
VB6.EXE /make Server.VBP /d DEBUGGING=<0|1>:PYMMO=<0|1>:UNIT_TEST=<0|1>:UNLOCK_CPU=1:BATTLESERVER=0 /out <build.log> /outdir <output>
server.exe --test-traps
server.exe --test-managed-config
server.exe
```

Use DEBUGGING=0 for production and DEBUGGING=1 for diagnostic test runs.
The last command runs the full existing suite in test builds. Use a disposable
database, complete resource fixture and a loopback-only unused port. Inspect
test_results.txt and test_progress.txt; an incomplete run is not a pass.

## Manual integration checklist

With a matching HOO client and local test server, verify startup/listening,
login, ordinary packet exchange, logout, reconnect, character switching and
clean shutdown. Check disconnect diagnostics and startup logs. Do not use
production accounts or data. Reverting the transport-removal commit restores
the prior build layout if necessary; no database/protocol rollback is needed.

## Validation recorded 2026-10-19

Base: `5d8135929e904863e551ede55f707438efe25191` (includes the current trap,
orb, taming ownership-refresh and startup/test-runner fixes).

Native compiler: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`.
All builds used the command above without any transport compiler argument,
UNLOCK_CPU=1 and BATTLESERVER=0. No COM registration or dependency change was
needed. Existing local COM path corrections were used but are not in the PR.

| PYMMO | UNIT_TEST | DEBUGGING | Native build | Full suite |
| --- | --- | --- | --- | --- |
| 0 | 0 | 0 | Passed | Not compiled |
| 1 | 0 | 0 | Passed | Not compiled |
| 0 | 1 | 1 | Passed | 427 passed, 0 failed (1869.01 ms) |
| 1 | 1 | 1 | Passed | 427 passed, 0 failed (971.99 ms) |

In both test configurations, focused traps passed 33/33 and managed
configuration passed 106/106. Full-run checkpoints confirm all 41 suites
completed, including Aurora serialization/packet validation, orb lifecycle,
taming ownership refresh and trap packet/lifecycle tests. Startup and the
normal test shutdown completed against disposable databases and local resources.

Both production binaries passed loopback smoke tests: startup/listening,
three connections with exact `02-00-01-00` framed greetings, invalid packet
rejection/disconnect, reconnect and normal form-driven shutdown. Listeners
closed, and neither production binary ran the test subsystem. Authenticated
HOO login, character switching and multiplayer gameplay were not exercised.

Static comparison of 11 retained source files against the previous Aurora
branch found identical code after resolving only the transport guards (ignoring
comments/blank lines). All compiler blocks balance, project source references
exist, and suites 1 through 41 are contiguous. PacketId.bas and binary-buffer
helpers are unchanged. `git diff --check` passed.

The tracked legacy-symbol search has no matches in active source, project,
configuration or scripts. Remaining mentions are explicitly historical records
in HunterTrapsProtocol.md, MagicLightOrb.md, TamingOwnershipRefresh.md, the
retired-client command in Agents.md, and this removal documentation.

Local evidence: `Logs/aurora-only/pymmo*-unit*/` contains build logs and
test reports/progress; `Logs/aurora-only/smoke-final.log` records the two socket
smoke runs. These local fixtures, executables and logs are not shipped.

## External CI follow-up

Read-only audit of `ao-org/ci-cd` at
`ecde199e101a0ce285f2b2bd4d436ac360cf445c` found one obsolete reference:
`heroes-of-old/Jenkinsfile.server-deploy`, line 53, appends `:DIRECT_PLAY=0`
to the server build definitions. Remove that argument in a separate CI change.
No external repository or Jenkins configuration was modified here. Jenkins
job definitions not stored in that repository were not accessible for audit.
