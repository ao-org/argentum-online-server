# Nonblocking housing admission coordination

This is only world admission/session coordination for PYMMO's optional paid
house exchange. VB6 does not select houses, quote prices, charge credits or
change the game protocol. Existing `house_key` loading is unchanged.

## Ordering

With the explicitly migrated, activated SQLite database:

1. `Protocol.HandleIncomingData` retains packet validation and queues a cloned
   login reader instead of immediately invoking either PYMMO login handler.
2. `modWorldAdmission` asynchronously validates both token values against
   `account_sessions` and claims the account using one `INSERT ... SELECT`.
   The same SQLite writer exclusion protects PYMMO's exchange transaction.
3. After completion, a claim-bound session read revalidates the token. Only then
   does the original login handler continue and load character/house keys.
4. The claim stays resident during admission, play and logout. Retiring the user
   retains its own outstanding asynchronous saves; ADO completion events,
   not dispatch return, establish when database work is finished.
5. Release asynchronously deletes only the exact claim and process instance.
   User generations and connection identity guard stale callbacks.

The runtime coordinator never polls in a wait loop, sleeps, drains saves or
holds a transaction during gameplay. Database work uses `adAsyncExecute`.
Completion callbacks record state; the normal main-loop poll advances it.
Contention rejects admission or retries cleanup on later ticks, without waiting
on the VB6 thread. Existing unrelated synchronous server persistence is not
redesigned by this change.

## Startup, recovery and failure

Migration/cutover belongs to PYMMO's `serverhelper.housing_schema` and requires
all old auth/world processes to be stopped. The default is inactive. Old tokens
are cleared and new legacy-token inserts rejected, so incompatible PYMMO=1
worlds cannot consume coordinated authentication. Do not enable this database
mode for standalone PYMMO=0 or stress-login builds.

Initialization precedes listening and uses no lock-wait timeout. A failed
coordinated initialization cannot be bypassed by the GUI listener restart.
Every process holds an exclusive OS handle beside the real local SQLite file.
Its exact lock path is stored with the instance ID, so database path aliases do
not cause another process to probe an unrelated, unlocked sidecar file.
Only an instance whose handle can be acquired is considered dead; no elapsed
timeout permits eviction. Startup removes only those dead-instance claims.
Other live worlds retain their claims. Do not relocate the database or delete
coordination lock files while a world is running. Network-share SQLite and
MySQL are not supported for this feature.

An uncertain save error retains the claim until the process is restarted. It
must not be interpreted as permission to purchase. Cleanup never deletes by
account ID alone, so a delayed old session cannot release a new session's claim.
Runtime claim/release statements are short autocommit statements. There is no
database transaction spanning play and no new network packet.

## Review fix: claim-specific persistence

Save tracking is keyed by an immutable execution command and its exact world
claim. Account A never waits for B's saves or unrelated SQL. `SaveUser` covers
full/final character saves and token removal; `SaveChangesInUser` now scopes
incremental saves as well. Admission scopes new-character writes. Main character,
inventory, bank, skills, pets, quests, skins and collectible writes inside these
paths all pass through the tracked `Execute` path.

Coordinated bank saves use a fresh command/parameter collection per dispatch;
cached prepared commands cannot overwrite an outstanding execution's identity.
Failure before Execute is entered retires unsubmitted tracking. A COM exception
from inside Execute is conservatively uncertain: remove the orphan tracking
entry but retain a sticky failure on only its owning claim until restart. Late
or duplicate callbacks cannot clear that uncertainty or affect a newer claim.
There are still no runtime waits or synchronous save drains.

## Tests

`PYMMO=1, DEBUGGING=1, UNIT_TEST=1` registers suite 42 after the existing 41
suites. `server.exe --test-world-admission` runs that suite through the existing
test reporting mechanism against a uniquely named disposable SQLite database,
without loading configured accounts or opening a game listener. It tests native
ADO contention responsiveness, token checks, duplicate claims, exact cleanup
identity, save completion and reused connection identity. Production builds
exclude this test entry point and assertions.

Live acceptance still requires matching HOO/PYMMO/VB6 builds: exchange versus
already-issued-token admission, new key loading, final logout save, reconnect,
two clients, interrupted admission and process crash recovery. Keep purchasing
disabled until these gates and the read-only ownership audit are satisfied.

Recorded validation (2026-10-19): Windows VB6 builds passed with `PYMMO=1`,
`LOGIN_STRESS_TEST=0`, both `DEBUGGING=0:UNIT_TEST=0` and
`DEBUGGING=1:UNIT_TEST=1`. The admission suite now has 23 passing assertions. The
full 42-suite run passed 461/461 on a disposable database and loopback fixture.
Review cases include claim-specific waits, definite/uncertain dispatch failure,
duplicate/late completion, newer-session isolation, unrelated SQL, and actual
removal of an unrelated account's claim despite another claim's failed save.
A separate production-build loopback smoke passed on disposable Windows SQLite
state: coordinated startup/listening, a new connection greeting while another
connection held the database write lock, exact-claim cleanup after valid-token
failed character admission, and restart recovery that cleared a dead process's
claim while preserving a live process using its original different lock path.
These are not a claim of a live HOO/PYMMO/VB6 purchase test. PYMMO=0 and MySQL
coordination are not enabled by this feature.
