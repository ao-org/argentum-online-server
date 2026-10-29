# In-game ordinary housing

This feature uses the authenticated VB6 world connection. PyMMO retains its
existing pregame first-claim operation. There is one authoritative VB6 world
process for the database; multiple PyMMO login processes are supported. External
ownership editors must not operate on a running world. This is not cross-world-
process key synchronization.

## Ownership and minimum period

The consolidated `house_key` catalogue is authoritative. `key_obj` identifies a
key/catalogue row; NULL `account_id` means available. Keys referenced by
`castle.inside_key_obj_id` are excluded. No catalogue rows are deleted.

Acquisition requires current account entitlement: Heroe (6057394), Leyenda
(6057395), Noble (28929328), or Emperador (28929365). VB6 does not revoke existing ownership when entitlement expires. Release remains
possible after the minimum period; the separate PyMMO expiry writer is discussed below.

The minimum ownership period is **72 hours from `house_key.assigned_at`**.
Release and exchange compare SQLite UTC `julianday('now')` with
`julianday(assigned_at,'+72 hours')`, inclusively, in the same `BEGIN IMMEDIATE`
transaction as the mutation. Claim/exchange writes `CURRENT_TIMESTAMP` to the
new house. Release preserves the old timestamp. Reads, rejected operations and
replayed requests never refresh it. An account with no ordinary house can claim
normally. There is no account cooldown, new timestamp column, or migration.
Invalid timestamps and multiple ordinary assignments are reported, never repaired.

## Transaction and live state

`modHousingDatabase` owns a dedicated ADO connection to the configured world
SQLite database. Parameterized commands and explicit SQL transactions avoid the
rotating asynchronous character-save connections. The installed x86 SQLite ODBC
stack is validated separately; `Timeout=0` is not fail-fast on that driver.
Housing uses `Timeout=1` and `PRAGMA busy_timeout=1` only on its own connection.
Lock contention returns Busy. This is a synchronous bounded operation; it does
not yield, use DoEvents, or retain transactions between loop iterations.

Inside the write transaction, ownership, entitlement, the expected old key,
minimum period, target availability, interior/threshold occupancy and every
loaded account character's key capacity are checked before writes. Exchange
releases A and assigns B atomically with exact affected-row checks. Rollback does
not apply prepared key changes. A commit error is not represented as confirmed
rollback: authoritative ownership is reconciled or ordinary-key access is blocked
until a successful status/login reconciliation. The client must read status and
must not automatically replay the mutation.

Following commit, all loaded characters (including retained disconnected users)
lose every released ordinary key and receive the acquired key. Castle and other
keys remain. Authorization changes precede packet notification, and empty slots
are sent. Physical inventory keys and both keyring/HOO key routes consult the
in-memory ordinary ownership cache, without per-door SQL. Ordinary unlocked-door
interaction remains available to guests.

Before admitting a login, authoritative keys are refreshed. A former owner or
visitor saved inside a mapped ordinary interior returns to their validated home
city in memory, using existing placement validation. Normal later character saves
persist that position. No character-save SQL or `is_logged` behavior changes.

## Explicit region coverage

`HousingRegions.csv` is required beside the server executable. Runtime map and
resource identities are separate columns and must match the loaded world. `D`
rows identify door anchors and the locked variant; `R` rows are explicit interior
tile runs. There is no runtime flood fill, map-range heuristic, or name matching.

Initial verified coverage:

| Key | Main map / door | Main interior | Basement resource |
| --- | --- | --- | --- |
| 54 | 1 / (60,23) | x54–61, y18–23 | 603, explicit top-left room |
| 65 | 1 / (50,23) | x44–51, y18–23 | 602, explicit top-left room |
| 69 | 1 / (40,23) | x34–41, y18–23 | 601, explicit top-left room |
| 73 | 1 / (45,40) | x39–46, y35–40 | 600, explicit top-left room |

Each basement contains several independent properties: the entire basement map
is **not** assigned to one house. Audited room exits return through the existing
return-position mechanism. The offline audit checked directional walk boundaries,
all doors, both interior exits and incoming links. Runtime validation checks map
identity/bounds, non-overlapping interiors, door variant/lock consistency and
coverage of every installed door using each mapped key.

Other ordinary houses remain in the catalogue but are not actionable through
this first in-game mapping. Unsafe unmapped release/exchange is Unavailable.
Expand coverage only with independently verified interiors and associated doors.
Occupants (owners, guests and retained logout users) and occupied thresholds block
release/exchange. Successful changes close and lock old doors through existing
collision and broadcast paths. There is no forced evacuation.

## World wire contract

Append-only IDs: client **317**, server **221** with `PYMMO=1`. The build without
PyMMO retains its conditional legacy enum offsets; verify both variants before
release. No ProtocolGenerator or new capability advertisement is involved.

All integers use the existing little-endian legacy encoding and String8 uses the
existing variable-length byte count / Windows-1252 representation.

Request: Int16 packet ID, Int32 positive request ID, UInt8 operation (catalogue=0,
claim=1, release=2, exchange=3), Int16 expected key (0 for none), Int16 target key,
Int16 cursor, String8 search (at most 96 bytes).

Result: Int16 packet ID, Int32 request ID, UInt8 operation, UInt8 code, Int16
current key, String8 description (512 bytes), String8 assigned UTC, eligible UTC,
server UTC (each bounded to 32 bytes), Bool can-acquire, Bool can-release, Int16
next cursor, UInt8 count (0–4), then count pairs of Int16 key / String8 description.
Codes: Okay=0, Busy=1, Unavailable=2, Ineligible=3, Stale=4, Occupied=5,
MinimumPeriod=6, Inconsistent=7, Invalid=8, Unresolved=9. No account ID is accepted
from the client.

One session high-water request ID and cached payload/result prevent duplicates
from executing again. Reused IDs with different contents and older IDs are
rejected. Slot reset and account/version changes clear this state. The 500 ms
request throttle is separate from the 72-hour ownership period. A new client
connection never replays a mutation. HOO uses a bounded response deadline and
requires status before another mutation after an uncertain/lost response.

## Rollout and validation

Deploy the matching world server and HOO together. PyMMO source, authentication,
first claims and payments are unchanged by this PR. Validate the consolidated
schema (nullable house owner), catalogue objects and region resource identities on
a staging copy/disposable database. Do not rerun old consolidation migrations.
Back up deployment configuration through the normal operator process; this feature
does not copy, migrate or edit production data automatically.

Manual matching-client smoke: acquire A, verify all loaded account characters get
the key; verify release/exchange are denied during the ownership period; use a
synthetic mature assignment, exchange A→B, verify A rejects old inventory/keyring
keys, B accepts its key, and A's doors close/lock. Retry with a guest inside and
on the threshold. Relog a former owner/visitor saved inside A and confirm the
home-city position. Restart and repeat. Do not adjust real assignment timestamps
to manufacture tests.

Rollback requires matched client/server versions and the former executable. Existing ownership changes remain
valid consolidated rows, requiring no schema rollback. Inspect unresolved housing
logs and reconcile status before permitting further changes; never blindly replay
an uncertain mutation.

### Audit revisions and resource checksums

Server base `083fd3e39a651cf787b4a47b4c5e478f4b23a3ba`, PyMMO `16a174f203341e6d9b9ca32e8661aa66331d7945`,
HOO `130f42680bcdced77c34296aa0a6f2df449f5b31`; installed resource checkout
`a8dbc41c996b336f562594e52e1cf1df30478ba9`. Exact audited file SHA-256 values:

```
mapa1.csm   34d9517b27401903ffd415ce258eeec9b23097932bf56226a2b83ca0ed80585e
mapa600.csm 5a94c0b08b98e863402c63322e8978bb44fb0bb8a3e0b4c14bad31fa869a3bd2
mapa601.csm 8dfc94b2508ea409c35aa349479d7fdff077dc3847bf100b8e6f4863b8c5ccc6
mapa602.csm d9a8f38de769f643250c2233c4e15f9ca299a6d09d8ad98f02fafd62ee1fa6a4
mapa603.csm ff4f2ddfc56107cd999e5a016dca417f95fd78922e35ce4de7d254b62db8c273
obj.dat    6691b8d166e837f6627eabd97749472fff1f94fe018e530b7a3296d2add47576
```

Verified non-PyMMO wire offsets: request 320 and result 222 (`PYMMO=0`). HOO
uses the existing PyMMO-world variant (317/221); it does not change the separate
legacy account-login variant's packet ordering.

Current external-writer audit found the patron gateway's
`SQLAuthHelper.release_patron_houses` deleting consolidated catalogue rows. The
PyMMO code remains unchanged at the user's request. That background cleanup can
remove ownership independently of the VB6 player release/exchange rule and cannot
reconcile live keys. It remains a rollout limitation, not a guarantee supplied by
this PR. `HousingService._claim` already uses
`BEGIN IMMEDIATE` and sets `CURRENT_TIMESTAMP` only on successful first claim.
Historical documentation mentions a separate game API; its current implementation
was not available in the checked-out repositories/search results. Its ownership
writes must be audited/disabled before rollout; this PR does not claim coordination
with an unverified external writer.

### Validation performed for this change

Native Windows VB6 `/make build-local-housing.vbp /d
DEBUGGING=0:PYMMO=<0|1>:UNIT_TEST=<0|1>:UNLOCK_CPU=1` passed in all four
configurations. The ignored local project preserves the machine's existing COM
reference paths; tracked references are unchanged. Each UNIT_TEST executable ran
`--test-housing`: **26 passed, 0 failed**, including two gated independent PyMMO
processes. Production executables were compiled but not started.

The disposable native ADO probe used MSDASQL.1, 32-bit `sqlite3odbc.dll` 0.9999,
SQLite 3.43.0. Explicit `BEGIN IMMEDIATE`/`COMMIT`/`ROLLBACK`, parameter binding,
`changes()` and independent-connection commit visibility were exercised. Observed
lock waits were about 264 ms with Timeout=50 and 51 ms with Timeout=1. Timeout=0
was not bounded on this installed driver and is not used. Regression tests require
contention to return in under one second. ADO client-side cursors cache the bounded
rows because this driver's forward-only text fields fail when accessed twice.

Coverage includes exact inclusive 72-hour SQL comparison, before-boundary denial,
reconnect/reopen, expired entitlement, duplicate/payload mismatch and slot reuse,
exchange rollback, injected pre-commit failure and lost post-commit reply,
failed-status uncertainty retention, key capacity/cleared slots/castles, occupants,
thresholds, locked-door normalization, login interior classification, and actual
native writer packet export and all three real key-use rejection paths. No synthetic timing result is described as proof of a
production transaction failure. Full resource-dependent world startup and a live
client exchange/relog were not run; follow the staging smoke above before rollout.
No real character/ownership rows were read or modified for these tests.

The native concurrency fixture lives in this server repository at
`Codigo/Tests/fixtures/native_pymmo_housing_peer.py`. Set
`HOUSING_TEST_PYMMO_PEER` to that file, `HOUSING_TEST_PYMMO_ROOT` to an unchanged
PyMMO checkout, and `HOUSING_TEST_PYTHON` to its native Python executable. No
PyMMO companion PR or PyMMO source change is included.
