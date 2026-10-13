# Hunter's Traps 2.0: authoritative lifecycle (Phase 1)

This is a required server/HOO protocol change, not a negotiated capability.
The retired VB6 client is not supported. This server-only phase adds no HOO
implementation, assets, database changes, or gameplay balance changes.

## Packet 192

`ServerPacketID.eUpdateTrap` keeps ordinal **192**, but replaces its former
5-byte layout with this **fixed 16-byte** message. All multibyte fields are
little-endian. There is no string, length prefix, owner ID, or account identity.

| Offset | Bytes | Type | Field |
| --- | --- | --- | --- |
| 0 | 2 | Int16 | packet_id = 192 (`C0 00`) |
| 2 | 1 | UInt8 | event_type: 1 Upsert, 2 Triggered, 3 Removed |
| 3 | 4 | Int32 | trap_id: positive EOT UniqueId |
| 7 | 2 | Int16 | map_id: server logical map |
| 9 | 1 | UInt8 | x: server tile |
| 10 | 1 | UInt8 | y: server tile |
| 11 | 1 | UInt8 | trap_type: 1 Paralysis, 2 Explosive, 3 Poison |
| 12 | 4 | Int32 | remaining_ms: nonnegative on Upsert, otherwise 0 |

Example: Upsert ID 16909060 (`0x01020304`), map 4660 (`0x1234`), tile
(99,88), Explosive, remaining 84281096 (`0x05060708`):

```text
C0 00 01 04 03 02 01 34 12 63 58 02 08 07 06 05
```

The fixture uses deliberately distinctive values to verify byte order. Normal
lifetimes come from EOT configuration, not this example. No packet IDs shift;
the packet has the same ordinal/layout with `PYMMO=0` and `PYMMO=1`.

## Events and ownership

- **Upsert** publishes a completely initialized trap: ID, owner reference,
  effect definition, elapsed time, lifetime, subtype and position. The packet
  omits the owner reference. Successful placement consumes the item and starts
  its existing cooldown; rejected placement does neither.
- **Triggered** is emitted exactly once after an authorized NPC activation.
  Repeat calls, PvP targets, invalid owners, denied attacks and paralysis-immune
  NPCs do not activate or emit this event. Existing damage, poison, paralysis,
  sound and FX calls remain responsible for gameplay/presentation; the new
  event does not execute them again.
- **Removed** terminates the same ID on expiration, fourth-trap replacement,
  death, invalid owner generation or other EOT removal. Cleanup is idempotent.
  A successfully triggered trap also emits Removed when EOT cleanup runs;
  clients must not replay a trigger animation on Removed.

Owner references validate both the existing generation and captured user ID, so
a recycled user slot cannot activate the old user's trap or receive its cooldown
or ownership notifications. Retirement detaches by object identity before reuse.
If the original owner cannot be verified, the retired instance is discarded,
never returned to the reusable pool: an obsolete owner list may still hold it.
That list can safely finish cleanup without touching a newer trap on the tile.
Ordinary retirement of a verified owner's trap still uses the pool.
IDs come from the existing global EOT counter, skip zero and wrap from maximum
positive Int32 to 1. They are not persisted and are scoped to a server session;
clients must discard old state on reconnect. A full counter wrap can reuse a
historical ID and is not a globally unique/persistent identifier.

## Lifetime, snapshots and ordering

`remaining_ms = max(configured TickTime - elapsed EOT time, 0)`. Removed and
pending-removal traps report zero and are excluded from area snapshots.
This uses the existing server EOT clock, including its 200 ms per-update hitch
clamp, not an independently introduced wall clock. No configured duration changes.

`ModAreas` sends `SendSnapshot` on each actual resident trap in its existing
x-then-y tile traversal. Late observers receive the original ID and current
remaining lifetime. Repeated snapshots are Upserts, not new trap instances:
HOO should update the existing entry by ID and never replay placement solely
because an already-known ID is repeated. Expiry estimates are visual only;
authoritative activation/removal comes from the server.

Broadcasts use the existing position-based area routing. The single-threaded
server sends placement Upsert before any activation, Triggered before its
cleanup Removed, and replacement Removed before the fourth trap's Upsert.
Already inactive EOT entries do not count toward the three-active-trap cap.
Delayed cleanup cannot clear a newer trap occupying the same tile.

Leaving an area does not manufacture a Removed event. A future HOO consumer
must handle its own map/session visibility lifecycle; snapshots do not constitute
a global list of every trap on the server. A client receiving a terminal event
for an unknown ID should tolerate it without inventing ownership or gameplay.

## Owner-only packet 208 (PYMMO=1)

`ServerPacketID.eHooOwnedTrap` is appended after mandatory taming result 207.
Reserved 206, taming 207 and every earlier live packet keep their IDs. The
supported HOO deployment and validation configuration is **PYMMO=1**. This is
not an owner field added to packet 192: the area packet remains exactly 16 bytes.

Packet 208 is also fixed **16 bytes**, little-endian:

| Offset | Bytes | Type | Field |
| --- | --- | --- | --- |
| 0 | 2 | Int16 | packet_id = 208 (`D0 00`) |
| 2 | 1 | UInt8 | event_type: 1 Upsert, 2 Removed, 3 Reset |
| 3 | 4 | Int32 | trap_id, positive except Reset |
| 7 | 2 | Int16 | server logical map_id |
| 9 | 1 | UInt8 | tile x |
| 10 | 1 | UInt8 | tile y |
| 11 | 1 | UInt8 | type: 1 Paralysis, 2 Explosive, 3 Poison |
| 12 | 4 | Int32 | nonnegative remaining_ms on Upsert; otherwise 0 |

Reset has **all bytes after event_type zero**. Example owner Upsert, using the
same distinctive fields as the area example:

```text
D0 00 01 04 03 02 01 34 12 63 58 02 08 07 06 05
```

Only `SendTarget.ToIndex` to the validated owning user is permitted. No owner
identity, account ID or user-slot number appears in either packet. No capability
negotiation is added. Invalid generations/identities receive no ownership data.

Ordering on the single server thread is:

- Placement: area 192 Upsert, then owner 208 Upsert.
- Activation: area 192 Triggered, then owner 208 Removed immediately. Later EOT
  cleanup emits area 192 Removed, but **not another owner removal**.
- Expiry, death, replacement or explicit removal: area 192 Removed, then owner
  208 Removed if the original owner is still valid, exactly once.
- Fourth placement: old trap removals precede new trap Upserts.
- Each `WriteChangeMap`, including initial map entry: map notification, owner
  Reset, then active owned-trap Upserts in owner EOT-list order. This full-state
  synchronization includes traps on other maps/outside the visible area. No
  periodic ownership polling or fabricated trap instances are used.

HOO must keep a separate owned-trap set keyed by `trap_id`. Upsert replaces that
ID's metadata/timer; Removed erases it idempotently; Reset clears the set before
the following snapshot entries. Do not clear owned state merely because an area
snapshot no longer contains a trap. Count the entries, not nearby area events.
For presentation, interpolate `max(remaining_ms - elapsed_since_receipt, 0)`;
never infer activation, ownership or server expiry from a visual timer. The
server EOT hitch clamp means presentation can reach zero before actual removal.
Ordered reliable delivery supplies snapshot ordering; there is no completion
marker and the set is rebuilt as its at-most-three Upserts arrive.

Disconnect clears client state. A new session receives a Reset on initial map
entry. Invalid-owner retirement only broadcasts the area removal; it never
notifies or edits the newer occupant of that user slot. Traps remain session
effects, not persisted ownership records.

## Correlated item-use result: packet 209 (Phase 3.1)

`eHooTrapUseResult` is appended after 208 without renumbering existing packets.
It is **209 with PYMMO=1**, and 210 with PYMMO=0 because of the existing
conditional account-character-list entry. HOO uses PYMMO=1 only. This is a
required message, not a capability, fallback or changed client request layout.

Exactly **9 bytes**, little-endian, including the ID:

| Offset | Bytes | Type | Field |
| --- | --- | --- | --- |
| 0 | 2 | Int16 | packet_id = 209 (`D1 00`) |
| 2 | 4 | Int32 | original positive UseItem/UseItemU request counter |
| 6 | 1 | UInt8 | original inventory slot, including invalid slots on rejection |
| 7 | 1 | UInt8 | source: 0 inventory click, 1 keyboard/hotbar |
| 8 | 1 | UInt8 | result: 0 Rejected, 1 AcceptedForTargeting, 2 NotATrap |

Example: accepted keyboard Use, counter `0x01020304`, slot 7:

```text
D1 00 04 03 02 01 07 01 01
```

`HandleUseItem` and `HandleUseItemU` retain their existing counter/security
validation, then call `UseInventoryRequest`. Every ordinary well-formed request
reaching that boundary receives exactly one result, including invalid/empty
slots, rejected restrictions, cooldowns, use intervals and handled errors.
Recognized ordinary items receive NotATrap (whether or not their use succeeds).
HOO ignores outcomes not matching an active trap request. Malformed/replayed
traffic still follows the original security/disconnect policy, not normal Use
acknowledgement. Results go only to the validated requesting user generation/ID.

The existing `UseInvItem` validates the item once. Only its successful trap
target-initiation branch reports readiness; it defers skill 25 to the caller.
The caller sends **AcceptedForTargeting, then WorkRequestTarget(25)** immediately
on the same single-threaded ordered connection. No other targetable item family
is deferred. Rejection sends no trap target request and leaves no newly armed
trap targeting state. Activation still consumes no item and starts no item
cooldown; the existing successful placement path does both.

HOO retains counter, source, original slot/item, map/session and character
identity. It must match all of these before authorizing the following skill 25;
that shared skill alone proves nothing. Rejected/NotATrap resolves only the
matching pending trap Use and allows retry under normal client rate limits.
Old or duplicate replies cannot cancel a newer request. Cancellation/ten-second
timeout must clear authorization, not leave a permanent unanswered-reply guard.
There is no new cancel packet: the next Use replaces old server trap intent,
and existing death/map/completion/rejection cleanup remains authoritative.
The result confirms targeting permission only, never trap placement/ownership.
Inventory updates and packets 192/208 remain authoritative for those changes.

## Inventory targeting and preserved mechanics

The pending item target captures slot, object definition and originating map. On the delayed
WorkLeftClick response, bounds, quantity, unchanged definition and targetable
object type are checked before indexing item metadata. A trap additionally
requires the trap item subtype and a valid trap EOT definition. Invalid requests
clear the pending item target; a completed request cannot submit twice.
`ClearPendingTargetableItem` clears only these three fields, not unrelated
inventory selection. Completion, cooldown/health/stamina rejection, invalid
slot/definition, placement rejection and errors all cancel the pending request.
The WorkLeftClick handler also clears it on sequence/vision/stun/dead/rest/bounds
rejection. Death and map notifications invalidate pending requests, including
leaving and returning to the original map. A new Use is required after rejection.
Resurrection, castle spawners, harpoons and hand cannons retain their dispatch
paths; no trap-only eligibility rule is imposed on those item families.

Existing configured placement range, cooldown values, item quantity consumed,
three-active-trap cap, PvE authorization, NPC immunity, death cleanup and effect
calls remain in place. Equal-age replacement deterministically selects the first
active trap. Item identity here means the definition in the pending slot, not a
new per-item GUID system.

Reference-only asset audit: `Recursos/Dat/EffectsOverTime.dat` defines trap EOTs
6 (Paralysis, 10000 ms), 7 (Explosive, 120000 ms), and 9 (Poison, 120000 ms,
applying EOT 8). Runtime configuration remains authoritative and was not edited.

## Coordinated deployment

1. Keep this server PR in draft until the separate HOO decoder/consumer is ready.
2. Implement HOO's exact 16-byte decoders for **192 and 208**, required nine-byte
   **209** Use result/correlation, the separate owned
   set/reset semantics, and idempotent ID-based lifecycle in
   that follow-up. There is no old-layout autodetection or capability fallback.
3. Validate matching development client/server builds before production rollout.
4. During a coordinated maintenance window, disconnect old sessions and deploy
   the matching HOO and server versions together. Do not expose the changed
   server to a client expecting the old layout or lacking packet 192/208/209 support.
5. Roll back client and server together if necessary; no data migration is needed.

HOO PR #666 contains the matching Phase 3.1 consumer, stacked on Phase 2 #663.
Keep this PR in draft pending coordinated review and matching-build gameplay
validation. No VB6 client or retired ProtocolGenerator changes are required.

## Validation

`Codigo/Tests/Unit_Traps.bas` is registered in `Server.VBP` and the existing
`UnitTesting.RunAllSuites` runner. Build with `UNIT_TEST=1`, `DIRECT_PLAY=0`
(Aurora), then use the normal server unit-test startup. Results appear in the
existing `test_results.txt` report. There is no separate test project or script.

For a focused run without database startup, listeners or installed resources,
the same test-enabled server executable also accepts:

```text
server.exe --test-traps
```

This calls the same registered suite and `UnitTesting` result reporting. The
24 groups cover exact area/owner wire bytes, late/repeated snapshots, all three real
activations, rejection/immunity, expiry, owner invalidation, death, replacement,
stale inventory, placement guards, pooled state/time clamping and positive ID
wrap, orphaned-object reuse across generations/identities, ordinary safe pooling,
remote owner snapshots, owner-only routing/event ordering, terminal rejection,
other targetable item families and actual WorkLeftClick handling/map cancellation.
Tests use the real server types, trap class, effects, inventory and writer.
Only outgoing packets are captured instead of sent. Test-only EOT helpers
isolate the active list and advance its existing update routine deterministically.
Idle user/NPC slots, touched map tiles, configuration and effect definitions are
saved/restored. Capture and helpers are excluded from `UNIT_TEST=0` binaries.
Live matching-client combat/area delivery remains an integration check.

Phase 3.1 native validation (2026-10-13):

- `PYMMO=0` and `PYMMO=1`, each with `UNIT_TEST=0` and `UNIT_TEST=1`,
  `DIRECT_PLAY=0`, `DEBUGGING=0`: all four builds passed.
- `server.exe --test-traps`: **33/33** groups passed in each configuration.
  Nine added groups cover both source layouts/ordinals, acceptance-before-25,
  cooldown rejection/retry, use intervals, dead/stunned/class/level/map
  restrictions, invalid/empty inventory, ordinary/replaced items, handled errors,
  replay security and exactly-once acknowledgements. Existing 24 groups pass.
- `server.exe --test-managed-config`: **106/106** checks passed in each.
- Full suites attempted with each build in a disposable runtime, synthetic key,
  empty test database and loopback-only port. Both crashed before producing a
  report: `sqlite3odbc.dll`, `0xc0000005`, fault offset `0x00010d3c` (PYMMO=0
  also reported `0x0000a247`). This matches the previously documented local
  driver blocker below; full-suite success is not claimed.
- No live matching-client/server gameplay validation was performed.

Suite registration has no conditional holes: network suite 33 is selected for
the transport, then all shared suites; Aurora adds trap suite 41. DirectPlay has
40 suites and does not claim to execute the Aurora-only packet-capture suite.
Missing suite registrations/exceptions/failure returns are reported as failures.
Individual trap failures also make `test_suite_traps()` return False.

Native Windows review validation (2026-10-13, separate temporary directories):

- PYMMO=1, DIRECT_PLAY=0, UNIT_TEST=0 and 1, DEBUGGING=0: builds passed.
- PYMMO=1, DIRECT_PLAY=1, UNIT_TEST=1: build passed (runtime not exercised).
- `server.exe --test-traps`: 24 groups passed, 0 failed.
- `server.exe --test-managed-config`: 106 checks passed, 0 failed.
- Conditional registration checked: 41 contiguous Aurora suites, 40 contiguous
  DirectPlay suites; no missing/duplicate indices.
- Full suite attempted in a disposable runtime (copied resources, empty test
  database, synthetic crypto key, loopback port 17667). Startup crashed inside
  `sqlite3odbc.dll` with `c0000005` before producing a test report. A separately
  built pristine `origin/master` (`08b447a722438ca061ca743ae9b99c6f6efa55cb`)
  reproduced the same driver failure, offset `00010d3c`, with the same setup.
  The full suite is **blocked**, not passed; no existing database was modified.
- Matching-HOO live gameplay/socket delivery has not been validated.

Commands:

```text
VB6.EXE /make Server.VBP /d DEBUGGING=0:UNLOCK_CPU=1:UNIT_TEST=<0|1>:PYMMO=1:DIRECT_PLAY=0:BATTLESERVER=0 /out <log> /outdir <temporary-output>
VB6.EXE /make Server.VBP /d DEBUGGING=0:UNLOCK_CPU=1:UNIT_TEST=1:PYMMO=1:DIRECT_PLAY=1:BATTLESERVER=0 /out <log> /outdir <temporary-output>
server.exe --test-traps
server.exe --test-managed-config
```

Manual follow-up with matching HOO, not claimed by this server-only patch:

- Place each subtype and inspect real ID/map/type/lifetime; observe an existing
  trap from a second client late and repeatedly without duplicate visuals.
- Activate once with an eligible NPC; verify existing damage/status, sound and
  particles occur once. Try immune, protected and player targets without activation.
- Place a fourth trap, let traps expire, die and disconnect/reuse an owner slot;
  verify matching Removed IDs and no stale tile or cooldown ownership.
- Change/empty/reuse the targeted inventory slot before clicking; confirm no
  wrong-item trap, item loss, extra cooldown or runtime error.
- Check development server logs and matching-client parsing across map/area
  transitions. Do not use an old client to assess the new wire format.
