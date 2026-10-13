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

Owner references include the existing generation check, so a recycled user slot
cannot activate the old user's trap or receive its cooldown-clear notification.
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

## Inventory targeting and preserved mechanics

The pending item target captures both slot and object definition. On the delayed
WorkLeftClick response, bounds, quantity, unchanged definition and targetable
object type are checked before indexing item metadata. A trap additionally
requires the trap item subtype and a valid trap EOT definition. Invalid requests
clear the pending item target; a completed request cannot submit twice.

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
2. Implement HOO's exact 16-byte decoder and idempotent ID-based lifecycle in
   that follow-up. There is no old-layout autodetection or capability fallback.
3. Validate matching development client/server builds before production rollout.
4. During a coordinated maintenance window, disconnect old sessions and deploy
   the matching HOO and server versions together. Do not expose the changed
   server to a client expecting the old layout (or lacking packet 192 support).
5. Roll back client and server together if necessary; no data migration is needed.

The inspected HOO packet coverage still lists 192 as unimplemented with the old
layout. This PR alone is **not deployment-ready**. HOO, VB6 client and retired
ProtocolGenerator are intentionally unchanged.

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
14 groups cover exact wire bytes, late/repeated snapshots, all three real
activations, rejection/immunity, expiry, owner invalidation, death, replacement,
stale inventory, placement guards, pooled state/time clamping and positive ID
wrap. Tests use the real server types, trap class, effects, inventory and writer.
Only outgoing packets are captured instead of sent. Test-only EOT helpers
isolate the active list and advance its existing update routine deterministically.
Idle user/NPC slots, touched map tiles, configuration and effect definitions are
saved/restored. Capture and helpers are excluded from `UNIT_TEST=0` binaries.
Live matching-client combat/area delivery remains an integration check.

Full native server build matrix (separate temporary output directories):

```text
VB6.EXE /make Server.VBP /d DEBUGGING=1:UNLOCK_CPU=1:UNIT_TEST=<0|1>:PYMMO=<0|1>:BATTLESERVER=0 /out <log> /outdir <temporary-output>
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
