# Identity-bound NPC bow attempts

This is required HOO/world protocol, not a capability or feature toggle.
Deploy the world-server build first, then the compatible HOO build. Legacy
clients remain unaffected: the server sends no new bow-stream messages to
clients that have not subscribed. The legacy tile-targeted projectile path
remains available for manual player targeting. Spell repetition,
spell-auto state, protocol generators and melee rules are not changed.

## Wire contract

All fields are fixed-width little endian. Identity words are opaque unsigned
32-bit bit patterns stored in VB6 Longs; neither word is a GUID. A pair of zero
words is never issued. Existing packet IDs and layouts are unchanged.

| Direction / message | PYMMO=1 | PYMMO=0 | Bytes including ID |
| --- | ---: | ---: | ---: |
| Client Subscribe | 315 | 318 | 2 |
| Client Attack | 316 | 319 | 24 |
| Server Context | 211 | 212 | 11 |
| Server Binding | 212 | 213 | 21 |
| Server Result | 213 | 214 | 19 |

- Subscribe: `int16 id`.
- Attack: `int16 id, uint32 context_low, uint32 context_high, int32 request_id,
  int16 character_index, uint32 token_low, uint32 token_high`.
- Context: `int16 id, uint32 context_low, uint32 context_high, uint8 phase`.
  Phase 0 begins a barrier; phase 1 marks it ready.
- Binding: `int16 id, uint32 context_low, uint32 context_high, uint8 kind,
  int16 character_index, uint32 token_low, uint32 token_high`.
  Kind 1 binds; kind 2 revokes.
- Result: `int16 id, uint32 context_low, uint32 context_high, int32 request_id,
  uint8 code, int32 retry_after_ms`.

Request IDs are positive and monotonically increasing within a context. Results
echo the request's context, including InvalidContext. Retry delays are nonnegative.
Codes are Accepted=0, RateLimited=1, OutOfRange=2, NoEnergy=3, NoAmmo=4,
InvalidEquipment=5, ActorBlocked=6, InvalidTarget=7, InvalidContext=8,
ForbiddenTarget=9, InternalError=10. Accepted means an attempted shot, including
a miss, not proof of damage.

## Identity and visibility

One unsigned 64-bit counter allocates contexts and binding tokens. It does not
wrap: exhaustion fails closed. The server process cannot retain TCP sessions
across restart. NPC's old wrapping Integer VersionId is an additional check,
never the lifetime authority.

On initial subscription the server sends:

1. Context Begin.
2. For each eligible visible NPC, a recipient-only CharacterRemove, fresh
   CharacterCreate and Binding.
3. Context Ready.

The explicit remove forces a fresh HOO actor generation even if the character
index had been reused before the client subscribed. It does not remove the
server NPC or affect other observers. A binding is issued only after that
recipient's Create was successfully queued. Failed queueing cannot authorize a
target or complete the barrier.

Normal duplicate area creates preserve a still-valid token. A repeated
subscription in an already-ready context replays Ready, without rotating the
context, clearing the request cache or restarting a shot.

Loaded-area checks use the same geometry as ModAreas. Departed bindings are
revoked before AreaChanged, which implicitly prunes HOO actors. Every NPC move
also checks subscribers, even when no area-cell boundary was crossed. Deletion,
NPC allocation/release, invisibility and wild-to-pet transitions revoke bindings.
Attacks independently recheck map, tile occupancy, health, live character index,
ownership, version and token. A creature returning to view gets a fresh token
and does not resume an earlier repeated attack.

Map transitions and death retire identity while retaining subscription intent.
The new barrier follows area setup or resurrection. A Subscribe received while
dead records intent but does not publish an empty ready context. Socket closure,
logout, character replacement and user-slot generation changes clear everything.
An ordinary client-side local-character representation refresh should stop the
selected attack, not discard an otherwise current stream context.

## Attempts, pacing and resources

The new endpoint validates the target, actor, equipment, ammunition, energy,
guild restrictions and UserCanAttackNpc before commitment. It never resolves an
identity request through the last tile cursor. Owned/summoned NPCs are not bound.
Existing authority continues to enforce faction, team, immunity, mounted-state,
consultation, range and attackability rules.

The retry delay is the maximum remaining magic-to-attack, attack-to-magic,
bow and equipped-item cooldown gate, using existing wrap-safe tick arithmetic.
Validation reads those gates without changing them. A committed attempt uses
the existing bow timing update and stamina cost, not a new gameplay timer.

UsuarioAtacaNpc explicitly reports commitment and completion. The new path pays
only after that routine's real early-rejection checks. The shared ranged-shot
helper snapshots target position before combat; killing/releasing a target must
not redirect a projectile to a reused slot. Ranged post-hit visuals skip a
released/replaced NPC. The paired effect callbacks still run: provocation must
restore the attacker's damage modifiers, while ApplyEffectOnHit retains its
existing target-generation guard. Melee keeps its existing path.

Both legacy and identity paths share projectile completion, ammunition updates,
free-consumable-zone rules, weapon sounds and non-ammunition item cooldowns.
Manual legacy rejection timing and projectile presentation remain unchanged.
The identity path also checks the inventory-removal helper's completion result.
An unconfirmed ammunition debit after combat returns cached InternalError, not
Accepted, and cannot be executed a second time by replaying the request.

## Bounded retries and failure

Each context retains 16 exact request/result records and a high-water request
ID. Exact duplicates replay the original outcome without repeating combat or
costs. Altered duplicates and older evicted IDs fail closed. A terminal uncertain
result is recorded before side effects; exceptions cannot re-arm the request.
Callbacks only write a result back if its original context still exists.

A separate transport budget admits at most 32 bow messages per second, including
duplicates and subscriptions. Over-budget messages are ignored without replacing
their original cached result. Once admitted again, a duplicate can still recover
that result. The budget is not a gameplay cooldown.

Bindings are allocated only for subscribed users and are bounded by MaxNPCs
(10,000), the maximum NPC pool. Revoked slots are reused; the replay cache is
fixed-size. There is no unbounded request log or per-frame networking task.

`DEBUGGING=1` builds additionally use the existing `LogInfoServidor` logger for
at most the first four admitted bow messages per user/second. Result diagnostics
include context, request, target character/token, outcome, retry delay and whether
combat committed (identity words are printed as hexadecimal `high:low`).
Cached duplicates retain the original committed marker,
including InternalError after damage or ammunition failures. The marker is
internal metadata, not an added wire field. `DEBUGGING=0` compiles out formatting
and logging; no production per-shot log is introduced.

## Validation

The native test runner registers `Unit_NpcBow` as suite 43. Its isolated entry is:

```text
Server.exe --test-npc-bow
```

Tests exercise real Aurora serialization, ordered Remove/Create/Bind barriers,
failed delivery, repeated creates, generation wrap, visibility/ownership/death
invalidation, exact-payload replay, malformed packets, abuse limits, range and
equipment restrictions, real NPC attack outcomes, resource charging, remaining
cooldown arithmetic and context teardown/recovery.

Native Windows validation (2026-10-20):

| Configuration | Native build | Focused NPC bow | Full existing runner |
| --- | --- | ---: | ---: |
| `DEBUGGING=1:PYMMO=1:UNIT_TEST=1:UNLOCK_CPU=1` | Passed | 53/53 | 532/532 |
| `DEBUGGING=1:PYMMO=0:UNIT_TEST=1:UNLOCK_CPU=1` | Passed | 53/53 | 532/532 |
| `DEBUGGING=0:PYMMO=1:UNIT_TEST=0:UNLOCK_CPU=1` | Passed | Not applicable | Not applicable |
| `DEBUGGING=0:PYMMO=0:UNIT_TEST=0:UNLOCK_CPU=1` | Passed | Not applicable | Not applicable |

All four builds were made with the installed VB6 compiler using
`/make /d <configuration> /out <build-log> Server.VBP`. Focused runs used
`--test-npc-bow`; full runs used the normal startup runner, copied development
resources, a fresh empty development database, and loopback port 17669.
No production database was used. Tests include real killing-hit cleanup,
deterministic misses, final-arrow depletion, free-consumable zones, cached
post-commit damage/ammunition failures, and exact high-bit identity wire bytes.

Live coordinated checks are still required: moving NPCs, final killing shots,
rapid selection changes with delayed replies, map travel, death/resurrection,
NPC respawn and taming during repetition, and manual player-targeted shots with
the new HOO build. Unit tests do not establish live multiplayer validation.
