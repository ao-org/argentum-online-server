# Inventory item destruction

Players explicitly prepare and confirm an exact quantity from one ordinary
inventory slot. This permanently removes that quantity without rewards or a
ground object. Drop bindings, `Intirable`, `Instransferible`, `Newbie`, the
legacy DeleteItem packet and GM ground-object DestroyItems command are unchanged.

## Deployment and capability

Deploy the server before the matching HOO client. The server advertises
`hoo-inventory-destruction-v1` through the existing client toggles packet and
accepts capability bit `0x40` in existing capability protocol version 1. Both
prepare and confirm require that negotiated bit. **Do not reuse `0x10`: that
bit is reserved for the retired orb extension, and `0x20` is reserved for
retired taming-result negotiation.**

This feature is an explicitly requested exception to the general policy against
new capabilities. ProtocolGenerator remains retired: its companion change only
documents and validates the additive ID contract, without regenerating templates
or outputs (the legacy randomizer would renumber live IDs). The contract must
remain synchronized with the server and coordinated HOO changes.
Unrelated packets and capability gates are unchanged. No database changes exist.

## Wire format

All numbers are little endian. No strings or array length prefixes appear.

| Message | PYMMO=1 ID | PYMMO=0 ordinal | Total bytes |
| --- | ---: | ---: | ---: |
| Client prepare | 313 | 316 | 20 |
| Client confirm | 314 | 317 | 18 |
| Server result | 210 | 211 | 40 |

Prepare, in order:

| Field | Type |
| --- | --- |
| Packet ID | int16 |
| Positive request ID | int32 |
| Inventory slot | uint8 |
| Expected object ID | int16 |
| Expected original stack amount | int16 |
| Expected equipped | uint8, 0 or 1 |
| Expected elemental tags | int32 |
| Exact quantity | int32 |

The expected tags are the client-visible union of stored tags and natural object
tags, as sent by ChangeInventorySlot. The server separately retains the raw stack
snapshot and the visible tags. Quantities are validated as Long against the
positive current stack amount before narrowing to Integer.

Confirm consists of int16 packet ID followed by exactly 16 opaque token bytes.
The server creates the token with Windows `CoCreateGuid`.

Result, in order:

| Field | Type |
| --- | --- |
| Packet ID | int16 |
| Original request ID | int32 |
| Opaque token | 16 bytes |
| Stage: 0 prepare, 1 confirm | uint8 |
| Result code | uint8 |
| Original slot | uint8 |
| Original object ID | int16 |
| Original stack amount | int16 |
| Original equipped state | uint8, 0 or 1 (not VB6 Boolean) |
| Original client-visible elemental tags | int32 |
| Exact quantity | int32 |
| Expiry seconds | int16 |

Code 0 means prepared at stage 0 or removed at stage 1. Codes 1..13 are:
Unsupported, InvalidRequest, Unavailable, Dead, Busy, ActiveBoat, ActiveMount,
EventObjective, StackChanged, Expired, InvalidToken, RateLimited, InternalFailure.
Prepare success expires after 30 seconds; all other results carry expiry zero.
An unknown confirm token is echoed with request ID zero and InvalidToken.
Known-token results keep the original request identity. Error item fields may be
zero. Invalid nonpositive prepare IDs are discarded without mutating inventory.

## Session, stale-stack and duplicate safety

One bounded destruction state lives in `t_User`, never in a database. It stores
the exact character ID and user-slot generation, stack snapshot, quantity,
wrap-safe preparation tick, mutation latch, token and last confirm result.
Preparing any replacement invalidates the previous token, including a rejected
replacement. Preparation is limited to one per 500 milliseconds.

Confirmation consumes the token before mutation. Expired/rejected confirmations
are terminal too. Duplicate confirms replay the last result without rerunning
removal. A session reset/CloseUser clears token and result state; character ID
and generation checks additionally reject stale/cross-session tokens.

The source slot mutation latch is independent of byte equality. Audited paths:

| Mutation | Invalidation |
| --- | --- |
| Consume, destroy, sell, bank deposit, trade, quest removal | `QuitarUserInvItem`, plus slot publication |
| Pickup/stack insertion | `TryInsertGroundObjectIntoInventorySlot`, `MeterItemEnInventario` |
| Swap, move, merge, repair, apply elemental tags | both slots invalidated at `HandleMoveItem` entry |
| Equip/unequip | `EquiparInvItem`, `Desequipar`, slot publication |
| Direct smelting, theft, crafting, bank withdrawal, `QuitarObjetos` writes | existing `UpdateUserInv` / `ChangeUserInv` publication before returning to dispatch |
| Full clearing | `LimpiarInventario`, `VaciarInventario`, full inventory updates |
| Death | `UserDie` invalidates immediately, even if revival restores identical state |
| Disconnect/logout/character replacement | `CloseSocketSL`, `CloseUser`, `ResetUserSlot`, `IncreaseVersionId` clear state |

`ChangeUserInv` invalidates even an identical replacement, and movement invalidates
before a move can fail or restore equal bytes. A raw stack comparison at confirm
provides a second check on object, amount, equipped state and tags. This does not
introduce persisted versions, UUIDs, global locks or an inventory rewrite.

## Restrictions and removal

Requires a logged-in connected living character, not saving/logging out,
trading, auctioning, crafting, special fishing, automated work or a pending
bar action. Only ordinary inventory slots are accepted. Gold balances, bank
slots and cosmetic collections are not addressed by this protocol.

The actual ship slot supporting navigation/swimming and actual currently ridden
mount slot are protected. Spare boats/mounts remain allowed. Capture-flag objects
are protected. Scenarios exposing `eInventoryInterface` (naval cargo) are refused
because their pickup/drop bookkeeping has no safe permanent-removal hook; calling
drop here would create objects or bypass the scenario lifecycle.

For the final equipped item, existing `Desequipar` must explicitly report
completion and leave Equipped=0 before removal proceeds. It already owns
attribute, appearance, aura and equipment-reference updates. Errors/refusals must
not become successful removal. Partial equipped stacks remain equipped.
`QuitarUserInvItem` performs quantity removal and marks `ModificoInventario`.
The new endpoint checks completion and the exact post-removal amount. Inactive
ship/mount cached slot references are cleared when their final stack is removed.
An old targetable-item cursor is not treated as an inventory reservation: it can
survive client-side cancellation. A pending use of this exact source slot is
cleared before removal so a delayed WorkLeftClick cannot consume the remaining
stack through the old request. Targeting another slot is left alone.

Successful removal publishes `UpdateUserInv`, then a stage-1 result, and logs
character ID/name, object ID, stored elemental tags, quantity and request ID using
the existing server logging infrastructure. Networking/logging retain their
existing error behavior; duplicate confirm cannot repeat the mutation.

No house ownership or house_key query is issued when destroying an inventory
key. The normal character save writes the resulting inventory through
`SaveCharacterInventoryDB` / `QUERY_UPSERT_INVENTORY`; no new persistence path
exists. A dirty-flag test is not a substitute for a live save/relog test.

## Validation

The existing suite runner includes suite 42, `Unit_InventoryDestruction`.
For isolated native checks without account/database/world connections:

```text
VB6.EXE /make /d DEBUGGING=1:PYMMO=1:UNIT_TEST=1:UNLOCK_CPU=1 /out build.log Server.VBP
server.exe --test-inventory-destruction
```

The focused suite uses the real Aurora reader/writer, production handlers,
ordinary inventory removal and real move/change paths. It captures outgoing
packets only under UNIT_TEST=1. The fixture restores the chosen idle user's
state and object definitions before returning to the full suite.

Native Windows validation on 2026-10-20:

| PYMMO | DEBUGGING | UNIT_TEST | UNLOCK_CPU | Build | Focused tests | Full suite |
| ---: | ---: | ---: | ---: | --- | --- | --- |
| 1 | 0 | 0 | 1 | Passed | Not applicable | Not applicable |
| 1 | 1 | 1 | 1 | Passed | 41/41 | 479/479 |
| 0 | 0 | 0 | 1 | Passed | Not applicable | Not applicable |
| 0 | 1 | 1 | 1 | Passed | 41/41 | 479/479 |

The full suites ran from disposable Windows temporary directories with copied
resources, fresh copies of `Empty_db.db`, loopback networking and a nonproduction
port. An initial PYMMO=0 startup exceeded a 120-second map-load window; rerunning
with a 300-second startup window completed successfully. No production database
was used. Production builds were compiled, not deployed or exercised live.

Manual coordinated HOO/server checks still required: modal confirmation and
cancel/input focus, each equipment family (including torch and stat accessories),
NPC/player trade, active navigation/mounting versus spares, live event objectives,
packet loss/late replies, session switching, and normal save/relog. Verify that
the exact quantity disappears, no reward/ground object appears, house ownership
is unaffected, and old clients cannot invoke the gated endpoint.
