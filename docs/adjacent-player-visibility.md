# Adjacent Player Visual Replication

The required HOO protocol adds a separate, render-only player stream. There is
no capability negotiation, ProtocolGenerator update, Aurora DLL change, or
change to NPC packets 214-216. Coordinated server/HOO deployment is required.
Production hooks run only under `PYMMO=1`; `PYMMO=0` retains its normal stream.

## Wire Contract

All integers/floats use the existing little-endian Network.Writer. Packet IDs
include the existing conditional enum offset: 217/218/219 under PYMMO=1 and
218/219/220 under PYMMO=0 (the latter are reserved, not emitted in production).

| Packet | Payload after Int16 packet ID |
| --- | --- |
| Reset 217 | Int32 positive scene epoch, Int16 viewer map |
| Remove 219 | Int32 epoch, Int16 user slot, UInt64 session (low then high Int32 words) |
| Upsert 218 | Fields below, in order |

Upsert has 47 fixed bytes including its ID, followed by eight String8 fields:

1. Int32 epoch, Int16 user slot, UInt64 opaque session ID (low/high words).
2. Int16 source map, Byte x, Byte y.
3. Int16 body, Int16 head, Byte heading.
4. Int16 weapon, shield, helmet, cart, backpack.
5. Float32 speed, Byte representation, UInt16 visual flags.
6. Byte status, privileges, user type, team, flag.
7. Authorized display name (4096 UTF-8 bytes maximum).
8. Seven persistent aura strings (1024 UTF-8 bytes maximum each), ordered:
   Weapon, Body, Shield, Head, Other, DM, RM.

String8 uses the existing 7-bit varint UTF-8 byte-length prefix, not an 8-bit
length. An oversized field fails serialization, clears the shared writer, and
does not replace accepted state. Corrected/unchanged failed updates are retried.

Representation values reproduce HOO `character_visibility.cpp`:

| Value | Color / Alpha | Dim Identity | Auras |
| --- | --- | --- | --- |
| 0 normal | white / 255 | no | yes |
| 1 staff | white / 100 | yes | only clan-visible or navigating |
| 2 clan | RGB 180,160,160 / 160 | yes | yes |
| 3 boat | RGB 125,125,125 / 125 | no | yes |

Flag bits: 0 idle, 1 navigating, 2-4 reserved zero,
5 draw auras, 6 snap, 7 same-clan label tint. Player idle is false, as in the
normal `MakeUserChar` snapshot. HOO derives mount/swim/rubber visuals from the
authorized body using the normal CharacterCreate metadata path. Actual server
gameplay transport flags can differ under mimic/skins and are not disclosed.
Snap is derived at send time, not cached as persistent visual state. A warp
revision forces snap even for a same-map, one-tile `WarpUserChar`; a normal
one-Manhattan-tile move can interpolate. Source-map changes and larger moves
snap. No transient combat, chat, audio, projectile, or effect event is included.

## Collection And Admission

`modAdjacentPlayerVisibility` samples every 100 ms through the existing
`TryGetAdjacentVisibilityPlacements` API. It intersects each unambiguous,
resource-map-only, at-most-two-edge placement with the viewer's inclusive
24-by-18 tile radius. It scans only those MapData rectangles. Because one tile
can hold only one `UserIndex`, it also examines that map's existing
`ConnGroups.UserEntrys`, filtering by position. This retains overlapping
ghost/staff characters without a global per-viewer user scan.

Candidates must be connected, logged in, have a positive character index owned
by their user slot in `CharList`, and match their indexed source map/position.
Viewers must likewise own an active character and a valid map position before
receiving a reset. Active-map characters are excluded. Slots are sorted before
authorization and the 128-player cap; NPC capacity is entirely independent.
Pending handoffs reserve capacity within that bound.

Authorization happens before building any location/name/appearance snapshot:

- Match `ModAreas.NotifyUser` first. Admin-invisible targets require
  `Admin.CanUserSeeAdminInvisible`; an authorized admin target takes precedence
  over the dead-viewer restriction, exactly as the existing ElseIf does.
- Otherwise, a dead viewer on an unsafe **viewer map** needs matching nonzero
  actual guild membership at the required level, or the **target's**
  `eTalkToDead`. Viewer talk-to-dead or staff privileges alone do not grant this.
- For hidden admitted players, apply HOO rendering priority: staff rank,
  clan exception, boat exception, otherwise exclude. `EsGM` is not the staff
  rank test because it omits RoleMaster. Casper bodies 829, 830, and 87 do not
  get the hidden-clan exception.
- HOO receives clan level zero when MakeUserChar hides names or uses a mimic
  name. The adjacent visibility decision uses the same effective clan level,
  not an undisclosed real level. The same-clan label bit is suppressed under
  disguise, hidden names, and a concealed pirate label; no guild ID is sent.

## Existing Rule Differences

These differences are deliberately documented, not normalized across gameplay:

- Some `modSendData` event routes admit dead same-clan viewers without checking
  clan level, and some use viewer talk-to-dead. Snapshot admission follows
  `NotifyUser`, not those transient event policies.
- Nearby hidden actors can generate location-bearing effects/chat events via
  distance/timer exceptions. Those events do not authorize a persistent visible
  adjacent character. Their exceptions are not copied into this stream.
- `MakeUserChar` builds `displayName {authorized alias} <guild>` (or a guildless
  consultation tag), while `RefreshCharStatus` uses a simpler label. The new
  snapshot follows MakeUserChar, including its feature/tier-gated
  `GetCharacterAlias`, `showName`, and exact `NameMimetizado` with no fallback.
- `RefreshCharStatus` explicitly blanks a concealed pirate's name, unlike
  MakeUserChar. The adjacent snapshot keeps that privacy restriction.

## Identity And Queue Safety

The opaque process-lifetime counter uses two unsigned 32-bit words stored in
VB6 Longs. Neither account/database IDs nor wrapping `VersionId` alone are used.
`ResetUserSlot` retires both viewer state and source session identity; a
connection/version signature additionally detects replacement. Source lifetime
IDs remain stable across ordinary map/visual changes. At UInt64 exhaustion new
identities fail closed, never wrap. Scene epochs are positive Longs and fail
closed at exhaustion. They are independent from NPC scenes and identities.

Only accepted packets update the cache. A rejected reset cannot initialize a
scene or permit upserts. Removals precede replacement identities, each accepted
difference commits independently, and unchanged rejected packets retry. Writer
errors clear the shared buffer; error logging is rate limited to once/10 s.

Adjacent-to-active handoff reserves a bounded retry before removal. The normal
CharacterCreate is deferred if removal is rejected. If removal succeeds but
the create fails, its retry remains reserved. Retries revalidate viewer
lifetime/map, exact subject session/map/character index, live CharList ownership,
area overlap, and current normal admission. Unrelated normal-create failures do not
enter this retry queue. Admission, movement, and gameplay are not deferred.

Same-map retries use normal `NotifyUser` admission, not adjacent render
eligibility. Ordinary `invisible` or `Oculto` players still need a normal record
for collision and a later stationary reveal. `AdminInvisible` is solely an
admission gate, not an extra normal hidden-state flag. Temporary admission denial
retains the bounded entry without a timeout; obsolete subject work is canceled.

A hidden handoff uses three individually confirmed stages after the adjacent
remove: a zero-appearance normal CharacterCreate, SetInvisible with the current
normal hidden state, then the full MakeUserChar refresh. The blank retains the
actual character index, position, status and privileges, but zeroes body, head,
equipment, FX, FXLoops and particle, and empties its label and seven auras.
This requires HOO's strict blank-PC guard: a matching create starts invisible,
including when replacing a previously visible record. Interleaved ordinary
CharacterChange packets therefore cannot expose a pending blank. A full refresh
preserves the prior invisible flag. No packet layout or transport framing changes.

Accepted stages are not replayed. If visibility changes before the full refresh
is accepted, its new value must be accepted first. Visible handoffs keep the
single-create path. NotifyUser suppresses its unconfirmed duplicate visibility
send while this handoff owns that state. Existing post-create orb/torch snapshots
remain unchanged. Temporary denial retires any accepted partial normal record
before restarting the blank stage. Obsolete subjects similarly retain a
cleanup-only entry until normal CharacterRemove is accepted. Cleanup matches
the pending session, subject, map and old character index against the accepted
normal-owner ledger, so it cannot remove a newer owner of a reused index.
Viewer disconnect/map reset discards the old world without subject cleanup.

If a real MakeUserChar request is blocked behind obsolete cleanup, that entry
also retains one separate replacement request. The old accepted identity/index
continues owning cleanup; the replacement captures its own viewer/subject
lifetimes, map and character index. It does not replace the old removal target.
Repeated notifications retain one request; a newer explicit request supersedes
only that request, never cleanup. Both share the existing 128-entry reservation.

After removal is accepted, retry promotes only a still-current saved request,
rechecking connection/session, character ownership, map, loaded area and normal
admission. Appearance and hidden state are rebuilt from current authorized data.
Failed replacement creation retains its own remaining stages, including the
blank/hide/full path. No additional notification or movement is required.
Cleanup-only work never creates the current occupant of a reused slot. Obsolete
replacements are discarded, and viewer teardown discards both responsibilities.

For the reverse handoff, accepted normal creates are tracked per viewer, keyed
by the bounded normal character-index namespace, with the actual
user/session/map/character index. A failed legacy remove can leave multiple
indices for one subject; they must all survive in this ledger. Before an
adjacent upsert, confirmed normal CharacterRemoves retire every matching index,
committing accepted removals independently. This remains correct if WarpUserChar
allocates a different current index. Any accepted normal
create, including an NPC, replaces the prior tracked owner of that index;
rejected creates cannot clear it. PYMMO=1 WriteCharacterCreate therefore requests
the existing transport queue-confirmation probe. Normal packet layouts are
unchanged. Map clear and viewer reset discard both tracking namespaces.

## Validation

`Unit_AdjacentPlayers` is registered as suite 44. It captures the production
writer and SendData acceptance boundary with controlled per-packet failures;
P0 explicitly exercises the collector test entry point while checking that
its production tick remains disabled. The suite preserves fixture users,
maps, map membership, guild data, character ownership, and recent-login cache.
Case 1 writes actual native-writer bytes to `adjacent-player-upsert.hex` in the
disposable test runtime for independent HOO literal-fixture validation.
Case 46 captures the accepted handoff sequence as six separate payloads named
`adjacent-player-handoff-{remove,blank,hide,full,torch-clear,reveal}.hex`:
adjacent remove (16 bytes), blank create (70), hide (7), full create (107),
normal torch clear (17), and stationary reveal (7). These are existing individual
frames, never concatenated packet payloads. The fixture uses viewer map 34 at
50,11 and character 2 at 50,12, scene epoch 1 and subject session 2.

Case 76 additionally captures `adjacent-player-replacement-{old-blank,remove,
blank,hide,full,torch,reveal}.hex`. This exercises the real NotifyUser path with
rejected obsolete cleanup, then resumes without another notification: remove
normal index 2, create hidden replacement index 3, then stationary reveal.
Cases 77-90 cover same-session index changes, session reuse, repeated failures,
independently accepted stages, repeated/superseding notifications, obsolete
requests, reused old-index ownership, cleanup-only cancellation, viewer reset,
the 128-entry cap and temporary admission denial.

Native Windows validation: all four configurations compiled successfully;
both U1 runs passed **767/767 tests across 44 suites** (677 baseline checks plus
90 adjacent-player cases). P0 conditionally skips P1-only lifecycle/handoff hooks
and verifies the disabled production stream. U1 executables ran in isolated,
loopback-only fixtures with empty test databases. U0 executables were compiled,
not launched. Fresh output timestamps were checked; the primary executable and
production data were untouched.

Deferred-handoff coverage includes both initial remove/create failures crossed
with ordinary invisibility and Oculto, further blank/visibility/full failures,
stationary reveal, changed visibility before a failed full refresh, permitted
admin-only parity with NotifyUser, temporary admission denial/recovery, obsolete
subject cleanup, reused-index ownership, viewer reset, and the 128-entry bound.
For the earlier hide/reveal fix, restoring the old `Representation` gate in an
isolated negative control compiled but failed **9 tests (744/753 passed)**,
including case 46.
For replacement recovery, case 76 was first extended against reviewed server
`e625140cdd5474b98b187511e44867212c95273f`, before changing production code.
That native negative control compiled but failed only case 76: **752/753**,
with `rejected cleanup retains requested replacement without another notification`.
The final fixed source passed the four-build matrix above. Seven actual native
writer captures feed HOO's replacement-event/collision/generation regressions.

### Compiler Matrix

Set `VB6_EXE` to the installed compiler, `ISOLATED_BUILD` to a disposable copy of
the project, and `VALIDATION_LOGS` to the output directory. The four tested flag
sets are:

```bat
"%VB6_EXE%" /make /d DEBUGGING=1:PYMMO=1:UNIT_TEST=1:UNLOCK_CPU=1 /out "%VALIDATION_LOGS%\pymmo1-unit1-build.log" "%ISOLATED_BUILD%\Server.VBP"
"%VB6_EXE%" /make /d DEBUGGING=1:PYMMO=0:UNIT_TEST=1:UNLOCK_CPU=1 /out "%VALIDATION_LOGS%\pymmo0-unit1-build.log" "%ISOLATED_BUILD%\Server.VBP"
"%VB6_EXE%" /make /d DEBUGGING=0:PYMMO=0:UNIT_TEST=0:UNLOCK_CPU=1 /out "%VALIDATION_LOGS%\pymmo0-unit0-build.log" "%ISOLATED_BUILD%\Server.VBP"
"%VB6_EXE%" /make /d DEBUGGING=0:PYMMO=1:UNIT_TEST=0:UNLOCK_CPU=1 /out "%VALIDATION_LOGS%\pymmo1-unit0-build.log" "%ISOLATED_BUILD%\Server.VBP"
```

### Captured Upsert

The actual P1 upsert capture is 92 bytes from the production writer, not a
hand-assembled fixture:

```text
DA 00 01 00 00 00 02 00 45 33 22 11 88 77 66 55 23 00 32 5A 0A 00 14 00 03 1E 00 28 00 32 00 3C 00 46 00 00 00 C0 3F 00 60 00 01 01 02 03 04 08 4A 6F 73 C3 A9 20 7B 7D 06 77 65 61 70 6F 6E 04 62 6F 64 79 06 73 68 69 65 6C 64 04 68 65 61 64 05 6F 74 68 65 72 02 64 6D 02 72 6D
```

Remaining runtime boundary: legacy `WriteChangeMap` still uses its existing
unconfirmed enqueue path. This change does not add general legacy-map-stream
recovery. It confirms the new scene packets and both character handoff barriers.
Live two-client testing with matching HOO builds was not performed.

Manual integration checklist:

- Start matching server/HOO test builds and inspect startup/packet/error logs.
- Observe direct and diagonal borders, overlapping characters, cap stability,
  movement/equipment/name/aura changes, and all four authorized modes.
- Toggle invisibility, clan level, admin rank, death/safety, mimic, and labels
  without moving. Unauthorized subjects must reveal neither name nor location.
- Cross both directions, warp one tile, logout/reconnect, reuse a user slot,
  and change viewer maps. Verify there is never an active/adjacent duplicate.
- Inject reset/upsert/remove/normal-create failures and verify ordered recovery.
- After an accepted blank, change the subject session/index and notify the
  viewer. Reject obsolete cleanup repeatedly, then accept it without another
  notification. Verify only the requested replacement appears, no old collision
  remains, hidden recovery/reveal is correct, and completed stages stay silent.
- Fail a hidden handoff's blank/hide/full sends, interleave appearance changes,
  then reveal without moving. Verify no premature rendering, light, duplicate,
  or lost collision record; obsolete partial records must eventually be removed.
- Roll back only as a coordinated HOO/server deployment; P0 does not emit the
  stream, but is not a promise of compatibility with old HOO builds.
