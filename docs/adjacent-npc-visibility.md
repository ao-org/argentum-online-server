# Adjacent NPC visibility — Phase 1

Status: draft implementation; requires matching HOO/server builds and Windows runtime validation.

## Scope

Render server-owned NPCs on overworld maps reachable through at most two topology edges. NPCs remain separate from the active-map character table and cannot be targeted, collided with, tamed, traded with or attacked across a seam. No ProtocolGenerator work and no new capability negotiation. Existing unrelated capability infrastructure is unchanged.

The server samples intersecting one- and two-hop tile rectangles every 100 ms and sends changes only. Interest is +/-24 tiles horizontally and +/-18 vertically in the viewer map's coordinate space, capped at 128 NPCs per viewer. Selection follows HOO's placement order: minimum topology depth, then map ID, then source tile order (y, x). This is a bounded Phase 1 region, not a guarantee of full viewport coverage at every resolution/zoom. No client-supplied map subscriptions are accepted.

Visibility composes signed edge offsets using 32-bit arithmetic. Identical placements
reached by multiple paths are scanned once. Returning to the active map is excluded,
and discovery stops after exactly two edge levels, even in cyclic topology. Empty
or out-of-interest intermediate maps do not prevent second-hop discovery. Only each
placement's intersection with the player-centered interest rectangle is scanned.
Conflicting offsets for one source map suppress that map for this viewer: the wire
identifies a source map, not a path, so selecting an arbitrary transform is unsafe.

For example, from map 80, both 80 -> 78 -> 34 and 80 -> 35 -> 34 place Nix at
(+74,+80). A guard there is sent once after the map-80 reset, and is sent again in
the new map-78 scene on return. No diagonal manifest edges are added; pursuit and
movement routing still use only the original edges. Perfect crossing continuity
remains Phase 2 work, not stale-scene retention.

The server, HOO and assets topology files must contain the same reviewed data.
The synchronized manifest has 437 maps and 1,494 directed edges, retaining the
additions from server #1873 and assets #676. It also restores the server's older
missing catacomb (40-45), southern-sea (529 and its neighboring seams), and
342/343 connections: seven maps and 22 directed edges. Each matches the existing
offline map audit's VerifiedSpatial classification. No existing edge is removed
or changed. These routes become eligible for existing NPC visibility and pursuit
under the unchanged toggle, safe-map, instance and movement checks.

Updating source files does not deploy the server's loose file or the client's
pack. Run HOO's `tools/check_topology_sync.py` against the selected three
checkouts before release and verify installed data separately.

Topology loading is independent of the pursuit switch. Only ordinary maps whose logical ID equals their resource ID are eligible. Missing topology means an empty adjacent scene; normal active-map play continues. Observed NPCs become eligible for the existing AI update, even when their map has no resident players; combat, movement legality and pursuit policy remain authoritative and unchanged.

## Visibility and send recovery

Adjacent collection excludes NPCs whose `flags.invisible` is nonzero or whose body
is absent. The visual-only packet has no invisibility representation, so hidden
NPCs are omitted even for staff or their owner; the normal NPC path has no
viewer-specific invisibility exception to reproduce. Player visibility exceptions
do not apply. The existing `MakeNPCChar` dead-viewer rule remains: on an unsafe
viewer map, only visible revivers are included. Instance restrictions are unchanged.
Becoming hidden removes an existing adjacent record; becoming visible sends a fresh
upsert without requiring movement or an appearance change.

The cache records only packets accepted by the reliable transport queue, not
confirmed client receipt. Aurora's `Send` COM return alone does not prove enqueue:
the adjacent-only opt-in checks its synchronous forwarding callback for the expected
connection and full payload length, successful send return, and unchanged user
identity. No transport DLL changes or client acknowledgements are required. Other
send callers retain their existing behavior.

A failed reset leaves the subscription uninitialized and sends no upserts. After a
successful reset, each accepted removal/upsert updates the known cache individually.
The first failed difference stops that refresh; the next 100 ms sample retries the
remaining differences. Already accepted changes are not resent. An old NPC lifetime
must be removed successfully before its slot's replacement is upserted. There is no
extra retry queue or unbounded snapshot loop. Shared writer bytes are cleared on
both accepted and rejected sends, and send-failure diagnostics are rate-limited.
User generation, connection, map replacement and disconnect invalidate obsolete
state. Current interest still drives the existing observed-NPC AI during retries.

## Wire contract

IDs below are for PYMMO=1. VB6's existing conditional PYMMO=0 ordinal shift remains; it is not a compatibility target. All multibyte values are little endian. The existing transport frame surrounds these payloads.

| ID | Name | Fields after the Int16 packet ID | Payload including ID |
|---|---|---|---|
| 214 | AdjacentNpcReset | Int32 scene generation, Int16 viewer logical map | 8 bytes |
| 215 | AdjacentNpcUpsert | Int32 generation; Int16 NPC index, version, source map; UInt8 x,y; Int16 body,head; UInt8 heading; Int16 weapon,shield,helmet,cart; Float32 speed; Boolean idle; Int16 NPC definition number; String8 authorized display text | 34 + varuint length-prefix bytes + UTF-8 text bytes |
| 216 | AdjacentNpcRemove | Int32 generation; Int16 NPC index,version | 10 bytes |

Upsert is a complete visual snapshot, serving creation and changed
position/appearance/speed/idle/definition/display-text state. Coordinates remain
1..100 in the source map; signed render offsets come from HOO's existing prepared
map placements, including diagonal maps. Health bars and transient combat effects
remain outside adjacent replication.

The original first 32 bytes are unchanged. The definition number is
`NpcList(index).Numero` at offsets 32-33, not the live NPC slot, character index or
lifetime version. String8 starts at offset 34: Aurora writes the UTF-8 byte length
as an unsigned 7-bit continuation varint, then the exact UTF-8 bytes with no NUL
terminator. Empty text is one zero byte (35-byte total packet). The existing
UInt16 transport payload limit is 65,535 bytes, so the maximum adjacent display
text is 65,498 UTF-8 bytes with a three-byte prefix. Oversized snapshots fail before
enqueue, clear the shared writer, and do not advance the accepted-state cache;
correcting the runtime text allows the next sample to retry. Errors use the
existing rate-limited logging. Text is never silently truncated.

`GetNpcDisplayName` is shared with normal `MakeNPCChar`: `showName > 0` sends
the runtime `.name & .SubName` without changing delimiters or substituting static
catalogue text; otherwise it sends empty text. Hidden text changes do not cause
updates until labels are enabled. The client must use its normal label parser and
visual metadata/body-selection rules with the source map's terrain. Normal NPC
creation now maps cart to `.Char.CartAnim` and helmet to `.Char.CascoAnim`, with
backpack, FX and FXLoops zero. The previous argument order incorrectly put helmet
and cart into FX and FXLoops. Area introduction and refresh both delegate to this
same production creation path; player creation is unchanged.

This extends packet 215 in place and is **not mixed-version compatible**: older
HOO expects 32 bytes and cannot consume the appended fields, while new HOO cannot
decode old upserts. Update the server and matching HOO together during a controlled
release that prevents mismatched clients from connecting. Neither arbitrary
client-first nor server-first rollout with mixed sessions is safe. Packets 214 and
216 and unrelated inventory/bow IDs are unchanged. No capability or generator is
involved. This work does not authorize deployment.

Scene generation is assigned by the server when a viewer session/map subscription is created. WriteChangeMap invalidates the subscription even for a reload of the same logical map. HOO clears adjacent state during map replacement and session reset and accepts updates only for its current generation. The server tracks user VersionId to detect slot reuse. NPC identity uses NpcIndex plus the existing VersionId; ordered transport and subscription resets bound its lifetime. It is not a globally unique persistent identity. Generation/version wrap under extreme churn needs additional Phase 2 coverage.

NPC versions increment on release through 0..32761 and then wrap to zero. Hidden
or out-of-interest NPCs can legitimately reappear with the same lifetime; a client
must not permanently tombstone that version merely because a removal arrived.
The ordered reliable stream, scene generation and accepted old-lifetime removal
before replacement remain the lifecycle guarantees; the packet carries no
per-update sequence number for arbitrary reordering within one scene.

Full snapshots interpolate a single-tile move using existing character speed timing; multi-tile gaps, source-map changes and new NPC identities snap. A one-tile explicit translation is currently indistinguishable from walking. Polling can omit lifetimes shorter than 100 ms and produce up to one sampling interval of handoff delay; seamless movement across maps is Phase 2 work.

## HOO integration

Decoded events flow through GameEventHandler into WorldState's separate AdjacentNpcScene. Its records never enter active-map occupancy or interaction state. The existing character scene list adds render-only entries with source-map pointers and signed placement offsets. Terrain ownership does not filter NPCs; records are deduplicated by NPC identity and lifecycle, and duplicate views render each NPC only once. Painter rows use composite coordinates. Roof visibility uses the source map and rendering uses a camera shifted by the map placement. The frame controller submits character texture demand, including all headings, through the existing streamer. Render-only flags suppress authoritative targeting effects and transient character-index effects.

## Validation

### Appearance and labels (2026-10-25)

- Native Windows VB6 production and test builds passed for both `PYMMO=0` and
  `PYMMO=1`. Each build used `VB6.exe /make /d
  DEBUGGING=<0|1>:PYMMO=<0|1>:UNIT_TEST=<0|1>:UNLOCK_CPU=1 /out <log> Server.VBP`,
  with DEBUGGING matching UNIT_TEST, and a fresh executable was verified.
- The full 43-suite run passed **670/670** under each PYMMO configuration, using
  an isolated source copy, loopback server and disposable SQLite fixture.
- Ten new production-path groups compare normal creation and adjacent equipment
  with distinct nonzero IDs and no FX; definition/name/description refreshes;
  showName hiding/revealing and empty labels; accented UTF-8, two-/three-byte
  String8 prefixes and the full 65,535-byte packet; oversize recovery; and failed
  text-update retry. Existing visibility, two-hop, removal and lifecycle tests
  remain green. The test reader now retains its borrowed input after returning
  from the helper, fixing an invalid-buffer lifetime exposed by the large packet.
- The production-captured Spanish upsert is logged as hex in `test_progress.txt`
  for independent client fixture verification. VB6 source encoding/CRLF and
  `git diff --check` were checked. User-local project settings were preserved.
- No matching HOO/server live rendering session was performed in this server
  pass. Ullathorpe boundary crossings, moving/stopping/idle visuals and water-NPC
  acceptance remain pending; client rendering validation belongs to the linked
  HOO change.

### Earlier Phase 1 checks

- Ten focused C++ tests pass in a standalone Linux GoogleTest build using production scene, packet decoder, WorldState and WorldMap sources.
- Changed C++ implementation translation units pass GCC C++20 syntax checks, including world rendering, event handling and the frame controller.
- The two-hop pass executed the full native VB6 unit suite in an
  isolated loopback test environment with `DEBUGGING=1:UNIT_TEST=1:UNLOCK_CPU=1`:
  **600/600 passed for PYMMO=1 and 600/600 for PYMMO=0**. This includes 11 new
  two-hop groups, the previous 28 production collection/refresh groups and nine enqueue-probe groups.
  Tests capture real serialized packets and inject failed acceptance; they verify
  packet ordering, cache contents, stale lifetime handling and shared-writer cleanup.
- Two-hop groups cover the checked-in Nix corner and fresh snapshots across
  78 -> 80 -> 78; all four signed corners; exact interest boundaries; one scan for
  convergent paths; empty intermediates; cycles, active-map exclusion and the
  two-edge cap; conflicting offsets; +/-60000 composition; instances and retries.
- HOO's production loader/prepared-scene/render-entry tests passed natively on
  Windows Release: 55/55 focused adjacent, overworld and map-placement tests,
  including three new corner regressions. HOO required no runtime changes.
- New coverage includes hidden/visible transitions, staff/owner and dead-viewer
  rules, reset/upsert/remove failures, partial success, unchanged-state silence,
  NPC-slot reuse, and map/user-generation/connection changes during recovery.
- Enqueue-probe tests cover complete/wrong/partial/duplicate callbacks, reentrancy,
  close invalidation and late callbacks. They are deterministic unit tests, not a
  real-socket queue-saturation test. The callback semantics were checked against
  the installed Aurora implementation; no Aurora files or dependencies changed.
- Native production builds passed for both `PYMMO=0` and `PYMMO=1` with
  `DEBUGGING=0:UNIT_TEST=0:UNLOCK_CPU=1`. All four builds used `VB6.exe /make /d
  <definitions> /out <log> Server.VBP`; each new executable was verified. Existing
  local project reference settings were left untouched and uncommitted.
- Matching live HOO/server seam checks were not performed in this pass. See the
  coordinated PRs for current build evidence and outstanding visual acceptance.

## Required runtime acceptance before merging

Use matching draft builds at a verified direct seam, for example map 34/35 (offset 0,-80 in the checked-in topology). Confirm the chosen positions are safe/legal in the actual resource set; do not alter gameplay restrictions for testing.

1. Observe an existing adjacent NPC without first entering its map: texture loads, correct position, facing and movement.
2. Spawn/remove an NPC and change its appearance/speed/idle state; check that unchanged NPCs are not continually resent.
3. Disable pursuit: adjacent rendering continues; pursuit behavior remains disabled.
4. Inspect a seam tree/roof and overlap roots; verify painter order and no duplicate visual.
5. Reload the same map, teleport, logout/relogin and switch accounts: no retained adjacent records.
6. Cross the boundary in both directions, including a pursuing NPC. Record polling gaps for Phase 2; no persistent duplicate/ghost is acceptable.
7. Check a logical map alias/instance: no NPCs from unrelated world instances appear.
8. Hide and reveal an adjacent NPC without moving it: removal followed by a fresh
   visible record, including staff/owner viewers and dead viewers near revivers.
9. Observe the Nix guard from map 78, enter map 80 and return. Each new scene must
   receive that guard again, once, at the appropriate composed placement.
10. With matching appearance/label builds, observe the same guards south of
    Ullathorpe on the active map and across the seam in both directions. Compare
    body, head, weapon, shield, helmet, cart, name and description; there must be
    no equipment-triggered FX. Wait through idle changes and observe a moving NPC
    stop without another update. Repeat crossings and reconnect to check for
    missing/duplicate sprites or labels. Check a water-specific NPC on its source
    map and from a diagonal prepared view where suitable test data exists.

These appearance checks require a real matching client/server development session.
Serialized packet regressions alone do not establish visual acceptance.

Later phases cover seamless handoffs, richer movement semantics, coverage beyond two edges, viewport/zoom guarantees and measured server/network load. Do not deploy this Phase 1 draft as a completed production feature.

## Topology deployment checklist

There are three separately maintained source copies:

- HOO: `resources/config/legacy_overworld_topology.json`.
- Assets: `config/legacy_overworld_topology.json`, packaged into `core.hoopack`.
- Server: `Recursos/config/legacy_overworld_topology.json`, loaded as a loose file
  from `App.Path` by `LoadAdjacentTopology` at startup.

Require byte-for-byte equality of the three source copies, checking both map IDs
and signed offsets. Review server additions explicitly because the graph also
supplies pursuit, then synchronize the complete approved file. Do not retain a
stale server subset. A client/assets PR or rebuilt Steam pack does not deploy
server data.

Deploy this server JSON beside the actual server executable under `Recursos/config`,
then restart through the normal release process and check `Adjacent topology loaded`
in the server log. Compare the installed file with the reviewed revision; counts
alone are not freshness evidence. Do not replace private server configuration.
This PR does not authorize a deployment or restart.

Validate 144 -> 145 and 145 -> 144 with a real visible NPC near the seam: observe it
before crossing, cross, and verify it remains available in the new adjacent scene.
Also test pursuit enabled/disabled and existing safe-map restrictions. Adjacent NPCs
remain visual-only across maps, and the existing polling/handoff delay still applies.

The checked-in-manifest VB6 regression now requires both Veril edges and actual
visibility placements in both directions. Run the existing NPC cross-map suite
against the deployment's manifest; do not substitute a synthetic manifest for this
release check. HOO's `docs/steam-resource-build.md` documents client pack verification.
