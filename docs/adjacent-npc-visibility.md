# Adjacent NPC visibility — Phase 1

Status: draft implementation; requires matching HOO/server builds and Windows runtime validation.

## Scope

Render server-owned NPCs on directly connected overworld maps. NPCs remain separate from the active-map character table and cannot be targeted, collided with, tamed, traded with or attacked across a seam. No ProtocolGenerator work and no new capability negotiation. Existing unrelated capability infrastructure is unchanged.

The server samples intersecting direct-neighbor tile rectangles every 100 ms and sends changes only. Interest is +/-24 tiles horizontally and +/-18 vertically in the viewer map's coordinate space, capped at 128 NPCs per viewer. Selection is deterministic by manifest neighbor order, then source tile order. This is a bounded Phase 1 region, not a guarantee of full viewport coverage at every resolution/zoom. No client-supplied map subscriptions are accepted.

The current server topology is a strict subset of HOO's manifest (20 directed edges are absent; shared edges agree). Phase 1 uses that shared subset and does not expand pursuit routes. Aligning the extra map coverage belongs to Phase 3.

Topology loading is independent of the pursuit switch. Only ordinary maps whose logical ID equals their resource ID are eligible. Missing topology means an empty adjacent scene; normal active-map play continues. Observed NPCs become eligible for the existing AI update, even when their map has no resident players; combat, movement legality and pursuit policy remain authoritative and unchanged.

## Wire contract

IDs below are for PYMMO=1. VB6's existing conditional PYMMO=0 ordinal shift remains; it is not a compatibility target. All multibyte values are little endian. The existing transport frame surrounds these payloads.

| ID | Name | Fields after the Int16 packet ID | Payload including ID |
|---|---|---|---|
| 214 | AdjacentNpcReset | Int32 scene generation, Int16 viewer logical map | 8 bytes |
| 215 | AdjacentNpcUpsert | Int32 generation; Int16 NPC index, version, source map; UInt8 x,y; Int16 body,head; UInt8 heading; Int16 weapon,shield,helmet,cart; Float32 speed; Boolean idle | 32 bytes |
| 216 | AdjacentNpcRemove | Int32 generation; Int16 NPC index,version | 10 bytes |

Upsert is a complete visual snapshot, serving creation and changed position/appearance/speed/idle state. There are no names, health bars or transient combat effects in this phase. Coordinates remain 1..100 in the source map; signed render offsets come from HOO's existing prepared map placements.

Scene generation is assigned by the server when a viewer session/map subscription is created. WriteChangeMap invalidates the subscription even for a reload of the same logical map. HOO clears adjacent state during map replacement and session reset and accepts updates only for its current generation. The server tracks user VersionId to detect slot reuse. NPC identity uses NpcIndex plus the existing VersionId; ordered transport and subscription resets bound its lifetime. It is not a globally unique persistent identity. Generation/version wrap under extreme churn needs additional Phase 2 coverage.

Full snapshots interpolate a single-tile move using existing character speed timing; multi-tile gaps, source-map changes and new NPC identities snap. A one-tile explicit translation is currently indistinguishable from walking. Polling can omit lifetimes shorter than 100 ms and produce up to one sampling interval of handoff delay; seamless movement across maps is Phase 2 work.

## HOO integration

Decoded events flow through GameEventHandler into WorldState's separate AdjacentNpcScene. Its records never enter active-map occupancy or interaction state. The existing character scene list adds render-only entries with source-map pointers and signed placement offsets, using prepared terrain ownership to suppress overlapping roots. Painter rows use composite coordinates. Roof visibility uses the source map and rendering uses a camera shifted by the map placement. The frame controller submits character texture demand, including all headings, through the existing streamer. Render-only flags suppress authoritative targeting effects and transient character-index effects.

## Validation

- Ten focused C++ tests pass in a standalone Linux GoogleTest build using production scene, packet decoder, WorldState and WorldMap sources.
- Changed C++ implementation translation units pass GCC C++20 syntax checks, including world rendering, event handling and the frame controller.
- VB6 topology/pursuit unit coverage is updated and five interest-boundary checks are added; these require Windows/VB6 execution.
- Neither full Windows client/CTest nor VB6 compilation nor live GL seam validation ran in the editing environment.

## Required runtime acceptance before merging

Use matching draft builds at a verified direct seam, for example map 34/35 (offset 0,-80 in the checked-in topology). Confirm the chosen positions are safe/legal in the actual resource set; do not alter gameplay restrictions for testing.

1. Observe an existing adjacent NPC without first entering its map: texture loads, correct position, facing and movement.
2. Spawn/remove an NPC and change its appearance/speed/idle state; check that unchanged NPCs are not continually resent.
3. Disable pursuit: adjacent rendering continues; pursuit behavior remains disabled.
4. Inspect a seam tree/roof and overlap roots; verify painter order and no duplicate visual.
5. Reload the same map, teleport, logout/relogin and switch accounts: no retained adjacent records.
6. Cross the boundary in both directions, including a pursuing NPC. Record polling gaps for Phase 2; no persistent duplicate/ghost is acceptable.
7. Check a logical map alias/instance: no NPCs from unrelated world instances appear.

Later phases cover seamless handoffs, richer movement semantics, corner/multi-hop coverage, viewport/zoom guarantees, complete visual parity and measured server/network load. Do not deploy this Phase 1 draft as a completed production feature.
