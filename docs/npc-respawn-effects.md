# NPC respawn presentation

The default-enabled **Settings → Graphics → NPC Respawn Effects** preference controls
one-second, client-only materialisation bursts. It is independent of graphics presets,
blood, weather and spell effects. Turning it off removes only these emitters/lights;
turning it on never replays missed events. English and Spanish labels are included.

## Authority and ordering

The server explicitly marks death-driven `MuereNpc` replacement work, both immediate
and queued, plus the explicit GM kill-with-respawn command. Administrative
repositioning remains excluded. `ReSpawnNpc` defaults to no visual notification: taming replacement,
AI recovery and administrative repositioning callers do not inherit death semantics.
The queue preserves the marker through unsuccessful attempts and schedule waits;
deferred initial population is explicitly excluded. Only a nonzero successful
`CrearNPC` result, valid active NPC, character index and matching map tile can notify.
Pets/summons are excluded. Generic creation, visibility refresh, reveal, crossing,
login and resynchronisation never trigger the presentation.

Notification 220 (`PYMMO=1`) is required protocol, without a capability flag.
All integers are little endian. The existing reliable frame contains:

| Field | Type |
|---|---|
| Packet ID = 220 | Int16 |
| Server respawn sequence, strictly positive | Int32 |
| Recipient's logical map at emission | Int16 |
| Spawn logical map | Int16 |
| NPC slot | Int16 |
| NPC lifetime version | Int16 |
| Character index | Int16 |
| NPC definition number | Int16 |
| Spawn x, y | UInt8 each |

Total: 20 bytes including ID. Invalid values, truncation and trailing bytes fail
normal packet validation. The sequence never wraps: after Int32 exhaustion,
cosmetic notifications cease until a server restart. It is not a gameplay identity.

Same-map recipients must be in the existing loaded-area bounds. A normal
recipient-specific CharacterCreate refresh is queued before the notification,
preserving appearance and dead-viewer rules and providing an ordering barrier.
Adjacent recipients are selected by the existing bounded two-hop visibility
collector, including its 128-NPC limit, resource-instance restrictions, invisible
NPC exclusion and unsafe-map dead-viewer rules. The current reset/upsert must
have been accepted before notification. No map-wide effect broadcast, retry queue
or retained effect subscription is introduced. A failed appearance enqueue means
no notification; later visibility retry contains appearance only.

## Client lifecycle

The decoder produces a typed event, GameEventHandler forwards it, and client
session wiring calls `NpcRespawnPresentation::receive`. It consumes the sequence
even when disabled or rejected. The recipient map must match current active map.
The current active-map NPC must match character index, definition and tile and be
visible/non-pet; adjacent NPCs additionally match the server NPC lifetime version.
The accepted one-shot immediately becomes an immutable map-local anchor. It never
looks up a character index again: removal, movement or slot reuse cannot attach
particles to another actor. TCP ordering means the appearance precedes its event
and a subsequent replacement follows it; there is no delayed character lookup.

Anchors retain both logical and resource map identity. Every frame projects them
through current prepared scene placements. Seamless spatial rebases preserve only
this controller's emitter IDs alongside existing adjacent environmental emitters;
reprojection translates existing particles without restarting their age. Ordinary
map replacement/teleport clears them. Session reset also clears the sequence.
If the source is no longer in the stitched scene, the effect ends immediately.
Missing textures drop the burst, without delayed playback or recurring logging.

## Visual and performance implementation

The existing packed atlas 4 / GRH 20216 white mote (544,672,32,32) was visually
inspected. No images, packs or deployment assets were added. The already-streamed
shared mote uses nonblocking residency checks and the existing texture-demand path.

Four components emit 192 particles over a complete burst: three soft violet feet
motes, 37 outward falling cyan-white sparks, two eight-sample curling streams
(128 motes), and 24 fading upward tail motes. Fixed-time helix samples make the
choreography independent of frame rate. Particles stay translucent and render
behind actors. A single smooth violet light per instance fades within 650 ms;
the colored particles remain available without HDR/bloom or dynamic lighting.
Reduce Screen Flashes halves opacity/light strength; enabling it cancels an
already brighter burst rather than flashing/restarting it.

At most four instances and sixteen emitters exist. Emission leaves 512 of the
shared 4096 particle slots for other effects; saturation skips cosmetic work and
never evicts unrelated particles. The shared capacity is unchanged. All emitters
are destroyed after one second or on disable/reset. No NPC behavior, collision,
movement, targeting, spawn timing or server probabilities are changed.

## Preview and manual checks

F1 → Particles → NPC respawn offers a real-effect preview at the local player's
tile and four simultaneous bursts, plus active/particle/accepted/skipped counters.
It creates no NPC and sends no gameplay packet. The player setting still applies.

With matching server/client builds:

1. Observe an actual death/respawn, both immediate and queued; verify NPC interaction
   is possible immediately and particles remain at its spawn tile while it moves.
2. Repeat at a stitched seam and cross during the one-second burst. Check both
   directions, two-hop viewers and no duplicate burst on entering visibility.
3. Check login, ordinary creation, summons, hidden NPCs, unsafe dead viewers,
   failed placements and visibility refreshes produce no inappropriate effect.
4. Preview in classic/modern, light/dark maps, varied zoom and HDR/bloom on/off.
   Check silhouette readability, Reduce Screen Flashes and simultaneous bursts.
5. Disable during a burst; verify this light/particles disappear while unrelated
   effects continue. Reenable, reconnect and teleport; check no replay.

## Deployment

Coordinate required server and HOO client versions. Publish the client that can
decode packet 220 before enabling the updated server, and ensure active clients
are updated before server rollout. Older clients cannot safely consume the new
notification; do not deploy the server alone. No asset-pack update, VB6 client,
ProtocolGenerator or HOO_CAP changes are required. No deployment is performed
by this change.

## Validation in this workspace

Client base: `6aa00b9daa1805af2bbd72bdb13a802d14208b24`.
Server base: `add6a089335de6c9575c8b39e0d1ebc4dc6b31d7`.

- Native Windows/MSVC Release client and unit-test binaries built successfully.
- 45 focused client tests passed (respawn, settings, dispatch, taming, adjacent NPCs).
- Full Windows unit run: 5,035 passed, 115 skipped, two failed. The failures are
  the existing topology-count and localization-identical-label assertions; their
  source/data were not changed by this feature.
- Windows CTest: 6/8 passed, including three GL smoke checks. The unit aggregate
  and topology check fail for the assertions above.
- Both localization JSON catalogs passed strict validation; git diff checks passed.
- Windows graphics client launched and remained running. This is startup
  verification, not visual or live-respawn signoff.
- Server: thirteen native regression cases were added, including queue provenance,
  visibility, active/adjacent ordering and failed-send suppression. They were not
  run: the documented VB6 compiler is absent from this Windows installation.
  Run a `UNIT_TEST=1, PYMMO=1` server with `--test-npc-respawn` in a disposable
  environment. Production notification call-site guards were audited separately.
- macOS compilation, multiplayer integration, the manual visual matrix and
  measured simultaneous-respawn frame-time validation remain outstanding.
