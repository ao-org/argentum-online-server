# Trigger and map schema migration

The server, Heroes of Old (HOO), and `argentum-online-assets` will move from exclusive numeric tile triggers to independent tile flags, map-wide zone flags, spatial roof regions, and castle references. This specification defines the coordinated implementation and data conversion. It is a design-only deliverable; the gameplay, database, and binary map migrations require later implementation. The preparatory GM-only `/trigger` command is delivered in separate HOO and server PRs: its set payload becomes a nonnegative 32-bit value, and HOO gains query and set commands.

The three planning PRs share branch `codex/trigger-map-schema-migration` and title **Prepare trigger and map schema migration**. Command implementation uses separate `codex/trigger-command-long` branches and PRs. HOO and assets keep repository-specific checklists in `docs/trigger_migration.md` and `docs/trigger-migration.md`, respectively. Review bit allocation, header layout, and roof seams as proposed contracts before implementing them.

## Repository responsibilities

| Repository | Required work |
|---|---|
| `argentum-online-server` | Flag definitions and consumers, map-wide prison checks, new map reader, castle schema migration and ID lookup, runtime entrance bindings, 32-bit GM command, authoritative tile-change notification, regression tests. |
| `heroes-of-old` | Matching flags and binary reader, world state and movement consumers, roof flood fill and fading, GM command and help, tile-change handling, parser and rendering tests. Keep implementation in the owning modules; the application composition root only wires services. |
| `argentum-online-assets` | Convert every distributed map to the new binary schema, update trigger labels and map tooling, publish format fixtures and a conversion manifest, rebuild asset packs containing maps. The local checkout currently named `Recursos` has this repository as its origin. |

Ship compatible server, HOO, and assets revisions together. The old VB6 client is outside the implementation target; an old client cannot safely send the new set-trigger packet or read the new maps.

## Tile flags

Use PascalCase for every new trigger member. Historical IDs identify conversion inputs only; they are not bit positions or new numeric values. Use a VB6 `Long` on the server and a 32-bit integer in HOO and serialized data. Keep bit 31 clear so values have the same nonnegative representation in both implementations.

| Flag | Bit | Decimal | Hexadecimal |
|---|---:|---:|---:|
| `None` | None | 0 | `0x00000000` |
| `UnderRoof` | 0 | 1 | `0x00000001` |
| `AntiNpcRespawn` | 1 | 2 | `0x00000002` |
| `WayUnblocker` | 2 | 4 | `0x00000004` |
| `PvPArena` | 3 | 8 | `0x00000008` |
| `AutoResurrection` | 4 | 16 | `0x00000010` |
| `SwimSuitPath` | 5 | 32 | `0x00000020` |
| `NoFishing` | 6 | 64 | `0x00000040` |
| `GhostOnlyTranslator` | 7 | 128 | `0x00000080` |
| `CastleFoundationPosition` | 8 | 256 | `0x00000100` |

The initial known mask is `511` (`0x000001FF`). Bits 9 through 30 are reserved. For example, `UnderRoof | PvPArena` is `9`; historical trigger 9 does not mean those two flags. The map format version determines which interpretation applies.

Test a single flag with `(flags And flag) <> 0`. When testing a collection, distinguish any matching bit from all required bits. The existing `IsSet` helper means any matching bit. Setting one flag must preserve other bits; removing a flag uses `And Not`. A GM set command replaces the complete value, with zero clearing it.

## Legacy conversion table

| Legacy ID | Old name | Conversion |
|---:|---|---|
| 0 | `nada` | `None` |
| 1 | `BAJOTECHO` | `UnderRoof` |
| 2 | `trigger_2` | `AntiNpcRespawn` |
| 3 | `POSINVALIDA` | `AntiNpcRespawn`; merge both NPC behaviors |
| 4 | `ZonaSegura` | Clear completely |
| 5 | `ANTIPIQUETE` | `WayUnblocker` |
| 6 | `ZONAPELEA` | `PvPArena` |
| 7 | `AUTORESU` | `AutoResurrection` |
| 8 | `DETALLEAGUA` | `SwimSuitPath` |
| 10 | `PESCAINVALIDA` | `NoFishing` |
| 11 | `VALIDONADO` | `SwimSuitPath` |
| 12 | `ESCALERA` | Clear completely |
| 13 | `WORKERONLY` | Clear completely |
| 14 | `TRANSFER_ONLY_DEAD` | `GhostOnlyTranslator` |
| 16 | `NADOBAJOTECHO` | Clear completely, as explicitly decided; do not infer roof or swimming flags |
| 17 | `VALIDOPUENTE` | Clear completely |
| 18 | `NADOCOMBINADO` | `SwimSuitPath`; no separate combined-swimming flag |
| 19 | `CARCEL` | Clear tile trigger; set `Prison` in the entire map's zone flags |
| 20 | `ONLY_PATREON_TILE` | Clear completely |
| 21 through 40 | Emperor castle entries 1 through 20 | Remove trigger identity; resolve through the old `castle.trigger` to `castle.id` mapping and create castle entrance bindings |
| 41 | `CASTLE_FOUNDATION_POSITION` | `CastleFoundationPosition` |
| 60 through 73, 90 through 99 | Roof group codes | `UnderRoof`; retain spatial separation through roof connectivity, not trigger values |
| 200, 201 | Undocumented marker and `BLOQ15` | Clear completely |

IDs 9 and 15 are not defined by the server. Any unknown value, including an unlisted value above 60, must appear in the conversion report and stop automatic conversion of that map until explicitly classified. Do not interpret arbitrary values as combinations or generalize the old `>=60` rule to unknown data.

The local baseline has 773 `.csm` maps. Relevant audit counts include 33,227 tiles with ID 2, 6,128 with ID 3, 990 with ID 19 in map 66, 150 with ID 16, five with ID 200 and eleven with ID 201 in map 320. ID 41 occurs on 21 tiles across 21 maps. No stored ID 20 or 21 through 40 was found in that baseline; castles create their entry triggers at runtime. Recompute counts and checksums against the actual release asset revision before converting.

Clearing a trigger removes its dedicated behavior. It does not delete graphics, objects, tile exits, or physical collision flags. In particular, removing the bridge exception may expose underlying water restrictions; removing a safe tile leaves existing map-wide safety rules in effect. Report these maps for gameplay review instead of silently inventing replacement flags.

## Behavior after conversion

`AntiNpcRespawn` prevents NPC spawning and ordinary NPC movement/pathfinding onto the tile. This intentionally adds the old ID 3 movement restriction to tiles that previously had ID 2. Preserve explicit pet and movement-ignore exceptions where the owning NPC functions already provide them. It has no rendering behavior.

`WayUnblocker` retains the current obstruction timer, warnings, reset behavior, and eventual disconnect. `AutoResurrection` retains resurrection and healing. `NoFishing` only controls fishing. `GhostOnlyTranslator` retains the dead-only tile-transfer rule and is evaluated before executing that transfer.

`SwimSuitPath` replaces the separate path categories for rubber suits, ordinary swimming suits, and combined swimming. Use one path predicate for movement, login equipment selection, and equip/unequip checks. Both existing suit families should qualify for this path during the transition; retain an already equipped valid suit and use a deterministic existing inventory order when selecting one at login. Removing item definitions or changing suit bonuses is a separate item-system change. Preserve water on converted ID 8, 11, and 18 tiles explicitly, because the old loaders restore `FLAG_AGUA` from these IDs. Do not restore water implicitly from cleared ID 16.

Remove numeric-range rules such as `trigger < 12`, `trigger > 10`, and `trigger < 50`. They currently mix NPC spawning, mounting, weather exposure, and warp placement with unrelated trigger IDs. Implement each rule from the relevant named property, physical terrain, tile exit, or existing map setting. Do not carry incidental restrictions into the new flags solely because a historical number happened to satisfy a range. Under-roof weather shelter comes from `UnderRoof` plus existing map environment settings. Any additional mounting policy needs an explicit design decision, not another ordering dependency.

## Arena checks and removal of the legacy result enum

`e_Trigger6` is actively used. Its consumers include `SistemaCombate.bas`, `Modulo_UsUaRiOs.bas`, `modHechizos.bas`, `Trabajo.bas`, and `InvUsuario.bas`. Remove the enum and `TRIGGER6_*` names only when all callers use predicates derived from `PvPArena`.

| Source in arena | Target in arena | Arena contribution |
|---|---|---|
| No | No | Apply normal combat rules |
| Yes | Yes | Apply arena combat rules |
| Yes | No | Arena boundary crossing |
| No | Yes | Arena boundary crossing |

Use small predicates such as `IsInPvPArena`, `BothInPvPArena`, and `CrossesPvPArenaBoundary`. Preserve existing explicit map-wide `SafeFightMap` overrides in the owning combat policy. Theft remains prohibited if either participant is in an arena. Death, drop, citizenship, and equipment callers that compare a user with itself use the single-user predicate. Healing must evaluate the boundary predicate before using its result; the current `modHechizos.bas` branch assigns a local result in one branch and examines it in another.

Two arena tiles can have different additional flags and still allow arena combat. Never compare their complete masks for equality.

## Map-wide prison property

Add `ZoneFlags As Long` to server map information and a matching 32-bit field to HOO map metadata. Define PascalCase zone members independently of tile flags: `None = 0`, `Prison = 1`.

If any legacy tile has ID 19, set `Prison` for the entire map and clear all ID 19 tiles. For the scanned baseline, map 66 becomes a prison throughout its bounds. Replace every former jail-tile restriction with a check of the current map's `Prison` property, including lobby, challenge, teleport, item, and command restrictions. This expansion from painted prison tiles to the whole map is intentional. Existing textual map zones and safety settings remain independent fields in the initial migration.

## Roof detection and fading

All retained roof markers become `UnderRoof`. HOO must stop using the trigger's numeric value as the identity of a roof. Derive temporary component IDs from roof connectivity; those IDs are renderer cache data and are never trigger flags or persistent castle-style identities.

1. Build a roof coverage grid using the same roof-layer placement, sprite footprint, and anchor rules used by the renderer. Associate coverage with authored `UnderRoof` areas. Graphics extending over multiple tiles must participate in the coverage of those tiles; a graphic's anchor tile alone is insufficient.
2. At the local character's interpolated position, compute the world-space head anchor from the character's actual rendering offsets. Probe the roof coverage above that anchor using the same world-to-map transform. Do not search for the nearest roof or select a roof merely because it is in the camera view.
3. If the probe has no qualifying roof coverage, fade the previous roof back in. Otherwise use the containing roof cell as the flood-fill seed.
4. Flood only four-way adjacent `UnderRoof` coverage belonging to the roof surface, with a visited set and a bounded iterative queue. Never cross an empty cell, map boundary, or explicit roof seam. Diagonal contact alone does not join roofs.
5. Fade only the render geometry associated with that component. Restore the previous component when the head probe moves to another roof or leaves roof coverage. If one render primitive spans separated components, split its render coverage or fix the asset; assigning it to both components would fade another roof.
6. Cache component labels on map load and invalidate affected labels when roof graphics, `UnderRoof`, or seams change. Recompute the selected component when the head probe changes; do not flood the entire map every frame or introduce another application loop.

A flood fill cannot distinguish two independent roofs if their coverage is connected without a boundary. Store sparse `RoofSeams` between adjacent cells where touching roofs must remain independent. During conversion, adjacent cells with different retained legacy roof-group IDs can establish seams; the converter must flag ambiguous geometry for map-author review. Disconnected areas with the same old ID naturally become independent components. No legacy group number is needed at runtime.

Seams describe spatial boundaries, not new trigger types. ID 4 and ID 16 do not create `UnderRoof` during conversion. Any subsequent repainting of those areas is an explicit map edit.

## Castle map references

Each map owns a sparse `CastleEntrances` dictionary from a bounds-checked tile key to a stable database `castle.id`. Allocate the dictionary only for maps with bindings. Keep ownership, whitelists, names, dates, and item links in `CastleData`; map data carries only the reference.

For an outside anchor `(x, y)`, the current castle builder creates four restricted approach tiles: `(x-1,y)`, `(x-2,y)`, `(x-1,y+1)`, and `(x-2,y+1)`. Register these four bindings after validating the complete footprint. Existing portal positions and `TileExit` destinations remain separate from the access-check tiles.

Resolve the castle by the entrance's ID, then check that castle's placement, activity, ownership, and whitelist before transfer. Owner lookup is a separate operation used when finding a player's castle. Remove the current owner-or-trigger search: it can select a player's own castle when that player is entering another one.

Creation, relocation, destruction, and map reload must update these bindings without clearing unrelated tile flags. Validate a relocation before publishing it, remove all old entrance bindings, and register the new ones with the matching portal changes. Unknown or invalid castle references deny access with an actionable log. Whitelists can store normalized character names as membership; their values no longer need to contain an old trigger number.

## Production castle baseline

The production screenshots supplied on 7 October 2026 establish 20 castle rows and 15 coordinate rows. Use a transactional database snapshot as migration input; copy full values from that snapshot rather than retyping the screenshots. The first castle name is clipped in the image, and whitelist rows were not included.

| Castle ID | Legacy trigger | Spawner object | Inside key object | Coordinate row ID | Interior map |
|---:|---:|---:|---:|---:|---:|
| 1 | 21 | 6362 | 6383 | 16 | 758 |
| 2 | 22 | 6363 | 6384 | 17 | 759 |
| 3 | 23 | 6364 | 6385 | 18 | 760 |
| 4 | 24 | 6365 | 6386 | 19 | 761 |
| 5 | 25 | 6366 | 6387 | 20 | 762 |
| 6 | 26 | 6367 | 6388 | 21 | 763 |
| 7 | 27 | 6368 | 6389 | 22 | 764 |
| 8 | 28 | 6369 | 6390 | 23 | 765 |
| 9 | 29 | 6370 | 6391 | 24 | 766 |
| 10 | 30 | 6371 | 6392 | 25 | 767 |
| 11 | 31 | 6372 | 6393 | 26 | 768 |
| 12 | 32 | 6373 | 6394 | 27 | 769 |
| 13 | 33 | 6374 | 6395 | 28 | 770 |
| 14 | 34 | 6375 | 6396 | 29 | 771 |
| 15 | 35 | 6376 | 6397 | 30 | 772 |
| 16 | 36 | 6377 | 6398 | Missing | Missing |
| 17 | 37 | 6378 | 6399 | Missing | Missing |
| 18 | 38 | 6379 | 6400 | Missing | Missing |
| 19 | 39 | 6380 | 6401 | Missing | Missing |
| 20 | 40 | 6381 | 6402 | Missing | Missing |

All 15 existing interiors have `(x,y) = (50,72)`. Coordinate primary keys 16 through 30 are not castle IDs; join using `castle_coordinates.castle_id`.

| Castle ID | Owner account | Owner character | Foundation date | Active | Outside map and anchor |
|---:|---:|---:|---|---:|---|
| 1 | 19241 | 15283 | `2026-07-25 18:40:40` | 1 | Map 27, `(49,69)` |
| 2 | 9 | 2897 | `2026-08-11 18:46:53.434` | 1 | `(0,0,0)` sentinel |
| 3 | 4147 | 1613 | `2026-08-31 21:13:43` | 1 | Map 546, `(50,30)` |
| 4 through 20 | NULL | NULL | NULL | 0 | IDs 4 through 15 have zero sentinels; IDs 16 through 20 have no coordinate row |

Preserve every castle ID, owner reference, object reference, activity value, name including NULL, and foundation timestamp including fractional seconds. The visible name of castle 3 is `Thunder Hold`; castle 2's name is NULL. Do not invent the clipped name of castle 1.

Castle 2 remains owned and active with an interior but no placement. Treat it as unplaced and create no outside entrance, without resetting its owner, date, or activity. Castles 16 through 20 remain unconfigured slots: keep their castle and object rows, but prohibit placement until an interior is explicitly assigned. Do not infer maps 773 through 777 or manufacture coordinate rows.

Expected runtime entrance bindings from this baseline are exactly:

| Map | Castle ID | Restricted tiles |
|---:|---:|---|
| 27 | 1 | `(48,69)`, `(47,69)`, `(48,70)`, `(47,70)` |
| 546 | 3 | `(49,30)`, `(48,30)`, `(49,31)`, `(48,31)` |

These eight bindings are rebuilt from production placement data at server startup. Do not bake production owners or movable castle placements into the shared asset repository.

## Castle database schema and migration

Use the following target structure, retaining the existing SQL naming style:

- `castle`: existing stable primary key, owner references, spawner and inside-key references, foundation date, activity, and name; remove `trigger` after all references have been converted.
- `castle_coordinates`: preserve existing `id` and unique `castle_id`, preserve non-null interior coordinates, and normalize an all-zero outside triple to an all-NULL outside triple. Require outside fields to be either all NULL or all valid positive coordinates. A castle may have no coordinate row while unconfigured.
- `castle_whitelist`: preserve IDs, names, and `castle_id` foreign keys. The table already refers to castle IDs and requires no trigger-number rewrite.

Keep `castle.id` independent of array order. Both loaders and writers must resolve by ID, including the current coordinate loader and `SaveCastleDataToDb`, which use loop indexes. Use `Long` for database IDs and owner IDs in VB6 rather than limiting future records to `Integer`.

Implement the database operation as a versioned migration with these steps:

1. Stop castle writes for the migration window and take a consistent SQLite backup. Inventory the actual schema, foreign keys, indexes, whitelist rows, and any other references to the castle tables.
2. Build a temporary legacy-trigger-to-castle-ID mapping from the actual `castle` rows. Verify unique IDs, trigger values, and item references. The supplied baseline satisfies `trigger = id + 20`, but the migration must use the recorded mapping rather than rely on this formula.
3. Preflight coordinate bounds and parent references. Classify castle 2 as owned and unplaced and castles 16 through 20 as unconfigured. Partial zero triples, orphan references, duplicate bindings, or invalid nonzero coordinates require correction before publication.
4. In one transaction, create and populate replacement tables using explicit `INSERT ... SELECT` columns. Preserve IDs and raw stored timestamp/name values. Convert only `(0,0,0)` outside coordinates to `(NULL,NULL,NULL)`. Preserve all 20 castles and exactly the 15 existing coordinate rows.
5. Use SQLite's table-rebuild procedure on the installed SQLite version, including preservation of indexes and inbound foreign keys. If foreign-key enforcement must be disabled for the rebuild, do so before beginning the transaction, never rely on changing it inside a transaction, and restore it afterward. Do not use a destructive historical castle seed script: it would erase the owners and dates shown above.
6. Verify foreign keys, bidirectional row comparisons excluding the intentionally removed `trigger` column, normalized coordinate comparisons, whitelist equality, and primary-key/autoincrement continuity before committing. Record the schema migration version. A repeated run validates the target version and performs no second conversion.
7. Rebuild the runtime entrance dictionaries and verify the eight expected bindings. Save and reload every retained castle state by stable ID. Keep the legacy mapping in the migration report for audit and rollback; gameplay no longer consults it.

Rehearse on a production backup. Rollback restores the pre-migration database together with the matching old binaries and assets; do not attempt a lossy reverse conversion after new combinations or relocations have been saved.

## Versioned binary map schema

Introduce a distinct `CSM3` signature. Do not widen a field inside the existing four-layer or five-layer format without changing its signature. Keep readers for the legacy formats during migration, using the conversion table to produce the same new in-memory representation. New writers emit only the new format.

All numeric fields use little endian, with explicit field widths and no implicit compiler padding or VB6 whole-UDT serialization. The proposed version 1 header is:

| Field in order | Encoding |
|---|---|
| Magic | Four bytes `43 53 4D 33` (`CSM3`) |
| SchemaVersion | UInt16, value 1 |
| LayerCount | UInt16, value 5 |
| ZoneFlags | UInt32, initially only `Prison` supported |
| BlockedCount | UInt32 |
| LayerCounts | Five UInt32 values |
| TriggerCount | UInt32 |
| LightCount | UInt32 |
| ParticleCount | UInt32 |
| NpcCount | UInt32 |
| ObjectCount | UInt32 |
| TileExitCount | UInt32 |
| CastleEntryCount | UInt32 |
| RoofSeamCount | UInt32 |
| Bounds | Existing four Int16 fields in order XMax, XMin, YMax, YMin |
| Metadata | Existing metadata fields, order, and string encoding, serialized explicitly |

After metadata, retain the existing section order: blocks, graphics layers 1 through 5, triggers, particles, lights, NPCs, objects, and tile exits. Append castle entries and roof seams. Preserve all unaffected record fields and their existing byte encodings.

| Changed or new record | Fields |
|---|---|
| Trigger | `x: Int16`, `y: Int16`, `flags: UInt32`; 8 bytes instead of the old 6 |
| Castle entry | `x: Int16`, `y: Int16`, `castleId: Int32`; 8 bytes |
| Roof seam | `x: Int16`, `y: Int16`, `direction: UInt8`; 5 bytes; 0 means east neighbor, 1 means south neighbor |

Store each seam once using east or south orientation. Both adjacent cells must be in bounds and have roof coverage. Omit zero-valued trigger records. Reject duplicate coordinates in a single trigger or entrance section, invalid seam directions, unknown mandatory flags/version, truncated sections, impossible counts, and positive castle IDs with no valid binding source. Validate counts before allocation and use checked size arithmetic.

The castle section holds authored static references, when present. The current asset baseline is expected to have zero such records. The server overlays dynamic production placements from the database; a static/dynamic conflict must be reported instead of silently choosing one. HOO does not need ownership or whitelist data to load a map.

The assets converter must support a dry run and separate output directory. Convert every map from its detected source version, preserve all unrelated content, and emit per-map source/output checksums, source version, old-value counts, new-flag counts, zone changes, seams, entrance bindings, and rejected values. Never overwrite the source batch before both new readers accept the complete result. An already converted file must not pass through the legacy numeric mapping again.

Update map editors, exporters, asset packers, fixtures, and caches that assume a six-byte trigger entry or the old header. HOO's `CsmTriggerEntry`, `WorldMap`, adjacent-map/environment handling, tutorial/path queries, and roof caches must all consume the same new representation. Server-generated or saved maps must use the same writer contract. Map graphics and independent collision flags remain separate structures.

## GM command and live edits

HOO registers `/TRIGGER [value]` as a Game Master command with case-insensitive lookup, console completion, and localized help. An omitted argument queries the current tile. A decimal integer from 0 through 2147483647 sets the value; negative values, overflow, fractions, hexadecimal, and extra arguments are rejected. The current command transport preparation accepts that numeric range; after the gameplay migration the server also rejects bits outside the known mask.

| Operation | Wire contract |
|---|---|
| Set | Existing client packet ID 164 (`eSetTrigger`) as Int16, followed by one Int32 value; 6 bytes total |
| Query | Existing client packet ID 165 (`eAskTrigger`) as Int16; 2 bytes total |

Use server `ReadInt32`, `Long` locals, and HOO Int32 serialization. Keep server authorization authoritative and retain existing GM audit logs and query messages. Both the set path and query path must handle values above 255 without truncation. No action counter or position is appended; the server uses the authorized GM's current position.

The preparatory command still writes the existing server trigger field; it does not reinterpret the existing enum or binary maps as flags. Its server confirmation does not currently refresh HOO's local map tile. In the full migration, add an authoritative typed tile-properties event carrying map, coordinates, and the 32-bit mask. Route it through HOO's game-event handler into world state and invalidate roof connectivity when required. Do not optimistically change local movement or roofs before server acceptance. This live-edit event must be implemented before using the migrated command as a synchronized map-editing tool.

Changing the set payload from one byte to four is a protocol compatibility change. Deploy the matching server and HOO revisions together and make the release/version check reject incompatible clients. Never guess the payload version from its numeric value or silently accept both lengths on an unversioned stream.

## Implementation order and acceptance

1. Review the implementation specification in three documentation-only PRs. Deliver the preparatory GM-only command in two separate HOO/server PRs, with tests for parsing, permission checks, packet bytes, large values, query behavior, and disconnected handling.
2. Freeze the bit values and new binary header. Add shared binary fixtures and both new readers, retaining explicit legacy conversion support. Update all runtime mask widths and replace equality/range consumers before activating flag data.
3. Implement arena predicates, combined NPC restrictions, swimming paths, and map-wide prison checks. Test every removal and intentional behavior change, including ID 16 clearing.
4. Implement castle ID resolution, dictionary lifecycle, and the database migration. Rehearse with the supplied 20/15-row production shape and a full whitelist export, including shuffled and non-contiguous IDs.
5. Implement HOO roof coverage, seam-aware flood fill, caching, and authoritative live tile updates. Test separate roofs sharing an old ID, touching roofs separated by seams, diagonal contact, head movement under a roof overhang, multi-tile graphics, component changes, and map unload/reload.
6. Convert the complete assets batch, update authoring tools, and compare both readers' normalized results. Rebuild packs and publish a manifest pinning server, HOO, assets, map format, and database schema versions.
7. Run native Windows builds and focused automated tests, then staged gameplay checks. Deploy with a consistent database backup and matching release artifacts. Do not migrate the production database as part of authoring this specification.

Acceptance requires all 773 baseline maps to parse after conversion, all retained ordinary map data to match, all expected removed/merged trigger values to follow the conversion table, map 66 to be prison-wide, exactly eight initial dynamic entrances for castles 1 and 3, no fabricated placements for castle 2 or interiors for castles 16 through 20, and isolated roof fading above the local character's head. Adjust the expected map count only if the release inventory has legitimately changed and the manifest records it.

## Source locations

- Server: `Codigo/Declares.bas`, `FileIO.bas`, `GameLogic.bas`, `MODULO_NPCs.bas`, `Modulo_UsUaRiOs.bas`, `SistemaCombate.bas`, `modHechizos.bas`, `Trabajo.bas`, `InvUsuario.bas`, `ModCastle.bas`, and `Protocol_GmCommands.bas`.
- Castle history: `ScriptsDB/20260624-03-create castle whitelist table.sql`, `20260624-04-create castle coordinates.sql`, and the July 2026 castle scripts. Use them to understand the schema, not to reseed production.
- HOO: `source/pymmoclient/assets/csm_map.h`, `assets/csm_parser.cpp`, `world/world_map.*`, `client/game_command_*`, `client/game_console_input_controller.*`, `client/ingame_ui_flow.cpp`, and `networking/game/vb6_protocol.*`.
- Assets: `Mapas/*.csm` and `init/triggers.ini`; inventory additional map readers/writers and packed copies before release.
