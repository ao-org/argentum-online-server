# Immediate tamed-pet ownership refresh

`Trabajo.DoDomar` publishes the existing tamed NPC to its owner through
`MakeNPCChar(False, UserIndex, NpcIndex, map, x, y)` after ownership assignment,
FollowAmo and the unchanged ReSpawnNpc operation. This sends the same character
index with EsNPC=2 to that owner only. Other observers retain ordinary NPC
semantics. ReSpawnNpc's replacement wild creature remains distinct and unowned.

The refresh precedes the existing success message and the unchanged correlated
packet 207 emitted by HandleWorkLeftClick. Failed rolls and rejected attempts do
not send it. There is no new packet, field, capability or client-side ownership
inference. Gameplay eligibility, probability, pet capacity and persistence remain
server-controlled.

On NoMascotas maps the active refresh is omitted and the existing successful
tame storage/removal path runs. Its condition now tests the Boolean directly:
FileIO loads True as -1, so the old `= 1` comparison skipped that path. The test
verifies success, preserved recorded pet type/count and no active phantom NPC.

The owner-only refresh is representation-only. The Boolean correction does
affect gameplay on restricted maps: it restores the already-intended NoMascotas
removal/storage restriction. This is a bug fix, not a balance or mechanics
redesign. Neither change alters tame probability, power or pet capacity.

UnitTesting suite 40 captures the real CharacterCreate serialization and checks
ordinary-to-owned same-character refresh, owner/non-owner isolation, failed roll,
rejection, NoMascotas and a real replacement wild spawn. It uses existing test
registration and SendData capture infrastructure, not a separate runner.

The SendData capture call is guarded by `UNIT_TEST = 1`.
The capture implementation is inside UnitTesting's `UNIT_TEST = 1` block.
Aurora.Network is now unconditional. Production `UNIT_TEST=0` excludes both;
there is no runtime production capture or test behavior.

Historical validation before transport removal (2026-10-19; see
[NetworkTransport.md](NetworkTransport.md) for current commands):

- DEBUGGING=1:PYMMO=1:UNIT_TEST=1:DIRECT_PLAY=0: build passed; full suite 427/427.
- DEBUGGING=0:PYMMO=1:UNIT_TEST=0:DIRECT_PLAY=0: build passed.

Tests ran in an isolated runtime directory with a fresh local DB, existing
public encryption fixture and localhost-only test port. The user's live/local
database and configuration were not used. Live multiplayer validation remains
separate. Deploy this refresh before relying on HOO's immediate pet-command
tutorial flow; ProtocolGenerator and packet layouts remain untouched.
