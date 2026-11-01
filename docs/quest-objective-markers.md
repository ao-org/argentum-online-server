# World-space quest objective markers

The server publishes accepted, unfinished kill and click-to-visit requirements.
HOO uses exact NPC template numbers, not names, sprites, positions, authored
placements or selected-target subscriptions. Inspection and gameplay are unchanged.

## Required protocol

Coordinate this server change with its HOO companion. No capability negotiation,
ProtocolGenerator, assets, database changes or new client requests are involved.
Deploy the client first, then the server during the coordinated rollout. The new
client shows no kill markers until authoritative eligibility arrives; an older
server cannot supply the complete projection. Older clients are unsupported:
they cannot decode the extended adjacent snapshot from the new server.

- `eQuestObjectiveTemplates`: 223 in PYMMO=1, 224 in PYMMO=0.
- Little-endian Int16 packet ID, Int16 count (0..2550), then count signed Int16
  requirements. Positive N means unfinished kill objective for template N;
  negative N means unfinished click-to-visit objective for template abs(N).
  Zero and -32768 are invalid. Both signs may be present for the same template.
  This is a complete replacement, including an empty snapshot.
- Maximum payload including ID is 5104 bytes: five accepted quests, two lists
  of at most 255 objectives. Duplicate requirements are removed server-side.
- Existing `eUpdateCharValue` gets type 2 (`eQuestKillEligible`): existing Int16
  character index, Int16 value type, Long Boolean. Sent after successful NPC
  creation and on taming to existing viewers. It does not alter ownership,
  collision, combat or targeting.
- `eAdjacentNpcUpsert` appends one Boolean kill-eligibility byte after its name.
  The appearance fingerprint includes eligibility, so taming updates adjacent
  viewers through the existing bounded visibility stream. HOO treats absence as
  unknown/ineligible; existing creation lifetime/scene-generation rules apply.

## Authority and updates

`ModQuest` supplies actual accepted slots and per-objective counters. Kill and
visit entries remain while current count is below required count; incomplete
metadata clears the projection instead of producing guessed partial results.
The event is sent at login, acceptance, kill/click progress, slot cleanup and
quest-list replacement. A per-user/version cache suppresses unchanged snapshots;
only accepted network enqueue updates that cache. Login forces a replacement.

`MuereNPC` excludes an NPC with a valid player master before kill quest credit.
Therefore kill markers also require explicit instance eligibility. This is
separate from EsNPC=2, which describes only the receiving player's pets.
Click-to-visit requirements do not inherit that kill-only restriction.

Normal character creation and the existing adjacent stream provide lifetime and
visibility scoping. No visibility is added, and no per-frame quest/network scan
or database query is introduced. Quest state survives map changes but not logout
or character/session changes; actor metadata is recreated with each actor.

## Validation

Suite 47 covers accepted versus unaccepted, completion, overlapping quests,
click-to-visit, malformed counters, exact matching, real kill progress callbacks,
accepted-enqueue caching, session reuse, the 2550-entry bound and distinct
kill/visit eligibility. Existing adjacent tests assert the added eligibility byte
and create/eligibility/respawn order, including maximum transport-frame size.

Local VB6 compiler is unavailable. Native PR validator results are recorded in
the PR, separately from source inspection. The validator does not establish a
PYMMO=0/PYMMO=1 configuration matrix or live gameplay validation.
