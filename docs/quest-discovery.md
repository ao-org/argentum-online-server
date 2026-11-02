# NPC quest discovery

The HOO companion displays an informational gold exclamation on disclosed NPCs
with an eligible offer. Discovery is independent of active objectives, turn-in
routes and target inspection; no live NPC positions or prerequisite details are
sent. Existing NPC interaction and acceptance remain authoritative.

## Acceptance rules

`ModQuest.QuestAcceptanceFailure` is the read-only predicate extracted from
`CanUserAcceptQuest`. The existing acceptance wrapper retains the same ordered
localized rejection messages and arguments. Worker/class restrictions, already
active membership, completed prerequisite, minimum/maximum level, required skill,
faction mask, non-repeatable completion, and global threshold/active-event/boss
rules are shared. Repeatables remain unavailable while active and can become
available after completion. `FreeQuestSlot` is checked by discovery just as in
`HandleQuestAccept`. Distance is deliberately evaluated by the existing NPC
interaction (and revalidated at acceptance), not the character-wide snapshot.
No rewards, costs, eligibility conditions or acceptance side effects were added.

`NpcInfoCache` is the authority for offering template/quest relationships, also
copied into live NPCs by the existing creation path. A lazy pre-index is rebuilt
when quest/NPC definitions reload. It holds quest-to-provider relationships and
only the relevant skill/global dependencies. There is no per-visible-NPC query,
per-hover subscription, disk I/O or database work during publication.

## Updates

Publication reuses `modQuestTurnIn`'s session-bound dirty-user dictionary and
end-of-network-poll drain (128 users per drain, 250 ms retry floor). Existing
quest/inventory/skill mutations already mark it. Level/class stat writers,
faction refresh/council transitions and completed-history mutations now mark it
too. Login/resync forces publication; logout and slot reuse invalidate old work.

A small input signature contains level, class, faction, active slots, completed
history count, relevant skill values and definition/global-gate revision. When
unchanged it skips eligibility scans entirely; changed inputs with identical
availability send no packet. Global changes compare pre-indexed threshold and
event/boss eligibility gates once before queuing users. Below-threshold
contributions do not reevaluate offers. Discovery-only global invalidations do
not reevaluate unrelated turn-in readiness; multiple changes coalesce in the
same dirty entry. Gameplay acceptance is still immediate and revalidated.

## Required coordinated protocol

New `eQuestDiscovery` is ID **226 with PYMMO=1**, 227 with PYMMO=0's existing
legacy enum insertion. HOO production integration uses PYMMO=1.

```text
Int16 packetId
Int32 sequence (>0, scoped to user session)
Int16 page (0..255)
Byte flags (1=final, 2=projection unavailable)
Int16 offer count (0..128)
repeat count:
    Int16 NPC template (>0)
    Int16 quest ID (>0)
    String8 title (at most 128 source characters; existing 7-bit length prefix)
```

Pages form a full atomic replacement, starting at page zero. Empty final page
zero clears availability. Partial enqueue failure leaves work pending; retry
resumes at the first unqueued page. At most four pages are sent per drain; changed
eligibility supersedes a pending transaction with a newer sequence. The sent cache advances only after every page queues.
The client rejects stale/duplicate completed sequences, missing pages, invalid
IDs/counts and duplicate template/quest pairs. It keeps the previous complete
snapshot while receiving a replacement and clears all state at session reset.

VB6 authored quest IDs are signed Int16 (maximum 32767). Each NPC's offer count
is a Byte (maximum 255). Pages carry at most **49931 payload bytes** (conservative UTF-8 encoding bound) and a
snapshot supports **32767 relationships**, with a pre-index limit of 131072.
The client validates 384 wire bytes / 512 normalized UTF-8 bytes per title and
at most 255 offers per NPC. Larger authoring/snapshots send an explicit unavailable empty replacement and
log the capacity error; no partial set is silently published or left stale.
Current local assets audit: 427 quest definitions, 173 unique valid offer links,
102 providers, comfortably inside these bounds. This is source/data validation,
not production runtime verification.

Deploy matching HOO and server builds together, before accepting game sessions.
An old HOO does not understand packet 226. There are no capability flags,
negotiation, compatibility fallback or version-enforcement changes, and no
ProtocolGenerator or assets update. Keep both PRs draft; do not merge/deploy
without the coordinated release.

## Validation and manual checks

Eleven deterministic cases are registered as suite 49: shared eligibility,
worker/class/level/skill/faction/prerequisite restrictions, active/completed and
repeatable quests, slot capacity, global gates, coalescing and unchanged
suppression, enqueue retry, slot reuse, and multi-page retry/layout. Existing
quest/turn-in tests remain enabled. New/touched VB6 sources use Windows-1252 and
CRLF. Packet IDs were source-checked in both conditional configurations.

The documented native VB6 compiler is unavailable locally. The existing native
PR validator compiled and passed **920/920 tests** at code commit
`bf8656559968c08ae56d188244c39e86e76e4274` (36,637.43 ms), including all eleven
new discovery cases and broader quest/turn-in regressions. The job does not expose
an independent PYMMO=0/1 native matrix; that matrix and live gameplay remain
outstanding. Do not read source checks as a separate native configuration pass.

Companion HOO PR: https://github.com/morgolock/heroes-of-old/pull/850.
Client code `68b9c39651523da1a244f9f51229b945dac5d046` passed native Windows
Release client/unit builds and 321/321 focused tests (61 suites, 14.488s).
EN/ES classic/modern-coordinate GL fixtures were inspected; they are labelled
synthetic data, not live production eligibility or stitched multiplayer checks.

Manual check using authored NPC 19 (Lyra), quest 320 (Resonancia anomala): use a
level 13..24 character with an available slot, no active copy, and verify `!`.
Accept through ordinary NPC interaction: objective marker wins. Complete/turn in
and verify readiness wins until successful turn-in; the repeatable offer can
return afterward. Move outside the level range or fill all five quest slots and
check availability clears. Also test a prerequisite-locked/non-repeatable quest,
level/skill/faction changes, reconnect and another character. In HOO check both
UI modes, adjacent maps, hidden/roof-covered NPCs, clipping/HUD occlusion, marker
setting, hover names and prepared magic/bow press/release safety.

## Targeted missing-marker diagnosis

Set `AO_TRACE_QUEST_DISCOVERY_NPC=432` before server startup to trace a specific
provider template. The environment is read once. At most64 log records cover index
construction, matching quest IDs, the shared acceptance failure enum/free slot,
failed enqueue and final committed sequence/row count. IDs only are logged; no
extra player packets, DB reads, subscriptions or global scans are introduced.
Unset and restart to disable/reset the budget. Diagnostics do not change eligibility.

Authored assets at `78d1253e6e662a53a78e58389bd4514cd58d0c77`: Alaric NPC432 offers
231 (minimum1) and230 (minimum1, completed231 prerequisite); Heather415 offers38
(minimum13); NPC408 (Romina in Spanish, Makrom in current English) offers427
(level1..8). No HP fields are authored for these noncombat providers. Controlled
fixtures use those exact acceptance/provider fields, not inferred subtitles.
The current dialog Available calculation is partial; actual acceptance and discovery
also enforce all shared eligibility checks and available accepted-quest capacity.

New tests cover initial publication with zero accepted quests, eligible231 alongside
blocked230, Heather/Romina level limits, production HandleQuestAccept invalidation,
failed-enqueue retry and reconnect. Native source `1c8085b5` passes **1013/1013**
full validator tests (71,012.87 ms), including all15 discovery cases. Local VB6
compilation and an independent conditional-build matrix are unavailable.

The companion HOO fix addresses a reproduced client-only 0/0-health update being
misclassified as death. Initial creation alone does not reproduce it; no live
Seneca/Alaric session has confirmed the reported login case. Collect the server
trace and matching `HOO_TRACE_QUEST_MARKERS=1` client page/commit/actor trace before
opening the quest dialog to identify the exact failing stage. No protocol or
coordinated deployment is required for this optional diagnostic addition.
