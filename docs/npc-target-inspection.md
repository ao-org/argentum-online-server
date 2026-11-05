# HOO target inspection

Required HOO/VB6 integration; no feature negotiation or ProtocolGenerator.
HOO selects an active-map NPC by character index and local lifetime generation.
Adjacent NPCs currently are visual-only representations, not combat selections;
no adjacent-map query is admitted. Border handoff creates a fresh subscription.

## Wire

Client packet `317`: Int32 positive request, Int16 viewer map,
Int16 character index, Int16 NPC template. Character zero unsubscribes.
Server `221`: echoed request/map/character/template, Byte available (0/1),
Byte threat (0 unknown, 1 low, 2 challenging, 3 dangerous, 4 deadly),
Byte offensive trait bits (1 poison, 2 projectile, 4 spells), Byte entry count
(low 7 bits, maximum 32; bit 7 flags omitted overflow), then entries: Int16 quest ID, Byte kind (1 kill, 2 ready turn-in, 3 click-to-visit),
Int32 current, Int32 required. Turn-in counts are zero.
Deploy matching server and HOO versions together. No existing packet is resized.
An older server cannot process the new request; old-server compatibility is not
advertised. New server subscriptions are opt-in by request, so legacy clients
receive no new messages. No database, assets or persisted player fields change.

## Authority and lifetime

One subscription per selected character lifetime. The server retains user ID,
user version, NPC slot/version, character index, NPC template and viewer map.
Only living, active, visible NPCs in the viewer's existing loaded-area policy
qualify, including the existing dead-viewer/reviver exception. A stale reference
sends unavailable and retires the subscription. Nothing changes gameplay targets.
Snapshots update at most every 250 ms and send only on change. Kill progress,
quest acceptance/removal also refresh immediately. No database query or random
roll occurs. Failed enqueues do not advance the comparison cache.

Kill objectives match `RequiredNPC.NpcIndex = Npc.Numero` exactly, as
`OnNpcKilledUpdateQuest` does, and disappear when current reaches required.
Turn-ins match `TalkTo` or the NPC's `QuestNumber` list, and use the existing
`FinishQuestCheck` (items, kills, spells, skills, target objectives and global
quest conditions with notifications suppressed), plus class/reward admission in `CanInspectQuestTurnIn`. `TalkTo` is not a separate talk-progress objective. Real click-to-visit objectives use `RequiredTargetNPC` and `NPCsTarget` (kind 3).

## Threat: intentionally limited estimate

For hostile attackable non-pet physical NPCs, subtract average equipped armour
from maximum physical hit: helmet weight 1/6; armour + shield weight 5/6, following
`NpcDamage` body-part probabilities. Clamp estimated residual damage to at least
1. Divide the residual by player maximum HP and normalize attack interval to
2000 ms. Bands: below .06 Low; [.06,.15) Challenging; [.15,.30) Dangerous;
at least .30 Deadly. Neither current HP nor random rolls enter the calculation.
Equipment, maximum HP and NPC attack interval updates naturally refresh it.

This estimates incoming physical pressure in a normal one-on-one encounter,
not the outcome. It excludes evasion, player damage/kill speed, consumables,
critical effects, poison damage, temporary buffs, ship/mount defence, cards and
special AI. Spells cannot be responsibly reduced to this physical estimate:
spellcasters report unknown and retain their Spellcaster trait. Missing data,
friendly/noncombat NPCs and owned summons are never defaulted to Low.

## Validation

Suite 45 exercises real subscriptions/packet writes, exact quest matching,
completion versus turn-in, multiple quests, invalid lifetimes/maps/visibility,
changed enqueue failure and stable current-HP-independent threat. Pure threshold
checks are included, along with conditional packet-ID assertions and absence of
quest-chat side effects during inspection.

The native PR validator built source commit `a572fedc7ba0c1e85743c73f260f7c69731198fc`
and reported **857/857 passed, zero failures**. The VB6 compiler is unavailable
locally; this is CI evidence, not a local native run. The report does not establish
a complete native configuration matrix; that remains unverified.
Live companion-client gameplay remains outstanding.
