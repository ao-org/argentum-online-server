# The Pack Hunt: first shared party quest

This is a coordinated server/HOO/content change. No protocol generation,
capability negotiation, PYMMO subsystem or database service is added.

## Authored content

The companion assets change adds quest **428**, `PartyMinimum=2`, offered by
**Alaric von Nix (NPC432)** after his unchanged personal offers **231** (The Needle
and the Claw) and **230** (The Roar of Bruin). The objective is **NPC558 (Wild Wolf), six**.
These are authored template IDs, not name matching. The new hunt requires level7,
matching existing wolf quest166’s entry level. Wild wolves already spawn on map35
immediately north of Nix, at(65,26) and(48,33). Quest428 uses the ordinary
Quests.DAT loader, NPC432's QuestNumber list and the existing world wolf spawns.

Reward: **360 wallet gold per eligible enrolled character**, multiplied by GoldMult
and frozen when enrollment starts. This is a conservative first hunt: existing
wolf quest166 offers 360 gold plus XP for a much larger hunt; Alaric's personal
chain rewards rare Bear Mount657 and is not repurposed. The first party definition
supports one kill objective and gold only. Extra reward/content types are outside
this delivery. No fixture alone makes the production quest playable: deploy the
companion assets Dat/Quests.DAT and Dat/npcs.dat changes.

## Enrollment and kill accounting

The leader proposes through the ordinary NPC quest acceptance entry point. The
eligible roster is bounded to six validated party reward recipients and sorted by
persistent character ID. Proposing is the leader's explicit acceptance. Other
members receive the existing client question modal with the objective and reward;
its reply carries the proposal token. A generic uncorrelated question reply cannot
enroll anyone. Proposals expire after90 seconds. The leader can cancel, or Start
with at least two accepted eligible players including the leader. Start revalidates
alive/connected/session, party leader+ID, same map, distance<20, quest rules and free
slots. It freezes only those who explicitly accepted and still qualify. Unanswered,
declined and invalidated members are not enrolled; roster additions need a new
proposal. No party ID is persisted as the run identity.

`AwardNpcPartyQuests` calls `PartyQuestNpcDeath` once before its per-recipient
personal-quest/drop loop, using the saved NPC and existing validated recipient
snapshot. The credited character must belong to the run, the victim must match,
and at least two enrolled members must qualify on the NPC's map. Support qualifies
without damage. The run increments once, caps at six, and records a six-bit mask of
qualified persistent roster ordinals for each counted kill. Personal quests and
ordinary rewards retain their independent paths. Pet owner attribution remains in
the existing combat/death path; training and player-owned victims remain excluded.

Disconnected members retain progress. Regrouping in a newly created party works:
run membership is persistent character identity, eligibility is the current party.
A late joiner cannot enroll or claim. Every enrolled character sees the shared
counter; only a member present for at least one counted kill becomes turn-in ready.
Abandon removes that character's participation, leaving other claims/counters
intact. An unfinished duo can always abandon if it cannot regroup.

## Persistence and claim boundary

Migration `20261103-01-party-quest.sql` adds `party_quest_run`, fixed
`party_quest_member` records and bounded per-run `party_quest_kill` records. An
active partial unique index enforces one active run per character. A transaction
commits enrollment before acceptance is announced; each kill commits its counter,
mask and participation before publishing progress. Login rebuilds normal quest
slots from these authoritative records, even if ordinary asynchronous quest saves
lagged. Claimed membership also restores completed-quest history if an older
asynchronous save removed its normal projection.

Claim revalidates Alaric, distance/map, completion, enrollment and participation.
One synchronous transaction conditionally claims that member, credits user gold,
records quest_done and removes the accepted projection. Only after commit does
memory/HOO receive success and the actual reward receipt. Another member's claim
is untouched. Full wallet capacity leaves the claim available for retry.

The existing character-save path dispatches asynchronous writes; it is not an
atomic reward boundary. `user.party_quest_gold_total` and the loaded
`t_User.PartyQuestGoldSeen` fence older snapshots: normal save SQL stores
`snapshotGold + committedCreditTotal - snapshotSeenTotal`. A queued pre-claim
snapshot therefore cannot erase committed hunt gold, and a post-claim snapshot
does not add it twice. This does not attempt to redesign unrelated async save
ordering. If a commit/rollback outcome is ambiguous, main character saving is
blocked until a subsequent claim interaction reconciles persisted gold/run state;
it must not overwrite a possibly committed credit with stale memory.

No DB work occurs during rendering/hover, and no periodic quest/player scan is
added. Only pending proposal expirations and failed enqueues are visited at the
network update boundary. Failed publication retries at most once per500ms per
session, invalidated by logout/reuse. Run cache storage is bounded by user capacity;
only runs without an online member may be evicted and later reloaded from storage.

## Required protocol

PYMMO=1 IDs: client **320 PartyQuestAction**, server **227 PartyQuestState**.
The VB6 enums retain their existing conditional ordering for PYMMO=0.

Action: Int32 request>0; Byte action(refresh/propose/accept/decline/start/abandon);
Int32 proposal/run identity; Int16 quest. Duplicate request IDs cannot repeat a
mutation; stale proposal/run identities are rejected.

Snapshot: Int32 revision, acknowledgement; Byte result, phase; Int32 proposal, run;
Int16 quest, provider, target, required, progress; Int32 gold, local character,
leader character; Bool can-propose/start/claim; Byte count0..6; then each row:
Int32 persistent character; String8 name; Byte vote; Bool participated; Byte status.
It atomically replaces client state. No distant live NPC positions are transmitted.
Normal accepted quests, objective snapshots, turn-in eligibility, reward receipts,
auto-pins and guidance remain the established maintained protocol.

Deploy matching server, HOO and content together with logins closed. Back up the
SQLite database before applying the normal automatic migration. Older binaries do
not understand these required packets and are not a supported mixed release.

## Validation / manual acceptance

Validation results are recorded in the draft PRs, including the tested commits.
Native Windows VB6 execution is delegated to the existing PR validator because the
local compiler is unavailable. Source and SQLite checks alone are not native VB6
or live gameplay validation.

Two-client checklist (requires the coordinated server/content):

1. At Alaric, verify quests230/231 retain their original eligibility, and428 shows
   Party · 2+ players. A solo character gets an explanation instead of enrollment.
2. Create a duo, propose, inspect objective/reward on B, decline; verify no start.
   Repropose, accept on B, Start on A; both journal/tracker rows show the named0/6.
3. Kill with A while B only heals nearby: both1/6. With three enrolled nearby,
   another wolf changes the counter by exactly one. A personal wolf quest also
   progresses normally. Check unrelated parties and a non-enrolled late joiner.
4. Disconnect B at4/6; A kills alone: remains4/6. Reconnect B, recreate/regroup the
   party and kill together: both5/6. Restart the server at this point and resume.
5. Complete6/6, return to Alaric using normal turn-in guidance. A claims360×GoldMult
   gold once. Retry/relog A. Restart before B claims; B can claim once independently.
   Verify a never-participating enrolled member cannot claim and can abandon.
6. Check EN/ES, classic/modern UI, hidden tracker/manual unpin, pending proposal
   timeout/cancel, roster changes before Start, full wallet and explicit abandonment.
