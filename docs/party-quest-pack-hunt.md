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

The unshipped draft migration `20261103-01-party-quest.sql` and base schema add
run revisions, fixed roster ordinals and bounded kill masks. Enrollment, kills,
claims and abandonment mutate one cached run and increment its dirty revision.
**Kills perform no SQL**; both clients receive the new counter immediately. The
journal is a projection, rebuilt from the shared run on login. Unexpected crashes
may roll back changes since the last successfully committed character save.

`SaveCharacterDB` (logout/full save) and `SaveChangesInUser` (periodic save) capture
a consistent run and whole-character batch for party-quest characters. Existing
main, inventory, bank, spells, skills, pets, skins and quest serializers build
immutable commands against a working copy, without clearing live dirty flags.
Account card changes use the same transaction. Unrelated characters retain their
existing persistence path. A save is successful only after `CommitTrans` returns;
ADO dispatch is not completion. Connection event fences wait for pre-existing
pooled writes to finish before the party transaction starts (without a busy loop).
See [ADO ExecuteComplete](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/executecomplete-event-ado).

Run revisions prevent older snapshots overwriting newer progress. A monotonic
`user.party_quest_save_revision` guards the whole captured character batch, rather
than adjusting or replacing gold. Only committed matching revisions acknowledge
saved baselines; later mutations stay dirty. Two members share one run identity.
New IDs are reserved from the database maximum once during initialization, not on
a kill. This assumes the existing single world-server writer for this database.

Claim validates the existing requirements and grants gold once in memory, marking
the member claimed and the run dirty. A save containing an unsaved claim includes
that claimant's captured character batch, even when another member initiates the
save. Wallet, ordinary character changes and claim status therefore commit or
roll back together. Subsequent earnings/spending remain in the next snapshot.
There is no save-blocked flag, wallet recovery assignment or claim watermark.
Full wallet capacity leaves the claim available for retry.

Failed/deferred saves retain immutable batches under persistent character identity,
independent of connection slots. The existing autosave cadence retries a bounded
number, coalescing to the latest run revision and online character state. Offline
snapshots survive logout; reconnect flushes retained work before reading the main
character row. Orderly shutdown retries remaining batches before closing the DB
and refuses to finish while a save remains pending. No new timer, database polling
loop, protocol field or general persistence framework is added.

Loading validates the complete run, roster and kill history in temporary state
before publishing to the cache; read failure allows a clean retry. Historical
`character_id` and ordinals are immutable. A separate nullable `user_id` foreign
key uses `ON DELETE SET NULL`, so character deletion preserves other members'
progress, participation masks and claim rights. Deleted active members become
abandoned tombstones on load. Dirty runs cannot be evicted. Storage is bounded by
user capacity; a full dirty cache rejects enrollment rather than losing progress.

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
