# Cooperative party quest rewards

## Recipient snapshot

`CollectPartyRewardRecipients` walks only the credited player's leader roster
(maximum six, including the leader), validates versioned references, logged-in
connection/alive state, matching party ID and leader identity, explicit map equality
and `Distancia(credited.pos, member.pos) < 20` (Manhattan distance: 19 yes,20 no).
The leader need not be nearby. Credited player comes first, duplicates are removed.
Invalid party metadata retains valid solo handling. Shared quest rewards additionally
require the saved defeated NPC map. No contribution ledger is used: support qualifies.

`MuereNpc` continues ordinary kill attribution, callbacks, loot/gold/guild rewards
and respawns. Only quest kill objectives and explicitly configured quest drops are
shared; rewards read `MiNPC`, not the removed/reusable slot. Existing pet-owner
attribution in `DoDamageOrHeal` reaches the same death handler once. Training and
player-owned victims and no valid credited player do not receive shared rewards.

## XP accounting

Damage XP remains incremental. Party damage-derived base XP keeps existing CLng
nearest-even rounding, damage cap and remaining pool cap. Eligibility resolves
before any debit. Each award divides integer base with quotient/remainder: credited
player then roster order receive at most one extra base unit, so shares sum exactly
to the single pool debit. The absent-leader phantom divisor/message is removed.
The saved death remainder uses the same party allocation and modifiers once.

ExpMult, per-recipient level penalty, card bonus and leadership bonus remain in
that order; existing Long rounding after party multipliers/penalties and final
stats assignment remains distinct from base allocation. Max-level recipients keep
existing participation/no-XP behavior; XP is capped. Solo damage and solo unmodified
death remainder retain their historical calculation/presentation. Ordinary gold,
loot, guild XP and scenario rewards are not multiplied.

## Quest state, inventory and HOO

Each changed kill objective caps at its requirement, marks normal quest persistence
dirty and publishes existing objective/inspection state once per member. Already
completed/unrelated objectives do not publish or repeat completion notices.
Each configured quest-drop entry independently checks accepted quest, current
missing amount and normal inventory capacity. Grants cap at remaining requirements
and a usable stack; overlapping entries recheck inventory. Full inventory uses
existing failure feedback. Success audio requires a positive actual inventory delta.

Normal inventory writes and objective publication dirty the existing turn-in queue;
it evaluates final kill-plus-item state at the normal boundary, without forced
intermediate snapshots. No packet, client, database schema, quest content, capability
or protocol-generation changes.

## Validation

Native VB6 compiler is absent locally. New suite50 contains41 deterministic cases;
configured drop probability1 guarantees rolls without modifying production RNG.
Native PR-validator results will be recorded after execution. Source checks preserve
Windows-1252/CRLF and run `git diff --check`.

Persistence tests use existing dirty detection and saved quest/inventory projections;
a real database save/reload and live HOO gameplay are separate manual checks.

## Two-client acceptance checklist

1. Two same-party characters accept the same kill quest. Kill near each other;
   both get one progress update. Let a support-only member deal no damage.
2. Repeat at distance19/20 and across maps. Put the leader on another map while
   nearby members kill; verify their XP is divided only among eligible members.
3. Hunt a configured quest item: each member gets an independent roll, capped at
   personal need. Fill one inventory; that member gets failure/no success sound,
   the other can still receive items. Free space and try a later kill.
4. Let a pet land the finishing blow; owner and eligible party progress once.
5. Finish kill+item requirements and observe HOO tracker/objective markers and
   ready-to-turn-in feedback without reopening the journal. No duplicate sounds.
6. Relog both clients and verify quest counts and actually granted items persist.
   Turn in normally. Test solo, maximum level, normal loot/gold and both UI modes.

Do not merge or deploy automatically.
