# Authoritative quest feedback

Companion HOO consumes `eQuestFeedback` (225 with PYMMO=1, 226 with PYMMO=0).
Coordinate client/server release; include `OGG/quest_ready.ogg` from the assets
companion in the normal audio pack. No capability flags, negotiation, protocol
generation or version-enforcement changes.

The existing dirty readiness drain supplies its ready collection to
`modQuestFeedback.PublishQuestFeedback`, after network Poll and before Flush.
No second readiness evaluation, per-inventory-slot evaluation, polling, database
lookup or ordinary-quest global scan is added. Required items use the same
`CountInventoryObjects` implementation as gameplay `TieneObjetos`, including
normal elemental-tag matching and wallet gold semantics. Kill/visit counts come
from the accepted slot's actual arrays, including final completed counts.

State snapshots are explicit silent synchronization on first publication or
forced login refresh. Acceptance lifecycles follow quest identity across slot
compaction and change on successful reacceptance. Fingerprints suppress unchanged
packets. Sequences and lifecycle counters belong to the player/session, not the
global population. Existing dirty work validates character ID/connection version,
retains rejected enqueue and uses the established 250-ms retry boundary.

`FinishQuest` captures grants at their real mutation sites: multiplied EXP only
when awarded, wallet/bank gold according to the existing destination branch,
actual item count increases after successful inventory insertion, and spells
only when newly assigned. Grant rules, rates, capacity checks and rewards are
unchanged. Completion/QUESTFIN is published after successful grant processing;
the typed receipt is queued after AddDoneQuest. Rejections and interrupted
operations do not produce a successful receipt. Old success console lines are
replaced by the HOO correlated receipt, avoiding static/duplicated reward claims;
failure/skip explanations and narration remain.

At most 16 receipts remain pending per player; accepted enqueue removes each in
order. Overflow logs a diagnostic without altering gameplay. No pending work is
carried into another connection identity. The client keeps an eight-notice visual
queue and an actual-grant console record, independently of pins/navigation.

## Layout

Int32 positive session-local sequence; UInt8 kind (0 complete state / 1 receipt),
UInt8 silent, UInt8 quest count (0..5; receipt exactly one). Per quest: Int16 ID,
Int32 acceptance lifecycle, String8 title (128 source bytes maximum), UInt8 ready,
Int16 goal count (0..765). Per goal: UInt8 kind (1 kill / 2 visit / 3 item), Int16
ordinal, Int16 template/item, Int32 current, Int32 required. Current is clamped
to required. Tail: Int16 actual reward count (0 for state; 0..512 for receipt),
then UInt8 kind (1 EXP / 2 gold / 3 bank gold / 4 item / 5 spell), Int16 ID
(0 for currency/EXP), Int32 amount. Empty state replaces all accepted entries.

Suite 48 covers native writer bytes, final kill/visit/item counts, initial silence,
unchanged suppression, actual successful/rejected turn-in, banked gold using the
configured multiplier, item grants, max-level EXP omission, ordered independent
receipts, retries, repeat acceptance and lifecycle preservation during compaction.
The earlier readiness batching/global regressions remain enabled.

For manual acceptance use authored quest 41 (Colmillos de Fuego), 25 template-576
kills, Desh template 1107 at map35 (45,78), base rewards 675 EXP/275 gold. Complete
it with a below-cap character and inspect the actual configured multipliers,
return guidance, one ready notification, grants and pin removal. Test an available
required-item quest for readiness loss/regain and repeatable acceptance separately.
No live gameplay acceptance is claimed by the automated harness.
