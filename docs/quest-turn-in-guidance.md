# Quest turn-in guidance

Packet `eQuestTurnIns` follows `eQuestObjectiveTemplates` (224 with PYMMO=1,
225 with PYMMO=0). Payload: UInt8 ready-quest count (0..5), then per quest:
Int16 quest ID, String8 title, UInt8 recipient count (0..128), then each
recipient's Int16 template ID and String8 name. Names/titles are limited to
128 source characters. This is a complete replacement, including an empty
snapshot. It contains no live NPC coordinates. More than 128 recipients keeps
readiness but publishes no partial recipient set.

Readiness delegates to `CanInspectQuestTurnIn`: the existing nonmutating
`FinishQuestCheck`, required class, free reward slots and reward-spell admission.
Actual turn-in continues to run the normal server checks and grants no rewards
from this snapshot. Recipients follow `EnviarQuest`: `TalkTo` and authored NPC
templates whose offered quest list contains the quest. Metadata is indexed once
per quest and invalidated on NPC/quest reload.

Updates follow login/quest snapshot publication, inventory-slot and inventory
capacity updates, spell/skill updates, natural skill gains, equipment skill
effects, and global event changes. Each per-user check examines at most the five
accepted slots. Global changes inspect only accepted slots of connected users;
no per-frame polling or database reads are added. Cache identity includes user ID
and connection version; only accepted enqueue updates the cache.

Release the matching HOO client and server together. There are no capability
flags, negotiation, compatibility fallbacks, or version-enforcement changes.

Suite 48 (`Unit_QuestTurnIn`) covers actual writer bytes, different recipients,
required-item gain/loss through inventory writes, authoritative removal, enqueue
failure, lost requirements and reward admission. Native VB6 compilation/tests
could not be run locally because the documented compiler is absent. Both
PYMMO configurations and the full harness require the PR validator.

Manual acceptance: complete an accepted quest, verify the tracker return row,
follow a route back to its recipient, inspect the checked scroll and hover text,
then collect the reward. Remove a required item before turn-in and verify readiness
disappears. Check rejected turn-in, abandonment, repeatable acceptance, hidden
tracker, manual unpin/reconnect and character switching with both HOO UIs.
