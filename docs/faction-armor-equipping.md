# Rejected faction armor replacements

`EquiparInvItem` must validate a new faction armor before calling `Desequipar`
on the currently equipped armor. Previously, rejection for incomplete stats
happened after removing the old armor, its body aura and any dependent skin.
The fix moves the existing removal block after the unchanged validation block.

The existing check covers health, mana, hunger and thirst for Royal/Chaos armor
outside safe maps and safe-zone triggers. It does not add a stamina requirement.
Login restoration and intentional unequipping keep their existing behavior.
The existing localized rejection message and all protocol packets are unchanged.
No HOO, legacy-client, resource, balance or database changes are needed.

The General native test suite adds 20 cases (Royal and Chaos variants): each
of the four incomplete stats preserves the old equipped slot, body, aura and
armor-dependent skin, and emits only the existing rejection message. Other
cases check full stats, safe maps, safe triggers, login restoration, ordinary
armor and deliberate unequipping. Fixtures restore the original user, object
metadata and map state; packet capture is compiled only with `UNIT_TEST=1`.

Validation in this workspace: byte comparison confirms that the existing
unequip block was only moved, the eligibility predicate/message are unchanged,
and intentional unequip still precedes the gate. `git diff --check` passed.
The documented VB6 compiler is unavailable here, so native compilation and the
20 new behavioral tests have not been run.

Manual check on a disposable test character: equip ordinary clothing with an
armor-dependent skin/aura, lower one stat, then attempt faction armor outside
a safe area. The message should appear while the original clothing, appearance
and equipment remain unchanged. Repeat with each checked stat and both factions;
then verify a valid replacement and deliberate unequipping. A server rebuild
and deployment are needed for the gameplay fix; neither was performed here.
