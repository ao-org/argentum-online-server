# World click target priority

`LookatTile` and `HandleLeftClick` use `ResolveWorldClickCharacter` to resolve
one consistent character target. The requested tile takes precedence. The tile
immediately south is considered only if the requested tile has no character,
preserving the legacy head-click fallback on empty tiles. On either chosen tile,
an NPC takes precedence over a player occupying that same tile, as in the
existing `LookatTile` behavior.

Previously `LookatTile` preferred the south tile, even when the exact requested
tile contained the intended NPC. A pet/elemental there could steal a manual spell
target. `HandleLeftClick` independently searched for players on both tiles,
allowing a player south of a merchant, banker or quest NPC to block interaction.
The shared resolver removes that disagreement. Interactable NPC actions are
forwarded to the resolved NPC's tile; direct player clicks do not open a nearby
NPC window. Hostile NPCs still do not trigger world actions.

The matching HOO change resolves the current NPC sprite under the pointer and
sends its actual tile. Deploy both changes. Packet layouts, timestamp validation,
spell eligibility, ownership, distance checks, NPC AI and combat rules are unchanged.
Multiple ambiguous NPCs on one client tile remain excluded from automatic targeting.

## Validation

`Unit_WorldClick` is registered in the existing UNIT_TEST suite. Its ten cases
cover direct NPC/player priority, pets or players below, shared tiles, empty
targets, and head-click fallback. `tools/test-world-click.ps1` extracts the actual
production helper and suite into an isolated native VB6 executable. Pass `-Vb6`
if the compiler is installed elsewhere.

Without VB6, `-ScriptHost` runs the same pure selection statements and test cases
through Windows VBScript, removing VB6 types and error handlers. This checks logic
only and does not replace native compilation or full-server validation.

Manual checks still required: merchant/bank/quest windows with players to the
south; offensive spell targeting and repetition with an elemental/pet to the
south; direct player clicks; empty-tile head clicks; bottom map edge; pet orders;
classic and modern HOO interfaces. Do not treat these checks as live signoff
until actually performed.
