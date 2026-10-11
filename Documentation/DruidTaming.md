# Druid taming: rejection audit and regression procedure

Audited server base: `16546f49f55c72a298e81ce8b2a2a0b15b00d85c`
(`ao-org/argentum-online-server`, current master fetched for this change).
Source locations below refer to that base unless explicitly marked as changed.
No particular production character/NPC or captured request was supplied, so this
does not identify which restriction affected that individual player. The code
does establish two bugs without assuming a different gameplay balance.

## Confirmed causes and corrections

1. **Impossible power check reported as random failure.**
   [Trabajo.bas:1749-1790](https://github.com/ao-org/argentum-online-server/blob/16546f49f55c72a298e81ce8b2a2a0b15b00d85c/Codigo/Trabajo.bas#L1749)
   combines `Domable <= puntosDomar And RandomNumber(1, 5) = 1`. Both a power
   rejection and an eligible losing roll enter the same message-659 branch.
   The correction classifies them separately: existing localized 718 means
   cannot tame at current power; 659 means an eligible attempt failed its roll.
   No new message ID, translation asset or packet is necessary. Message 718
   intentionally remains shared with untameable NPCs; it does not disclose a
   new numeric power requirement to the client.
2. **NPC ownership check tested a Boolean as a positive number.**
   [Trabajo.bas:1744](https://github.com/ao-org/argentum-online-server/blob/16546f49f55c72a298e81ce8b2a2a0b15b00d85c/Codigo/Trabajo.bas#L1744)
   uses `IsValidNpcRef(...MaestroNPC) > 0`. VB6 Boolean True is -1, so a valid
   NPC owner fails this test. Removing `> 0` restores the intended existing
   ownership rejection (656). User ownership checks remain unchanged.

Changed implementation: [Trabajo.bas](../Codigo/Trabajo.bas),
`GetTamingPower`, `GetTamingAttemptMessage`, and `DoDomar`.
The power helper preserves the original calculation, including its Integer
intermediate. The classification helper permits deterministic testing without
changing the random generator or adding a client-facing result protocol.

## Complete request path and restrictions

| Stage | Existing behavior / location |
| --- | --- |
| Dispatch | `Protocol.HandleIncomingData`, dispatch cases at 369/381; existing session/packet validation applies before work handlers. The PYMMO logged-user gate is at 271-308. |
| Request target | `Protocol.HandleWork`, 2712-2728: reads skill and counter, rejects dead users, cancels logout, sends the existing WorkRequestTarget for Domar. The hiding-specific counter check does not apply to Domar here. |
| Target packet | `Protocol.HandleWorkLeftClick`, 2942-2975: reads x/y/skill/counter; validates the WorkLeftClick timestamp/rate policy; records work target; rejects dead, resting, out-of-map, stunned and out-of-vision attempts; stops meditation and cancels logout. |
| Resolve target | `LookatTile` (`GameLogic.bas:1075`, NPC selection at 1254) establishes the target reference from the clicked tile. `IsValidNpcRef` checks bounds/version (`MODULO_NPCs.bas:59`). No valid NPC gives 719. |
| NPC eligibility | `Protocol.bas:3281-3304`: `Domable <= 0` gives 718 before distance/combat/DoDomar. Missing difficulty defaults to zero and is not tameable. |
| Range | Same branch: Manhattan distance `Abs(dx) + Abs(dy) <= 4`, not a square/circular radius. Too far gives 8. General visibility validation also applies. |
| Combat claim | `GetOwnedBy` (`MODULO_NPCs.bas:2480`) must return zero. It checks a recent AttackedBy name, elapsed claim time and a valid user reference, including the requesting player. A live claim gives 717. This is not simply the NPC's hostile flag. |
| Already yours | `DoDomar:1736-1740`: valid MaestroUser equal to requester gives 655, before pet capacity checks. |
| Staff restriction | `DoDomar:1742`: Consejero privilege exits silently. No new class or staff permissions are introduced. |
| Capacity | `DoDomar:1743` requires both NroMascotas < MAXMASCOTAS (3) and a free MascotasType slot; otherwise 660. |
| Existing owner | `DoDomar:1744-1747`: valid MaestroNPC or MaestroUser gives 656. This is the Boolean bug corrected above. |
| Power and level | Calculates class-specific power; MinTameLevel greater than character level gives 1321 with the required level. No roll/skill attempt follows a level rejection. |
| Attempt | Compare positive difficulty to power and draw 1..5. Exactly roll 1 succeeds for eligible power. Below-power attempts still draw once and still reach SubirSkill. |
| Result | Success binds the pet, follows the owner, performs the existing respawn handling and sends 657. A NoMascotas map stores the pet for later and additionally sends 658. Failed eligible rolls send 659; insufficient power now sends 718. |

There is no extra taming-specific mounted/invisibility, faction, class whitelist
or two-of-one-species guard in this current path. `PuedeDomarMascota` describes a
same-type limit but has no caller. `ModDomar` is also unused. Wiring either into
the active path would change balance and is deliberately out of scope.

## NPC configuration, math and progression

`MODULO_NPCs.LoadNpcInfoIntoCache` loads `Domable` at 1077 and `MinTameLevel`
at 1119 (default 1); `OpenNPC` copies them to the live NPC at 1463/1494.
Configuration/cache difficulty is Integer; live MinTameLevel is Byte
(`Declares.bas:3534`) while the cache uses Integer. Out-of-range configuration is
not repaired by this change. `FileIO.BackUPnPc/CargarNpcBackUp` also read/write
Domable but have no callers in this revision; they are not the normal loader.

Read-only data cross-check: `ao-org/Recursos` at
`5888d173d62f0752cbc32f939dd7a043380ccc87`, committed `Dat/npcs.dat`,
contains Gallina NPC501 with difficulty 17, Wolog NPC527 with 283, and Pantera
Negra NPC1141 with 283 and minimum level 31. These examples are not assertions
about deployed production data. No resource file was edited.
Existing Spanish/English messages were checked in committed
`init/SP_LocalMsg.dat` and `init/EN_LocalMsg.dat` at the same resource revision.
The server uses the unchanged `WriteLocaleMsg` path (`Protocol_Writes.bas:899`)
and constants in `modMessageIDs.bas`.

Active formulas, deliberately unchanged:

```text
Integer product = CInt(charisma) * CInt(taming skill)
Druid:     Integer power = product / 6
All other classes:       = product / 118
```

Assignment rounds to nearest with ties to even. Examples: 21*1/6 = 3.5 becomes
4, whereas 15*1/6 = 2.5 becomes 2. Do not replace `/` with integer division or
compare difficulty to the unrounded quotient. Druid charisma 21/skill 80 gives
280, and skill 81 gives 284; difficulty 283 crosses eligibility there. The unused
ModDomar's Hunter/Cleric divisors are not the current gameplay formulas.

Attributes and skills are Byte (`Declares.bas:2548-2550`), but the supported skill
cap is 100. The largest supported product is 255*100 = 25500, safely below 32767.
Corrupt skill values such as 255 can overflow the existing Integer product.
That behavior is explicitly regression-tested and not silently widened/rebalanced.

VB6 `And` evaluates both operands: the old code consumed a random draw even when
power was insufficient. The new code keeps that draw so it does not inadvertently
shift the following progression roll. `SubirSkill` still runs once after success,
eligible failure **and insufficient power**; it does not run after the earlier
target/range/combat/owner/capacity/level rejections. Its existing cap, level,
hunger/thirst and progression RNG checks remain in `Modulo_UsUaRiOs.bas:1655-1718`.
No skill-rate change is included.

## Pet state and repeated messages

`FreeMascotaIndex` checks MascotasType, not just live NPC references: stored pets
still occupy slots. The separate count check prevents exceeding three; if count
and slots disagree the existing checks can reject conservatively. No speculative
pet reconciliation or storage migration is introduced.

Success still increments NroMascotas, sets a versioned MascotasIndex reference,
stores the NPC type, binds MaestroUser, marks ModificoMascotas and calls FollowAmo
and ReSpawnNpc. `FollowAmo` clears combat targets and changes AI; the NoMascotas
branch removes the live pet but restores its type/count for storage.
`CharacterPersistence.SetupUserPets:442` restores types/count; save logic at
876-912 and 1260 onward preserves stored pet types and excludes timed summons.
`MODULO_NPCs.QuitarMascota:1874` clears the corresponding live/type slot and count.
None of these operations or persistence formats were changed.

Repeated 659 suppression was not permanent: `General.bas:1473-1478` clears
UltimoMensaje after three message-timer updates. Macro stop/reset paths also
clear it (`Protocol_Writes.bas:253`, `Protocol.bas:6988`). Keep that policy.
The correction suppresses repeated identical failure *reasons*, so a transition
from insufficient power (718) to an eligible losing roll (659) is visible.
Success records its result instead of leaving a stale 659 that could suppress
the next creature's failed attempt. No new cooldown/timer is added.

## Automated validation

`UnitTesting` suite 41 registers four pure tests plus a DoDomar ownership
regression. The ownership fixture uses an NPC-owned, deliberately impossible
difficulty and maxed skill, so a regressed guard changes the result without
creating a pet or advancing a skill; modified fixture fields are restored.

For reproducible native, database-free checks, run from Windows PowerShell:

```powershell
.\tools\test-taming-vb6.ps1
```

The script extracts the exact production helpers, enum/message constants and
existing pure test bodies, compiles them with installed VB6 and runs an isolated
temporary executable. It does not translate/reimplement the formulas or start
the server. Four tests cover class formulas, half rounding, all 25,856 supported
charisma/skill pairs for both divisors, equality boundaries, all five rolls,
untameable difficulty and preserved overflow for corrupt skill values.
The DoDomar ownership fixture needs the initialized server test environment and
is **not** executed by this pure runner.

Full native compilation commands (run in the repository directory):

```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make Server.VBP /out taming-build.log
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make Server.VBP /d UNIT_TEST=1 /out taming-unit-build.log
```

The second build includes suite 41. Running that server build executes the
existing full runner after startup, writes `test_results.txt`, then shuts down.
It must use an isolated development database/resources and a free loopback port:
`General.bas:666` starts listening before the test runner at 679. Do not run it
against production configuration merely to obtain a test report.

## Manual integration matrix (development server only)

Prepare a level-40 Druid with charisma 18, taming skill 100 (power 300), no pets,
and a fresh unowned NPC fixture with Domable=300, MinTameLevel=31. Use a disposable
test database/resource copy and restart or invalidate the NPC cache after fixture
changes. Keep the production data untouched. Record numeric locale IDs as well
as rendered text. For deterministic success/failure inspection, pause in the VB6
IDE immediately after the ResultMessage calculation and inspect the roll/result;
pure tests exhaust all five rolls without modifying production RNG.

| Case | Expected result and invariant |
| --- | --- |
| Eligible success / failure | Repeated real attempts eventually show 657 or 659. For deterministic branch debugging choose the result at the breakpoint in the test IDE only. Success binds one pet; losing rolls never bind one. Skill progression still uses existing SubirSkill rules. |
| Insufficient power | Difficulty 301 with power 300 always gives 718, never 657/659. One attempt still reaches SubirSkill. To observe advancement use a non-maxed skill, sufficient level/food/water, and a difficulty above its calculated power. |
| Level | Level 30 / MinTameLevel 31 gives 1321; equality passes this gate. No roll or progression on rejection. |
| Untameable | Domable 0, missing, or negative gives 718 before DoDomar; no progression. Do not make these creatures tameable. |
| Ownership | Your own pet gives 655. Another user's pet gives 656 with spare capacity. A valid MaestroNPC must now give 656; a stale version is not a valid owner. No ownership/reference/count mutation. |
| Range / target | Distances 4/5 pass/fail the taming range respectively; diagonal (2,2) passes, (3,2) fails. Empty/stale targets give 719 or fail existing target validation. Out-of-map/vision packets do not reach DoDomar. |
| Combat | A recent valid attack claim, including your own, gives 717. Retry after expiry; stale/nonexistent attacker references do not hold a valid claim. |
| Pet limit | Three pets give 660. Three stored MascotasType slots with no live references still give 660. Free count with no free slot also rejects. No out-of-range slot write. |
| Other gates | Dead, resting, stunned and Consejero restrictions remain as described above. No new non-Druid prohibition: its existing /118 formula applies. |
| Repeated attempts | Within the suppression interval, duplicate 659/718 is suppressed. Switch from power rejection to eligible losing roll: 659 must appear. After success on one NPC, a losing roll on another must not be hidden by the old failure flag. Existing timer expiry restores feedback. |
| Storage / persistence | Tame on a NoMascotas map: 657 then 658; type/count retained without a live pet. Save/relog and leave the map to verify restoration. Normal pet logout/relog retains type and ownership; timed summons retain their existing persistence policy. |

At startup check for NPC load errors, invalid MinTameLevel or overflow logs. During
tests check server `DoDomar`/packet errors and duplicate pets. Use the existing
legacy client's taming action; no new packet/client is required. If reverting,
revert this code change and rebuild: no database/resource migration or client
rollback is needed. Manual network/persistence scenarios remain distinct from
the isolated arithmetic tests.

## Results from this change

- Native VB6 compilation of baseline master passed before editing.
- Final normal and `UNIT_TEST=1` full server builds passed with VB6 `/make`;
  `server.exe` was checked after each (3,416,142 and 3,633,230 bytes respectively).
  The unit executable contains all five new suite-41 test names.
- `tools/test-taming-vb6.ps1` compiled and ran its isolated native executable:
  **Total=4, Failed=0**. This validates actual VB6 rounding/overflow semantics,
  not a Python approximation. No server, database or resource startup occurs.
- The fifth, initialized-server ownership regression was compiled but not run.
  The full server test suite, network gameplay and save/relog matrix were not run.
- `git diff --check` passed. Existing `.bas` Windows-1252 encoding and CRLF line
  endings were preserved. No project references or COM registrations changed.
- No HOO/client, resource, NPC balance, database or protocol file was modified.
