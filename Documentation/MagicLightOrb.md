# Following Orb of Light (spell 352)

Deploy the matching HOO client and resource update before this server. No new
packet IDs, capability flags, database fields or objects are introduced. Restart
the server during deployment; live floor-orb instances are not migrated.

Spell data now uses Target=1 and AutoLanzar=1. Tipo=13, Particle=354, duration 60,
mana 20, stamina 5, cooldown 2 and bindability remain unchanged. Mage, Druid and
Bard restrictions and scroll learning remain authoritative. Obsolete floor-light
colour/range/cast-range fields are removed. No terrain target is requested.

The existing UseSpellSlot -> PuedeLanzar -> user-spell handler validates and
consumes costs once; successful casting starts the normal cooldown. The raw-tick
sentinel accepts negative timestamps so cooldowns work beyond 24.9 days uptime.

One entry per user slot holds a versioned owner reference and original deadline.
Recasting refreshes it without a removal packet. Movement, map changes and death
preserve the remaining lifetime. Disconnect/expiry remove it; recycled owner
references are discarded without touching the replacement occupant.

Existing eParticleFX (113), 17 bytes, little-endian:

| Offset | Type | Meaning |
| --- | --- | --- |
| 0 | Int16 | 113 |
| 2 | Int16 | caster character index |
| 4 | Int16 | 354, dedicated orb identity |
| 6 | Int32 | -1 upsert, 0 removal |
| 10 | Boolean byte | false upsert, true removal |
| 11 | Int32 | GRH override 0 |
| 15 | UInt8 | X 0 |
| 16 | UInt8 | Y 0 |

Upserts use the normal living-character area visibility rules. Snapshots are sent
after CharacterCreate in MakeUserChar, including the caster after map changes.
Repeated snapshots are idempotent in HOO and do not refresh server time. Admin
invisibility remains filtered; HOO suppresses the orb and light when invisible.
An invisible caster's snapshot sends SetInvisible before ParticleFX, even though
NotifyUser also sends invisibility later, so split reads cannot reveal the light.
Removal uses the ordinary area route so an observer who died after seeing the
orb can still remove it. Only particle 354 is removed, never all character FX.
The spell no longer emits floor packets 114/121 or changes MapData lights.

Native test builds can run `server.exe --test-light-orbs` through the existing
UnitTesting framework. The same suite remains registered in the full server
suite. Tests intercept outbound packets only under UNIT_TEST=1, DIRECT_PLAY=0.
The ordinary suite uses a saved/restored spell fixture and does not require an
unmerged resources branch. Before coordinated deployment, run the full test
server as `server.exe --validate-light-orb-resources` with the matching Dat path.
This additionally validates the loaded spell's self-cast target, type, duration,
costs, cooldown, particle and bindability. Old ground-targeted data must fail
this explicit integration check; passing unit tests alone does not validate a
deployment. Do not deploy this server with the old resource definition.
See HOO's `docs/magic-light-orb-spell.md` for visual design and manual acceptance.

Validation (2026-10-14): native Windows VB6 builds passed with DIRECT_PLAY=0,
PYMMO=0/1 and UNIT_TEST=0/1. Focused runs passed 25/25 in both test builds;
the full isolated server suite passed 421/421 in both PYMMO configurations,
including the loaded spell-data contract. No live client/server gameplay test
or production deployment was performed.

Review validation (2026-10-15): all four native build configurations passed
again. Focused suites passed 25/25, ordinary full suites 420/420, and explicit
`--validate-light-orb-resources` full suites 421/421 in both PYMMO modes. A
negative run with AutoLanzar=0 / Target=4 failed only the resource contract,
as intended. The former unconditional assertion reproduced a 420/421 result
with assets master, even though the fixture-based gameplay tests passed.
CI report publication: startup/suite checkpoints are written to
`test_progress.txt`, never `test_results.txt`. The validator polls for the final
report's existence, so publishing an incomplete checkpoint at that path caused
it to reject and terminate a still-running server. Final reports are written to
a same-directory temporary file, closed, then renamed into place. A terminal
startup error still publishes an explicit failure. Checkpoints and this report
handling are compiled only into UNIT_TEST builds; gameplay is unchanged.
Test-build spell/object/EOT/map loading and migration failures propagate to that
startup report instead of waiting for an unattended modal dialog. Production
error dialogs are unchanged.
