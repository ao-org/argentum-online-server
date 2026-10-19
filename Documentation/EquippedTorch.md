# Equipped Torch

The production lit torch (OBJ2675) keeps its resource definition but no longer
occupies `EquippedShieldSlot`/`EquippedShieldObjIndex`. Its existing inventory
`Equipped` flag is authoritative and persists normally; a bounded inventory scan
finds it, so slot swaps require no new cache or database field. Equipping another
torch replaces the previous one. Shield/weapon restrictions remain unchanged.

Existing packet 113 (`eParticleFX`) carries dedicated effect 355. Its unchanged
17-byte little-endian layout is packet Int16, character Int16, effect Int16,
lifetime Int32 (-1), remove Bool, GRH Int32 (0), x UInt8 (0), y UInt8 (0).
This is an idempotent carried-light marker, not an asset particle allocation.
No new packet ID, capability, resource change or protocol layout is needed.

Equip publishes the marker to the visibility-filtered living-player area.
Unequip/removal publishes its removal to the area including dead viewers, so
death clears the owner's light too. Item removal and death use `Desequipar`;
inventory reset also clears the marker. Login restores the inventory flag
without premature broadcasts. After CharacterCreate, each observer receives a
positive/negative snapshot; hidden carriers receive SetInvisible before a
positive marker. Administrative visibility follows existing recipient checks.

HOO suppresses the light while invisible/hidden and resumes it when visible.
It preserves shield appearance, skin and displayed defence. Existing torch art
uses an empty off-hand only; with a shield, the shield remains drawn and the
independent torch remains a light source. Brightness/range are unchanged.

Deploy matching HOO before this server change. No legacy-client changes or
negotiation are provided. Test equip orders, duplicate torches, removal, death,
relog, map/area entry, stealth and reappearance with a local two-client session.

## Native Validation

Windows VB6 was invoked with `/make Server.VBP /d
DEBUGGING=<UNIT_TEST>:PYMMO=<0|1>:UNIT_TEST=<0|1>:UNLOCK_CPU=1:BATTLESERVER=0
/out <build.log> /outdir <isolated-output>`. All four combinations built.
Both test-enabled executables ran in disposable local directories/databases:
438/438 full tests, 33/33 trap tests, 106/106 managed-configuration checks.
Eight new General-suite regressions exercise actual equip/unequip, replacement,
death, hidden snapshot ordering, two-handed restrictions, item removal and login
restoration. The orb capture fixture ignores the separate torch marker without
changing its orb assertions. Packet capture is compiled only with UNIT_TEST=1.

The user's existing Server.VBP reference/configuration edits were preserved.
Live two-client testing has not been performed; no deployment was made.
