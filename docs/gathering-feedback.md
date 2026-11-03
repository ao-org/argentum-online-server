# Authoritative gathering feedback

HOO companion uses two new required packets (PYMMO=1): client 318
`eGatheringStart`, server 222 `eGatheringFeedback`. No capability bit and no
ProtocolGenerator. Deploy server first, then the coordinated HOO client before
allowing the new feature; an old server cannot process request 318. Existing
packet layouts remain unchanged. PYMMO=0 enum ordinals are 321 and 223.

All integers are little endian. Start payload: positive Int32 request, Byte x,
Byte y, Byte skill (17 lumberjacking, 18 fishing, 19 mining). It enters the same
validation/start handler as the old automatic request. Smelting retains the old
request. The result is 24 bytes including the Int16 packet ID:

| Field | Type |
|---|---|
| Request, sequence | Int32 each, positive |
| Source map | Int16 |
| Skill, outcome, reason, active | Byte each |
| Awarded object | Int16 |
| Awarded quantity | Int32 |
| Equipped working tool | Int16 |

Outcomes: 0 accepted start, 1 inventory award, 2 valid unsuccessful roll,
3 blocked, 4 stopped, 5 award dropped on ground. Reasons: 0 unspecified,
1 stamina, 2 inventory full, 3 invalid resource target, 4 missing/unsuitable tool,
5 distance, 6 fishing location/tool restriction, 7 depleted, 8 special-fish
challenge, 9 no floor space. Unspecified does not imply success or a guessed failure reason.

Only real extraction branches emit awards. Wood/ore overflow still falls on the
floor; fishing's inventory-full branch still stops without granting the catch.
Bonus fish items and completed special-fish challenges report their actual award.
Special catches retain their original request through a later activity change.
Only the existing gathering award call is observed synchronously before/after insertion, so stack limits cannot inflate the reported quantity. No inventory-sync observer, rates, rolls, costs, allocation or retry changes.
Legacy manual WorkLeftClick still follows its existing handlers (which do not
implement repeated extraction); no activity is inferred from sending Work.

The per-user context includes ID/version. ResetUserAutomatedActions emits one
stop before clearing gameplay state. A blocked result suppresses a redundant
stop. A stopped context cannot become active again, but can report a legitimate
late award. A rejected replacement leaves the old server action unchanged; HOO retains its confirmed status. Accepted replacement stops only the old feedback context before starting the new one.
A bounded client request history and monotonically increasing sequence reject
duplicates and stale status; they cannot infer changes from console messages.

Failed chopping retains sound 64 for other nearby viewers. For a correlated HOO
request, its local copy is played only from outcome 2 on the Effects bus with a
single label and reduced gain. Success sound 13 and other audio are unchanged.
No assets are added. Auditory tuning remains a live-client check.

## Validation

Suite 46 contains packet bytes/IDs, accepted start, exact rewards, unsuccessful
roll continuation, block/stop deduplication, late award, user lifetime reuse,
legacy request exclusion, rejected replacement, real stamina gate, tool
validators, ground award and zero-yield cases. Both PYMMO enum configurations
are asserted. Native VB6 compiler is unavailable at the documented local path;
PR validator results must be reported separately from local checks.

Manual: exercise all three jobs, full inventory (floor vs no award), depleted
resources, movement/death/tool cancellation, special fish and its bonus items,
failed chop audio/mute, rapid restart, map/disconnect and late queued results.
