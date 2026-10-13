# Correlated taming presentation results

The HOO presentation companion consumes a mandatory protocol acknowledgement
of the **existing** WorkLeftClick request. No request packet or gameplay cooldown
is added. The 20% `RandomNumber(1, 5) = 1` roll, class/skill formula, minimum level,
ownership/capacity checks, pet following, respawn and skill progression remain
unchanged.

`WriteHooTamingResult` sends resolved outcomes directly, without capability
checks or feature advertisements. HOO is the supported client; retired clients
and older servers are unsupported. Deploy the matching HOO client and VB6 server
together. ProtocolGenerator is retired and is not part of this workflow.
Older clients may reject packet 207.

New functionality MUST NOT introduce `HOO_CAP_*` flags, per-feature negotiation,
or feature-support advertisements. It belongs to the normal required protocol.
Existing unrelated capability infrastructure remains until a separate audit;
do not extend it or introduce alternate compatibility paths.

Appended server ID **207** contains, in existing little-endian framing:
Int32 original WorkLeftClick counter, Int16 logical map, UInt8 submitted X,
UInt8 submitted Y, Int16 NPC character index, UInt8 outcome (0 success,
1 eligible failed random roll, 2 rejected). Including the Int16 packet ID, the
payload is 13 bytes. No credentials or account identity is included.

The character index is captured before `DoDomar`, since success can immediately
remove a pet on a NoMascotas map. `DoDomar`'s optional presentation output defaults
to rejection; success/failure is assigned only after normal processing completes.
Other callers remain source-compatible. Existing locale messages, including
suppression of repeated identical failures, are untouched. The structured
acknowledgement is not suppressed by `UltimoMensaje`.

Malformed/rate-rejected requests and early WorkLeftClick validation exits may
produce no acknowledgement; clients must silently expire their pending visual.
Missing targets have character index zero and must not produce guessed effects.
A structured result does not replace authoritative character/ownership updates.

`UnitTesting.test_taming_results` retains its five-roll probability regression and
adds pure outcome mapping and fixed packet-ID assertions. The companion HOO tests
verify exact field decoding and stale/cross-map/counter/character rejection.
`tools/test-taming-protocol.ps1` extracts the production functions/writer verbatim,
compiles them with VB6 and uses the actual Aurora writer with a capture-only send
sink. It tests four classic taming regressions plus exact 13-byte serialization,
all three outcomes, missing NPC and unresolved-result suppression. It does not
open sockets, start the server, load resources or touch a database.

Native Windows validation (2026-10-13): all five harness checks passed. Full
server builds with `PYMMO=1`, `UNIT_TEST=0` and `UNIT_TEST=1` passed. The full
server executable was not started; live success/failure/following/save-relog
remain manual validation, not claims of the isolated harness.

`Codigo/PacketId.bas` reserves ID 206 and assigns taming 207 in the supported
HOO `PYMMO=1` build, without renumbering earlier packets. `PYMMO=0` retains its
existing extra account packet and shifted tail; it is not HOO's wire layout.
The authoritative server enum and HOO decoder are verified directly, including
the exact native-writer byte fixture. No generator regeneration is required.
The client keeps the attempt visual at two seconds and correlation at a bounded
30 seconds, checks map/revision/generation/counter/tile, and never infers ownership
from particles or sound. NoMascotas despawn suppresses the visual, not the matching
acknowledgement; newer attempts on the same creature supersede earlier ones.
