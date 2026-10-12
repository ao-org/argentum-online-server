# Correlated taming presentation results

The HOO presentation companion consumes a capability-negotiated acknowledgement
of the **existing** WorkLeftClick request. No request packet or gameplay cooldown
is added. The 20% `RandomNumber(1, 5) = 1` roll, class/skill formula, minimum level,
ownership/capacity checks, pet following, respawn and skill progression remain
unchanged.

`PrepareActiveToggles` always advertises supported wire capability
`hoo-taming-result-v1`. A version-1 client must acknowledge bit `0x20` before
`WriteHooTamingResult` sends anything. This is not an optional gameplay rule.

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
adds pure outcome mapping; the existing capability suite checks the new bit and
protocol-version rejection. The companion HOO tests
verify exact field decoding and stale/cross-map/counter/character rejection.
VB6 compilation and live success/failure/following/save-relog tests require a
development server; source-level checks alone do not constitute that validation.

`Codigo/PacketId.bas` reserves retired magic-orb ID 206 (and capability bit
`0x10`) instead of reinterpreting older clients. It appends taming without
renumbering existing packets. Keep this extension when regenerating the packet catalogue;
the separate ProtocolGenerator working tree is not modified by this change.
