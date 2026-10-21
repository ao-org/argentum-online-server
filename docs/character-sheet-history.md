# Character history snapshot

`WriteMiniStats` now includes three authoritative values requested by the modern
Character panel: `Stats.JineteLevel` (Byte), `Faccion.FactionScore` (Long), and
`Hogar` (e_City, encoded as Int16). No stats, faction scores, homes, or gameplay
rules are changed. JineteLevel is not the old Riding skill value.

The existing request is client packet 87. Server packet 79 retains its original
fields and appends UInt8 riding level, signed Int32 faction points, and signed
Int16 home city. The resulting payload is 37 bytes including the packet ID, or
39 bytes including the two-byte frame length. Signed faction penalties are
preserved; the client displays unknown home IDs as Unknown.

Coordinate deployment with the client branch `fix/character-history-riding-faction-home`.
The updated decoder requires these fields. Old clients cannot decode the longer
response; updated clients cannot decode the older response. Deploy the matching
client and server together. No new packet ID, HOO_CAP flag, feature negotiation,
or ProtocolGenerator modification is involved.

The client requests this snapshot through C and displays the new fields on
History & Status → History (2/2). It does not implement or send `/EST` and does
not parse localized console messages. `SendUserStatsTxt` remains unchanged.

Validation: the client tests assert exact wire order, all truncated lengths,
trailing-byte rejection, signed score boundaries, all ten city IDs, localization,
model refresh/reset and page navigation. A native server build and live paired
client/server test must also be performed before deployment; neither is implied
by the client unit tests.
