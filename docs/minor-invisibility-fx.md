# Minor Invisibility activation FX

Audited server base `3d11b70d95f0a108d6dd2c1ce5a0d160a1ee2140`, assets master
`29156e971e3a5a1cf0872a3540223d630cf91888`, HOO base
`2a04c09d4380b4550d4ff6f953a088fb462fd012`.

Spells 24 and 30 are Minor Invisibility. Both have Invisibility=1, Duration=15,
WAV=16 and no Particle/FXgrh. Mana costs differ (400/275). Existing successful
application randomizes the duration to 11–16 seconds. No gameplay values change.
Other invisibility spells and many unrelated spells share sound 16; it is not
an activation identifier.

After all existing invisibility rejection checks, a previously visible target
receives the new presentation marker, then the ordinary visibility update.
The marker uses existing eCreateFX: Int16 packet ID, Int16 character index,
Int16 FX=30001, Int16 loops=0, Byte x, Byte y (10 bytes total).
30001 is a reserved semantic effect ID, not an asset/GRH index. No new packet,
capability negotiation, ProtocolGenerator, data or asset changes.

Only successful spell 24/30 applications emit it. Already invisible, hidden or
admin-invisible targets emit no marker. Existing ToPCAliveArea routing preserves
area, instance and dead-viewer/admin rules. This is the existing same-map FX
recipient scope; it does not expand spell presentation to adjacent-map viewers.
The client captures the visible/interpolated anchor before SetInvisible, applies
visibility immediately, then animates at that fixed source-map location.
Sound, counters, spell feedback, costs and duration are unchanged.

Deploy the companion HOO first, then this server. New client + old server retains
original presentation. Old HOO does not implement the reserved effect; coordinated
upgrade is required for the new visuals. Disabled enhanced FX ignores the marker
and preserves original sound/visibility (there was no legacy graphic to replace).

Validation: six native unit regressions added to suite General: spell identity,
hidden-target exclusion, unrelated/shared-sound exclusions, and the existing wire
layout. Native VB6 compiler is unavailable locally; these are pending the existing
PR validator. Source encoding remains Windows-1252/CRLF. `git diff --check` passed.
Live multi-client visibility and successful/failed casts remain unverified.
