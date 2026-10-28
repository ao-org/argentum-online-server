# Guild menu experience overflow

Guild experience and level requirements are VB6 Long values. Narrowing them to
Integer in SendGuildNews or the leader response raises error 6 above 32,767,
aborting the menu response before HOO can display its roster and actions.

Keep these values as Long through both writers. Existing packet IDs are unchanged:

- GuildNews (89): String8 news, String8 guilds, String8 members, UInt8 level,
  **Int32LE current XP, Int32LE required XP**.
- GuildLeaderInfo (94): String8 guilds, String8 members, String8 news,
  String8 applicants, UInt8 level, **Int32LE current XP, Int32LE required XP**.

Strings and separators are unchanged. There is no database migration, XP clamp,
permission change, new packet or capability negotiation. ProtocolGenerator remains retired.

Deploy the companion HOO decoder/model update and this server together. Neither
old-client/new-server nor new-client/old-server combinations are supported.
Do not deploy either half independently.

Suite 33 exercises the actual member and leader packet writers with XP values
0, 32767, 32768, 1000000 and 2147483647. Captures are UNIT_TEST-only; tests check
field order, full-width values and no trailing bytes without network/database I/O.
Native VB6 execution is required to validate these cases.

The reported SendGuildNews log identifies the member response path. It does not
prove a leadership-data defect; this fix preserves existing role selection.
