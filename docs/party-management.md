# Party management

The extended UpdateGroupInfo packet is the single authoritative management/roster
stream. It includes a correlated request/result, local versioned user identity,
invitation safety, party ID/roster revision, explicit leader identity, and up to six
versioned member identities, names, character/head IDs and HP/MP values. Empty
snapshots clear membership. Values are Int32 for HP/MP; zero maxima mean unavailable.

Client request319 (production PYMMO=1): Int32 positive request ID, Byte action
(0 refresh,1 create,2 invite by name,3 kick,4 leave,5 disband), Int32 party ID,
Int32 roster revision, Int16 target user slot/version, String8 name <=30 characters.
String8 uses the existing variable-length prefix/UTF-8 codec. State response uses
the existing eUpdateGroupInfo ID with its new required layout. Result codes0..18
are localized by HOO. Name/row counts and identities are validated before mutation.
No feature flags, retired-client compatibility, proto generation or DB changes.

Create is idempotent and installs a real one-member roster immediately. Joining
an existing party does not reallocate identity or reset fair XP rotation. Leaving
or kicking the last nonleader retains the leader's party. Leader disconnect or
explicit disband ends it. One-member parties take the historical solo XP paths;
OnlyGroups admission still requires at least two valid members. Removal evicts a
lone leader from restricted maps through their configured exit, when present.

Named and targeted invitations share the same predicate: connected/alive players,
same map and existing Manhattan range<=8, GM/safety/faction/level/leadership rules,
leader authority, capacity and pending-question protection. Existing Accept/Decline
questions are reused. Proposals retain the inviter's versioned identity and issued
party ID; acceptance revalidates all requirements and current capacity. Old proposals
cannot join a recreated party. Invitation success requires the question to enqueue.

Kick/leave/disband use party identity and roster revision; kick additionally uses
the intended versioned member identity, never a shifted slot. Retired uncorrelated leave/disband and positional
kick requests are rejected with refreshed state. Duplicate/newer action correlation
is cached once per current user session, bounded by the user pool. No action is
optimistically applied by HOO. Live party/member combat targeting remains separate.

## Validation

The AO Server PR Validator built/tested source commit `19457183`: **1004/1004
passed**, including all 31 new management cases and the existing 53 party reward
cases (67,220.88 ms total). This includes production packet dispatcher and
Accept/Decline integration. The documented local VB6 compiler is unavailable;
this is native CI evidence, not a claimed local build. Separate PYMMO=0/1 and
UNIT_TEST=0/1 matrix builds are not exposed independently by this validator.
Windows-1252/CRLF and `git diff --check` pass locally. No live gameplay has been
verified. Client draft: https://github.com/morgolock/heroes-of-old/pull/853.

## Coordinated release and two-client checklist

Deploy matching HOO/server together while logins are closed; the roster packet's
required layout changes. No assets or database migration is needed. Do not merge
or deploy automatically.

1. Open Party from the main menu or configured Party binding. Create, verify one
   named leader/You row and no optimistic success before the server replies.
2. Enable invitation popups on the second character. Send a named invitation within
   eight tiles on the same map, accept/decline through the existing question.
3. Repeat through right-click Invite. Check blocked, far/different-map, level/faction,
   full party, pending question and expired/recreated party outcomes.
4. Fill six rows; check HP/MP, selection, confirmation naming, stale kick after a
   simultaneous membership change, and action buttons at supported UI scales.
5. Leave/kick the last member: leader remains and can invite again. Disband and
   leader disconnect clear all members. Relog/character switch clears pending UI.
6. Verify one-person XP is solo and group-only map admission requires two. Repeat
   shared quests/items and small-XP fairness tests in a multi-member party.
7. Check classic/modern UI, compact HUD visibility/targeting, text focus, keybinding,
   pointer press/release consumption, overlay/focus loss and session resets.

## Invitation operation feedback

Roster synchronization and operation feedback are separate. The required
UpdateGroupInfo layout now includes an Int32 feedback sequence immediately after
the result byte. State/refresh snapshots use result0 and sequence0. Each actual
outcome has a positive, session-scoped sequence; duplicate correlated replies reuse
their original sequence. Result17 explicitly means an invitation was sent;
result18 means the recipient joined. Generic success is never inferred to mean
an invitation was sent. Member broadcasts remain silent.

HOO immediately presents each new outcome inline when Party is open, or in the
existing localized system console when closed. Closing while waiting does not
hide the eventual result. Refresh and invitation-safety synchronization preserve
the last outcome. Sequence suppression prevents duplicate feedback; session reset
clears transient state. A stale roster revision cannot discard an independent
operation result. The redundant legacy inviter-decline console line was removed;
the invitee's existing decline confirmation remains.

Matching client/server builds are required; no new negotiation or versioning
mechanism is introduced. Manual checks: with Party closed, invite a player who
blocks popups, invite successfully and decline, fill the final slot before another
invitation is accepted, and close Party while a named invite awaits its reply.
Check each explanation once in EN/ES and classic/modern UI, then refresh and
verify no replay. These live two-client checks remain unperformed.

Feedback source `29418f78` passed the native full validator: **1009/1009**,
including all 36 management cases (67,526.81 ms). Final acceptance-result
assertions are additionally verified by the latest PR validator run. The local
VB6 compiler and independent conditional build matrix remain unavailable.
