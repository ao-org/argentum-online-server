# Character save completion and ordering

This is a **world-server-process-local** guard, not account/session coordination
between processes. It does not read or write `is_logged`, change the schema or
wire protocol, or change which character fields are saved.

## Driver evidence

Native Windows VB6 / 32-bit ADO `MSDASQL.1`, SQLite3 ODBC
`sqlite3odbc.dll` version `0.9999` was exercised against a disposable database.
The independent observer used a second connection.

| Check | Observed result |
| --- | --- |
| Async recursive-CTE update (4 million iterations) | `Execute` returned in 0 ms; callback/readback completed in 891 ms |
| Connection state before completion | `adStateOpen` (1), despite unfinished execution |
| Invalid-table async update | Submission returned in 0 ms without an exception; `ExecuteComplete` reported `-2147467259`, no such table |
| Explicit SQL BEGIN / UPDATE / COMMIT | Independent reader saw the old value before COMMIT and the new value after its successful callback |
| Explicit SQL rollback | Independent reader retained the previous committed value |
| Deferred-FK COMMIT failure | COMMIT callback reported failure; subsequent successful ROLLBACK restored the previous value |
| Unrelated synchronous write while a snapshot transaction was open | Calling thread blocked **5375 ms**, then reported database locked |

The last check matters: the server still has synchronous writes in character
creation, guild operations, shop operations and login mapping. A transaction
held across main-loop callbacks would block those calls while requiring that
same loop to submit its next statement. Therefore this implementation deliberately
retains **per-statement SQLite autocommit**, not a multi-statement transaction.
It does not claim snapshot atomicity or rollback after a partial save. Other
installed driver versions/configurations still require these checks.

## Save execution

`LoadCharacterFromDB`, shared by name-based and ID-based login, checks the local
pending/failed registry before reading character state. It also reserves bounded
save capacity and a monotonically increasing process-local save-session ID.
Exhaustion refuses admission; IDs never wrap/reuse within a process.

`SaveCharacterDB` and `SaveChangesInUser` capture SQL and typed parameter copies
before returning. Their job retains the database character ID, save-session ID,
snapshot, completed-statement position, result and error independently of the
player slot. Pet persistence decisions are captured before logout removes pets.

Two retained ADO connection/event wrappers run jobs. Each wrapper owns exactly
one command at a time. It waits for a matching connection **and command object**
in `ExecuteComplete`; callback arrival during submission is supported. Callbacks
record a result; the normal main loop advances execution. An idle connection,
successful submission or elapsed time never establishes completion.

Every required statement must complete successfully before the job succeeds.
The main character update must affect exactly one row. SQLite autocommit errors
are failures, not success. No prepared command or connection belonging to an
unfinished character save is reused. The old rotating bank-command cache is
removed; unrelated generic database operations are not rewritten.

For each admitted character there is room for one autosave and one final save.
Additional autosaves defer without clearing dirty state. A final save is queued
after the existing job, and no further save from that session is accepted.
The final bank snapshot and completed-quest state are captured even when equal
to the last acknowledged state: the older queued job could still change them.

Acknowledgement updates the persisted baseline, not gameplay state. Skill,
quest and collectible-card dirty state is reconciled against the captured
values. Changes made while an older autosave ran remain dirty. A reused player
slot must match both character and save-session IDs before receiving an
acknowledgement. The autosave timestamp advances only on complete success.
The existing manual-save completion message is also deferred until the tracked
queue completes; autosave logs count requests, not falsely acknowledged saves.

### Account collectible-card dependency

Collectible cards are a 1024-byte **account** blob, not character-owned data.
When a full save requires that blob, the job captures its authenticated account
ID and registers an account dependency **before** constructing SQL or enqueueing.
This closes the logout-A/login-B race: a different character on the same account
cannot load an old blob while A's save remains pending, failed or uncertain.

The name/ID admission paths share the guarded character loader. New-character
admission checks before creating its row, and the collection loader checks again
before reading or resetting its arrays. Rejection uses the existing dynamic
message box. A successful character reservation clears its failure message so
inventory loading does not mistake a successful reservation for a load error.

Each dependency retains the exact immutable job, not a user slot or an account
counter. Only that job's verified whole-save completion removes it. Multiple
card-bearing jobs retain separate dependencies; one completion, a duplicate or
an old session's callback cannot release another. Capture/enqueue failures retain
the dependency. Provider-confirmed failure uses the existing reviewed recovery;
the account becomes loadable after **all** its card-bearing jobs succeed.
Failure logs include `card_account` alongside character/session identity.

Unrelated accounts and saves without dirty cards are unaffected. With collectible
cards disabled, the existing no-load/no-write behavior remains; re-enabling does
not discard an unresolved dependency. Dirty-state reconciliation is unchanged.
There is no per-account online-session limit or account SQL transaction.

This only protects subsequent loads in this world-server process. Two characters
that already loaded the same account blob before either save was registered can
still overwrite each other's changes; that preexisting simultaneous-online case
and cross-process writers are outside this fix. The dependency is not durable
across a crash or restart.

**Scope decision:** the manual-save review finding is explicitly deferred.
This card-race correction does not change `GuardarUsuarios`, its notifications,
or manual-save scheduling.

## Paths covered

- Main character fields, spells, inventory (including elemental tags), bank,
  skills, pets, quests, completed quests and inventory skins keep their existing
  SQL/serialization. Ordinary autosaves retain conditional table writes.
- Full saves retain the existing collectible-card feature/dirty condition;
  the blob write now participates in completion instead of swallowing errors.
- Logout token removal is the last ordered statement, not an independent
  rotating-pool submission. `last_logout` keeps its existing SQL semantics.
- `resetPj` quest cleanup/reset marking and marketplace publication-before-kick
  are captured in the same ordered pipeline, rather than bypassing it through
  `Execute`. A reset deferred behind an autosave captures a follow-up snapshot;
  its acknowledgement cannot clear a newer reset generation. Marketplace packet
  handling, price validation, charges and packet layouts are unchanged.
- New character creation keeps its existing synchronous main-row insert. Its
  child-table writes use a retained job, and initial `ConnectUser` occurs only
  after those writes complete. Disconnect during creation retains the job.
- `UserDisconnected`/scenario cleanup still precedes snapshot capture;
  `CloseUser`/`ResetUserSlot` may release the live slot without deleting the job.

This does not serialize arbitrary administrator SQL, offline GM edits, other
processes, guild history or other account-wide writes. Operators must not edit a
blocked character behind its retained snapshot. Those operations are not
character-session snapshot writers and are not covered by this completion result.

## Failure and recovery

A failed job stays at the head of its character's queue; later jobs and login
remain blocked. Other characters on that account also wait if the failed snapshot
includes cards; unrelated accounts can continue. The database error log records
character/session/card-account IDs, whether this was final, completed-statement count, the
failing SQL prefix and the provider error. A submission exception has an
**uncertain outcome**: its exact connection and command are quarantined, never
reused or automatically retried. Missing callbacks remain unresolved; no timeout
declares them successful.

For a provider-confirmed statement failure:

1. Keep the server process alive. Preserve its logs and take an appropriate
   database backup before administrative repair.
2. Inspect the error and partial database state. Correct the underlying problem
   without replacing the character's state or changing completed statements.
3. Use the server form's **Guardar Usuarios** button. When failures exist it
   offers recovery; enter the database character ID and explicitly confirm that
   the failure/partial state has been reviewed. Cancel leaves the job untouched.
4. Recovery resumes at the failed statement; it does **not** re-execute prior
   completed writes. The final queued snapshot still runs afterward. Login is
   enabled only once the retired session's entire queue has succeeded.

For capture failures, uncertain dispatch, or a missing callback, this UI refuses
retry. Keep the process/snapshot for debugging and investigate the retained
command/provider before deciding on manual reconciliation. Do not clear the
guard merely to permit login. There is no claim that an arbitrary uncertain
operation can be retried safely.

Orderly shutdown stops new admission, disconnects users and captures their final
saves. It keeps pumping the normal loop until saves complete; pending/failed
jobs are logged and prevent process exit. It never sleeps or synchronously
drains saves. The operator can review failures while the process remains alive.

**The queue is in memory, not crash-durable.** Forced termination, power loss or
process failure can lose pending data and the local guard. Before restarting
after an unresolved forced shutdown, inspect logs and reconcile/restore affected
characters from the database/backups. This PR does not claim cross-process
exclusion or a durable recovery journal.

## Validation

`Unit_CharacterSaves` is suite 44 in the existing native runner. It exercises
the production scheduler, retained commands and completion callbacks against
disposable SQLite databases, including an original-schema test fixture clone.
It covers ordered delayed autosave/logout, reconnect guarding, provider errors,
partial failure/reviewed resume, independent characters, slot reuse, dirty-state
acknowledgement, bounded admission/final capacity, shutdown, callbacks during
submission and stale/duplicate callbacks. The snapshot case executes real
creation/save SQL and checks the final main/inventory/bank/skill/logout values.

The card regressions clone that disposable fixture and execute the real full-save
capture, `AddCards` UPSERT, asynchronous worker callbacks, character/new-character
admission and collection loader. They verify delayed A/logout versus B/login,
both additions surviving B's later save, multiple outstanding dependencies,
unrelated accounts, reused slots, duplicate/stale callbacks, provider-confirmed
card failure and reviewed recovery, uncertain submission, capture/enqueue failure,
dirty changes during an older save, and feature disable/re-enable behavior.

**Separate preexisting driver limitation, not corrected here:** on the tested
SQLite3 ODBC 0.9999 stack, a declared `BLOB` column reports a defined size of 255.
The production `adUseClient` read exposes only 255 bytes; an `adUseServer` read
reports 1024 bytes but zero-fills the tail. An independent 32-bit ADO probe using
the unchanged table declaration and SELECT reproduced this without the save
scheduler: SQL reported a 1024-byte blob with final byte `0A`, whereas the client
cursor returned 255 bytes. Therefore the race fixtures use full 1024-byte blobs
with distinct nonzero cards within indices 1-255 and zero padding afterward.
Tests do not establish preservation of cards above that range on this driver.
That existing loader/provider issue requires a separate scoped fix; this change
does not alter BLOB reads, values or the SQL representation to hide it.

Build definitions (no transport flag):

Windows validation on 2026-10-24, PR base master
`68efe0f856cbf2837496a5c832d318c586c6626c`:

| PYMMO | DEBUGGING | UNIT_TEST | Native build | Executed suite |
| --- | --- | --- | --- | --- |
| 0 | 0 | 0 | Pass | Not run (production build) |
| 1 | 0 | 0 | Pass | Not run (production build) |
| 0 | 1 | 1 | Pass | 701/701 passed, 44 suites |
| 1 | 1 | 1 | Pass | 701/701 passed, 44 suites |

Compilation used a separate source copy while the user's VB6 IDE remained open.
The test executables ran on loopback port 17669 with disposable fixture databases;
no production database or client/server gameplay session was used. The original
project's local COM-reference paths and compilation defaults were not committed.

```text
VB6.exe /make /d DEBUGGING=1:PYMMO=1:UNIT_TEST=1:UNLOCK_CPU=1 /out <log> Server.VBP
VB6.exe /make /d DEBUGGING=1:PYMMO=0:UNIT_TEST=1:UNLOCK_CPU=1 /out <log> Server.VBP
VB6.exe /make /d DEBUGGING=0:PYMMO=1:UNIT_TEST=0:UNLOCK_CPU=1 /out <log> Server.VBP
VB6.exe /make /d DEBUGGING=0:PYMMO=0:UNIT_TEST=0:UNLOCK_CPU=1 /out <log> Server.VBP
```

Manual acceptance still required: matching development-client logout/reconnect,
new-character admission, inventory/bank/quest/card changes during autosave,
logout while an autosave is slow, orderly shutdown with connected players, and
operator recovery with the deployment's exact driver and database settings.
Do not test these against production data.
