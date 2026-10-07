# GM trigger command regression tests

Run `./Codigo/Tests/TriggerCommand/run.ps1` in PowerShell with VB6 and the server's registered Aurora.Network dependency available. Pass `-Compiler` to override the compiler location.

The script extracts the current `HandleSetTrigger`, `HandleAskTrigger`, `e_PlayerType`, and `e_Trigger` declarations into an ignored build directory and compiles them with a small test harness. Logging and outbound messages are captured in memory; it does not start the server, open sockets, read maps, or access a database.

Coverage includes values 0, 255, 256, 32768, 65536, 16909060, and 2147483647; both allowed roles; all four restricted roles, zero/unknown privileges, and conflicting role bits; negative and truncated payloads; query/confirmation text; audit calls; and preservation of the next packet after the four-byte set payload. Results are written under `build/trigger-command-tests` and a failed check returns a nonzero process exit code.

The command transport changes from a one-byte value to a signed four-byte value. Release it with matching HOO support. The existing scalar trigger meanings and map file schema are unchanged; this PR does not perform the later bitmask migration.
