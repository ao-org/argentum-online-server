# Native light-orb regression fixture

This VB6 project compiles the production `Codigo/modMagicLightOrbs.bas` with
deterministic map, clock, capability and packet-capture stubs. It checks class
and placement rules, replacement, overlapping casters, remaining-time
snapshots, expiry, tick wraparound, disconnect and map cleanup. It does not
connect to a database or start the game server.

From the server repository root, create the ignored `build` directory, then
compile `Fixtures/magic_light_orb/light_orb_check.vbp` with the installed VB6
compiler's `/make` option. Run `build/magic_light_orb_check.exe`; its result is
written to `build/magic-light-orb-results.txt`. Success is `PASS 26 checks`.
If the compiler has a Windows compatibility setting that requests elevation,
set `__COMPAT_LAYER=RunAsInvoker` for this compiler process.

Also compile the complete `Server.VBP` with the normal production conditional
constants before deployment. The fixture isolates orb ownership and timing;
the full build checks its integration with spell dispatch and the wire writer.
Live casting still requires HOO and the companion resources, with
`hoo-magic-light-orbs-v1` enabled in the server feature toggles.
