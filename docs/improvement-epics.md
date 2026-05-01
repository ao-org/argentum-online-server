# Argentum Online — Improvement Epics

Ordered from quick wins to most difficult. Each epic includes scope, goals, and actionable requirements.

---

## Epic 1: Architecture Documentation
**Effort:** Low | **Impact:** Medium | **Estimated Time:** 1–2 weeks

### Goal
Create living documentation that helps new contributors onboard quickly and gives the team a shared understanding of the system.

### Requirements
- [ ] Write an onboarding guide for new contributors (setup, build, run, test)
- [ ] Document the high-level architecture of both client and server (expand on existing `ProjectOverview.mmd`)
- [ ] Create Architecture Decision Records (ADRs) for key past decisions (VB6, SQLite, Aurora.Network, CryptoSys, feature toggles)
- [ ] Document the game loop lifecycle (server tick, client render loop)
- [ ] Document the module dependency graph for the server (which `.bas`/`.cls` calls which)
- [ ] Document the module dependency graph for the client
- [ ] Add inline comment standards and examples for future contributions

---

## Epic 2: Document Network Protocol
**Effort:** Low | **Impact:** High | **Estimated Time:** 1–2 weeks

### Goal
Create a complete reference of the client↔server binary protocol so protocol changes are safer and onboarding is faster.

### Requirements
- [ ] Document every `ServerPacketID` enum value with: packet name, direction, payload fields, field types, and byte sizes
- [ ] Document every `ClientPacketID` enum value with the same detail
- [ ] Document the packet framing format (header, length prefix, serialization order)
- [ ] Document the handshake/login sequence (connection → authentication → character selection → gameplay)
- [ ] Document version negotiation and how client/server protocol sync is enforced
- [ ] Add notes on which packets are rate-limited and how
- [ ] Keep the document in a format that can be referenced by the ProtocolGenerator project

---

## Epic 3: Configuration Validation & Centralization
**Effort:** Low | **Impact:** Medium | **Estimated Time:** 1–2 weeks

### Goal
Prevent server crashes from misconfigured INI files and make configuration easier to reason about.

### Requirements
- [ ] Create a startup validation routine that checks all required INI files exist before the server starts
- [ ] Validate required keys and value types (numeric ranges, non-empty strings) for `Server.ini`
- [ ] Validate `feature_toggle.ini` structure (TOGGLECOUNT matches actual toggle count, names are non-empty, values are 0 or 1)
- [ ] Validate `Configuracion.ini` required fields
- [ ] Log clear error messages when validation fails, including the file name, key, and expected format
- [ ] Centralize path resolution (DatPath, MapPath, etc.) into a single module instead of scattering path reads across multiple modules
- [ ] Add a `DefaultConfig` fallback mechanism for non-critical missing values

---

## Epic 4: Centralize Constants & Eliminate Magic Numbers
**Effort:** Medium | **Impact:** High | **Estimated Time:** 2–3 weeks

### Goal
Replace hardcoded numeric literals scattered throughout the codebase with named constants, making game balance tuning and bug hunting much easier.

### Requirements
- [ ] Audit server code for hardcoded map IDs, item IDs, NPC IDs, spell IDs, and class/race IDs
- [ ] Audit client code for the same categories of magic numbers
- [ ] Create a `GameConstants.bas` module on the server with named constants grouped by category (maps, items, NPCs, spells, classes, races, limits)
- [ ] Create a matching `GameConstants.bas` on the client for client-side constants
- [ ] Replace all identified magic numbers with references to the new constants
- [ ] Move game balance values (damage formulas, XP curves, gold thresholds) to named constants or config-driven values
- [ ] Add constants for protocol limits (max inventory slots, max spells, max bank slots, etc.)
- [ ] Ensure `Consts.bas` (server) is reviewed and merged/deduplicated with the new constants module


---

## Epic 5: Standardize Error Handling & Logging
**Effort:** Medium | **Impact:** High | **Estimated Time:** 2–4 weeks

### Goal
Eliminate silent failures, make production issues diagnosable, and establish a consistent error handling pattern across the codebase.

### Requirements
- [ ] Audit all `On Error Resume Next` usages in the server — classify each as intentional or accidental
- [ ] Audit all `On Error Resume Next` usages in the client — same classification
- [ ] Remove accidental silent error suppression and replace with proper `On Error GoTo` handlers
- [ ] Standardize error handler naming convention (e.g., `ErrHandler` label in every procedure)
- [ ] Create a centralized `LogError` routine that captures: module name, procedure name, error number, error description, and timestamp
- [ ] Ensure all error handlers call the centralized `LogError` routine
- [ ] Implement log levels (Error, Warning, Info, Debug) in the `Logging.bas` module
- [ ] Add log rotation or size limits to prevent unbounded log file growth
- [ ] Ensure critical game events are logged at Info level (player login/logout, trades, bans, guild actions)
- [ ] Review and improve the existing `Logging.bas` module on the server to support the above

---

## Epic 6: Expand Unit Test Coverage
**Effort:** Medium | **Impact:** High | **Estimated Time:** 3–4 weeks

### Goal
Increase confidence in core game systems by expanding the existing unit test framework to cover critical modules that currently lack tests.

### Requirements
- [ ] Identify all server modules without corresponding test suites (compare `Codigo/Tests/` against `Codigo/*.bas`)
- [ ] Identify all client modules without corresponding test suites
- [ ] Write unit tests for `Protocol.bas` — verify packet serialization/deserialization round-trips
- [ ] Write unit tests for `Database_Queries.bas` — verify SQL query construction
- [ ] Write unit tests for `CharacterPersistence.bas` — verify save/load data integrity
- [ ] Write unit tests for `AI_NPC.bas` — verify NPC behavior state transitions
- [ ] Write unit tests for `EffectsOverTime.bas` — verify effect application, duration, and removal
- [ ] Write unit tests for `InvUsuario.bas` — verify inventory operations (add, remove, equip, stack)
- [ ] Write unit tests for `modHechizos.bas` — verify spell effect calculations
- [ ] Write unit tests for `ModGrupos.bas` — verify party/group logic
- [ ] Ensure all new tests follow the existing `UnitTesting.bas` framework conventions
- [ ] Add a test coverage checklist to the contributing guide so new features require tests

---

## Epic 7: Security Audit
**Effort:** Medium | **Impact:** High | **Estimated Time:** 2–3 weeks

### Goal
Identify and fix security vulnerabilities in authentication, data handling, and server administration.

### Requirements
- [ ] Audit password hashing implementation — verify salt generation, hash algorithm strength, and timing-safe comparison
- [ ] Audit all SQL query construction for injection vulnerabilities (string concatenation vs parameterized queries)
- [ ] Remove admin credentials from `Server.ini` example file or document that they must be changed
- [ ] Audit `AO20CryptoSysWrapper.bas` usage — verify encryption is applied correctly to sensitive data
- [ ] Review rate limiting on login attempts, character creation, and trade operations
- [ ] Audit packet handling for buffer overflow or malformed packet vulnerabilities
- [ ] Review anti-cheat implementation (`AntiCheat.bas`) for bypass vectors
- [ ] Verify that GM commands are properly permission-gated (no privilege escalation paths)
- [ ] Document all findings and create follow-up tasks for each vulnerability found
- [ ] Ensure recovery codes and validation codes are generated with cryptographically secure randomness

---

## Epic 8: Database Layer Hardening
**Effort:** Medium | **Impact:** High | **Estimated Time:** 2–3 weeks

### Goal
Improve database reliability, performance, and safety.

### Requirements
- [ ] Audit all queries in `Database_Queries.bas` for SQL injection (string interpolation vs parameterized)
- [ ] Convert any string-concatenated queries to use parameterized queries via ADO
- [ ] Implement a connection management wrapper that handles open/close lifecycle and error recovery
- [ ] Add retry logic for transient SQLite errors (SQLITE_BUSY, SQLITE_LOCKED)
- [ ] Implement a simple in-memory cache for frequently read, rarely changed data (item definitions, NPC definitions, map metadata)
- [ ] Verify that WAL mode is enabled at startup and log a warning if it's not
- [ ] Add database migration rollback support (currently only forward migrations exist in `ScriptsDB/`)
- [ ] Add indexes for frequently queried columns not yet indexed (audit slow query patterns)
- [ ] Ensure all database writes within a logical operation use transactions (e.g., character save should be atomic)

---

## Epic 9: Extract Shared Code (Client/Server)
**Effort:** Medium | **Impact:** Medium | **Estimated Time:** 2–3 weeks

### Goal
Reduce code duplication between client and server for modules that are functionally identical or near-identical.

### Requirements
- [ ] Diff `PacketId.bas` between client and server — document any divergences and determine if they should be unified
- [ ] Diff `AO20CryptoSysWrapper.bas` / `basCryptoSys.bas` between client and server
- [ ] Diff `clsIniManager.cls` between client and server
- [ ] Diff `clsNetReader.cls` / `clsNetWriter.cls` between client and server
- [ ] Diff `ModAreas.bas` between client and server
- [ ] Diff `modElapsedTime.bas` between client and server
- [ ] For each shared module, decide: (a) keep in sync manually, (b) extract to a shared repo/DLL, or (c) accept divergence
- [ ] If option (b), create a shared VB6 project or document a copy-sync process
- [ ] Ensure the ProtocolGenerator project (private repo) is the single source of truth for `PacketId.bas` and `Protocol.bas`
- [ ] Document the sync strategy in the contributing guide

---

## Epic 10: Refactor Large Modules
**Effort:** High | **Impact:** Medium | **Estimated Time:** 4–6 weeks

### Goal
Break down monolithic modules into smaller, focused units to improve readability, testability, and maintainability.

### Requirements
- [ ] Identify all server modules over 500 lines — rank by size and coupling
- [ ] Identify all client modules over 500 lines — same ranking
- [ ] Refactor `GameLogic.bas` (server) — extract distinct responsibilities into separate modules
- [ ] Refactor `Protocol.bas` (server) — group packet handlers by game system (combat, commerce, guild, etc.)
- [ ] Refactor `Protocol.bas` (client) — same grouping strategy
- [ ] Refactor `General.bas` (server) — extract utility functions vs game logic
- [ ] Refactor `SistemaCombate.bas` — separate damage calculation, hit resolution, and death handling
- [ ] Refactor `TileEngine.bas` (client) — separate rendering, map loading, and tile logic
- [ ] For each refactored module, ensure existing unit tests still pass
- [ ] For each refactored module, add new unit tests for the extracted pieces
- [ ] Update the module dependency documentation after each refactor

---

## Epic 11: Introduce Layered Architecture
**Effort:** High | **Impact:** High | **Estimated Time:** 6–10 weeks

### Goal
Separate concerns into clear layers (data access, business logic, protocol/UI) so that changes in one layer don't ripple across the entire codebase.

### Requirements
- [ ] Define the target layer boundaries: Data Access → Game Logic → Protocol Handler → Network I/O
- [ ] Create a Data Access Layer (DAL) module that wraps all `Database_Queries.bas` calls behind a clean interface
- [ ] Ensure no game logic module directly executes SQL — all DB access goes through the DAL
- [ ] Create a Game Services layer that encapsulates business rules (combat, commerce, guilds, quests)
- [ ] Ensure protocol handlers (`Protocol.bas`, `Protocol_Writes.bas`) only parse/serialize packets and delegate to Game Services
- [ ] Remove direct UI manipulation from game logic modules on the server (e.g., no direct form references from `.bas` modules)
- [ ] On the client, separate rendering logic from game state management
- [ ] Define clear interfaces (VB6 abstract classes) for each layer boundary where feasible
- [ ] Migrate one game system end-to-end as a proof of concept (suggested: commerce or banking — well-scoped, has DB + protocol + logic)
- [ ] Document the layered architecture pattern and update the contributing guide
- [ ] Ensure all existing tests pass after each migration step
