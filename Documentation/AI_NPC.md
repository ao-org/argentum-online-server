# AI_NPC Module Documentation

## Module Purpose

The `AI_NPC.bas` module orchestrates every behaviour exhibited by server-controlled characters. It coordinates movement, combat, spell casting, and situational awareness across the different AI archetypes (`e_TipoAI`). The module pulls data from global registries such as `NpcList`, `UserList`, and `MapData` and relies on helper modules (pathfinding, combat resolution, networking) to enact each decision.

## High-Level Overview

This update delivers a sweeping refinement to ranged NPC movement, pathfinding efficiency, and the long-term stability of every timer-driven behaviour. The following sections expand on each pillar so future maintainers can trace the feature set back to the concrete helpers introduced in PR #1014.

### 1. Smarter Ranged Movement
Ranged NPCs now seek a defensive orbit instead of charging blindly. `ComputeNpcRangedRetreatDestination` blends the backward vector with a configurable tangential component so casters keep their preferred distance while circling the target. `AI_CaminarConRumbo` consumes that destination, applies any pending strafe offset, and drives the NPC around the player with periodic orbit flips coordinated by `EnsureNpcOrbitDirection` and `FlipNpcOrbitDirection`. The result is a fluid orbit that alternates direction when lanes clog, producing lifelike repositioning.

### 2. Reactive Strafing When Hit
When an NPC receives damage from its current target we register a short-lived strafe offset via `RegisterNpcDamageStrafe`, `SetNpcStrafeOffsetFromAttacker`, and `ApplyNpcStrafeToDestination`. The trio nudges the next waypoint sideways, letting the NPC sidestep before resuming its orbit. The offset expires after `NPC_STRAFE_DURATION_MS`, preventing the behaviour from sticking around longer than the configured reaction window.

### 3. Controlled Path Recalculation
`AI_CaminarConRumbo` guards expensive A* calls with `NextPathRecomputeAt` and the `[AI]` setting `PATH_RECOMPUTE_COOLDOWN_MS`. Paths are reused until the cooldown lapses, dramatically reducing CPU spikes during large engagements while keeping the movement responsive.

### 4. Orbit Reevaluation and Timing Stability
All temporal behaviours—strafe windows, orbit reevaluations, and path cooldowns—share the same wrap-safe helpers (`AddMod32`, `TickAfter`). `EnsureNpcOrbitDirection` and `FlipNpcOrbitDirection` store `OrbitDirection` and `OrbitReevaluateAt` on `t_NpcPathFindingInfo`, making orbit changes deterministic even once the millisecond counter wraps after ~49.7 days of uptime.

### 5. Configurable Parameters
Every new timing and weighting constant is tunable from the `[AI]` section of `Configuracion.ini`, with defaults defined in `Consts.bas` and surfaced through `ServerConfig.cls`. Operators can tailor combat pacing and difficulty by tweaking these knobs.

| Setting | Default | Description |
| --- | --- | --- |
| `NPC_PATHFINDING_NOISE` | `0.75` | Adds subtle randomness to pathfinding so multiple NPCs do not pick identical corridors. |
| `NPCCrossMapAllowSafeMaps` | `1` | Allows hostile NPC cross-map pursuit routes to enter or pass through safe maps. Set to `0` to stop pursuit at those boundaries. Requires `npc-cross-map-pursuit` to be enabled. |
| `NPCCrossMapDiagnostics` | `0` | Opt-in routine pursuit diagnostics, read from `[AI]` in `Configuracion.ini`. Does not enable or disable pursuit. Missing key means off. |
| `NPC_ORBIT_REEVALUATE_MS` | `1800` | Minimum duration before an NPC can flip its orbit direction. |
| `PATH_RECOMPUTE_COOLDOWN_MS` | `250` | Minimum delay between A* recomputations once a path exists. |
| `NPC_ORBIT_TANGENT_WEIGHT` | `0.35` | Balance between backward and sideways motion when orbiting a target. |
| `NPC_RETREAT_DISTANCE_BUFFER` | `0.75` | Buffer distance that triggers a slide instead of pure retreat. |
| `NPC_STRAFE_DURATION_MS` | `900` | Lifespan of the reactive strafe after receiving damage. |

### Cross-map pursuit diagnostics

Routine pursuit messages are disabled by default, independently of the
`npc-cross-map-pursuit` gameplay feature toggle. To enable them temporarily, set
the following in the active **Configuracion.ini**, not Server.ini:

```ini
[AI]
NPCCrossMapDiagnostics=1
```

Use the server window's **Recargar configuracion.ini** button to apply it through
`LoadMainConfigFile`, or restart the server. A successful reload is sufficient;
no restart is required. Editing the file alone has no effect until reload/startup.
The existing reload also reloads events/invasions; avoid using it during an
active invasion, as described in the external-configuration documentation.
When using `/config-dir`, edit that directory's managed Configuracion.ini (see
[External configuration](ExternalConfiguration.md)). To disable diagnostics,
set the value to `0` or remove the key, then reload or restart again. Missing
configuration leaves these diagnostics off; existing production configuration
does not need to be changed. The setting is cached in `ServerConfig`, with no
file access or configuration parsing during pursuit processing.

This gates transition checks, retention scans (including zero-target scans),
retained targets, selected/crossed exits, expected route rejections and ordinary
chase cancellations. When enabled, existing NPC index/lifetime, user, map, exit,
target and reason details remain available. String formatting is guarded at the
call sites. Pursuit, targeting, routing, return-home and visibility behavior do
not depend on this diagnostic setting.

Exception handlers still call `TraceError` unconditionally. The two forced
return-home fallbacks, `return route unavailable; warping home` and
`return path failed; warping home`, also remain visible regardless of the setting:
they record an exceptional recovery rather than an ordinary cancellation, and
can help diagnose missing routes or repeated path failures. Their existing
informational severity is preserved, not promoted or downgraded.

`Logging.LogInfoServidor` continues emitting `vbLogEventTypeInformation` through
Windows Event Log source `Argentum20`, event ID 20. `TraceError` emits
`vbLogEventTypeError`. This change does not alter global logging or unrelated
events. No monitoring trigger/export configuration is checked into this
repository, so the rule behind the reported Warning/PROBLEM alert remains
**unverified**. An alert label alone does not establish the source event's
severity; do not broadly exclude all `NPC cross-map` messages in monitoring.

### 6. Structural and Initialization Changes
`t_NpcPathFindingInfo` now stores `StrafeOffset`, `StrafeExpiresAt`, `OrbitDirection`, `OrbitReevaluateAt`, and `NextPathRecomputeAt`. These fields are reset inside `ResetNpcMainInfo` and `OpenNPC`, eliminating the risk of leaking movement state between instances. Consolidating AI constants into `Consts.bas` and `ServerConfig.cls` keeps configuration handling in one place.

### 7. Pathfinding Noise and Randomization
Pathfinding integrates heading shuffling and the `NPC_PATHFINDING_NOISE` scalar. The extra noise diversifies routes chosen by simultaneous movers, giving ranged NPCs distinct orbit lanes and preventing clumped zig-zag patterns.

### Summary
The ranged AI overhaul now combines reactive strafing, orbit reevaluation, and throttled path recomputation to produce adversaries that feel deliberate without overloading the server. Because every timer uses wrap-safe arithmetic and every tuning knob is configurable, shards can run for weeks while fine-tuning combat personality.

## Key Public Subroutines and Functions

* **`NpcDummyUpdate(NpcIndex As Integer)`** – Updates practice dummies, including their health regeneration and idle logic.
* **`NpcAI(NpcIndex As Integer)`** – Main entry point for NPC logic. Dispatches to the appropriate AI routine (`AI_RangeAttack`, `AI_GuardiaPersigueNpc`, etc.) based on the configured movement mode.
* **`AI_RangeAttack(NpcIndex As Integer)`** – Handles ranged target selection, attack cadence, and now feeds retreat/orbit destinations into `AI_CaminarConRumbo`.
* **`AI_CaminarConRumbo(NpcIndex As Integer, ByRef rumbo As t_WorldPos)`** – Applies strafe offsets, honours pathfinding cooldowns, and drives the NPC toward the supplied waypoint.
* **`EnsureNpcOrbitDirection` / `FlipNpcOrbitDirection`** – Cache and flip the active orbit side, scheduling reevaluations with wrap-safe timers.
* **`ComputeNpcRangedRetreatDestination(...) As t_WorldPos`** – Generates the blended retreat/orbit point around the current target.
* **`RegisterNpcDamageStrafe` / `ApplyNpcStrafeToDestination`** – Store and consume the short-lived strafe offset after taking damage.
* **`SelectNearestUser`, `SelectNearestNpc`, `SelectCurrentTarget`** – Determine the most relevant hostile or friendly entities within range.
* **`GoToNextWp(NpcIndex As Integer)`** – Retrieves waypoints for scripted patrols and battleground behaviours.

## Notable Dependencies

The module interacts with numerous systems:

* **Global Data Structures:** `NpcList`, `UserList`, `MapData`, `Hechizos`, `Invasiones`.
* **Supporting Modules:** Pathfinding (`SeekPath`, `FollowPath`), combat (`SistemaCombate`), networking/protocol helpers, logging utilities, and server configuration accessors.
* **Utility Functions & Constants:** Distance helpers (`Distancia`), random number generation, heading math, and behaviour flags that gate each action.

Together these dependencies provide the raw data and services needed for the AI routines described above.
