# AI_NPC Module Documentation

## 1. General Purpose

The `AI_NPC.bas` module is responsible for managing the Artificial Intelligence (AI) of Non-Player Characters (NPCs) within the Argentum 20 game server. It defines how NPCs behave in various situations, including their movement patterns, combat strategies, interactions with players and other NPCs, and responses to game events. The module uses a state-based approach, where an NPC's behavior is determined by its assigned AI type (`e_TipoAI`).

## 2. Main Public Subroutines and Functions

Here's a list of the primary public subroutines and functions provided by the `AI_NPC` module:

*   **`NpcDummyUpdate(NpcIndex As Integer)`**: Manages updates for "dummy" NPCs. These appear to be simple target practice NPCs that regenerate their health over time.
*   **`NpcAI(NpcIndex As Integer)`**: This is the central AI processing subroutine. It's called periodically for each NPC. Based on the `NpcList(NpcIndex).Movement` type, it dispatches control to more specialized AI routines (e.g., static, random movement, following master, attacking, participating in invasions).
*   **`SelectNearestUser(NpcIndex As Integer, ByRef NearestTargetDistance As Single) As Integer`**: Scans the NPC's vicinity to find the closest user (player) that is a valid target.
*   **`SelectNearestNpc(NpcIndex As Integer, ByRef NearestTargetDistance As Single) As Integer`**: Scans the NPC's vicinity to find the closest NPC that is a valid target.
*   **`SelectCurrentTarget(NpcIndex As Integer, ByVal NearestUser As Integer) As t_AnyReference`**: Determines the NPC's current target. It prioritizes targets based on factors like who attacked the NPC (`AttackedBy`) or if the NPC is under a 'taunt' effect.
*   **`AI_RangeAttack(NpcIndex As Integer)`**: Implements the logic for NPCs that perform ranged attacks. It handles target acquisition, aiming, attacking, and repositioning based on preferred range.
*   **`ComputeNextHeadingPos(NpcIndex As Integer) As t_WorldPos`**: Calculates the map coordinates of the tile directly in front of the NPC, based on its current heading.
*   **`NPCHasAUserInFront(NpcIndex As Integer, ByRef UserIndex As Integer) As Boolean`**: Checks if there is a user character located on the tile immediately in front of the NPC.
*   **`AI_GuardiaPersigueNpc(NpcIndex As Integer)`**: Defines the behavior for guard NPCs that are tasked with pursuing other (presumably hostile) NPCs within a certain area.
*   **`AI_SupportAndAttackNpc(NpcIndex As Integer)`**: Controls NPCs that have a dual role of supporting friendly entities (e.g., healing) and attacking hostile targets.
*   **`AI_BgTankBehavior(NpcIndex As Integer)`**: AI logic for NPCs acting as "tanks" in battleground scenarios, focusing on engaging the nearest enemy (user or NPC).
*   **`AI_BGSupportBehavior(NpcIndex As Integer)`**: AI logic for support-oriented NPCs in battlegrounds, prioritizing helping allies and then potentially attacking or moving strategically.
*   **`AI_BGRangedBehavior(NpcIndex As Integer)`**: AI for ranged NPCs in battlegrounds, managing attacks and movement to maintain optimal distance.
*   **`AI_BGBossBehavior(NpcIndex As Integer)`**: AI for boss NPCs in battlegrounds, with specific logic for engaging targets and a leash range to return to origin.
*   **`AI_BGBossReturnToOrigin(NpcIndex As Integer)`**: Handles the behavior of a battleground boss returning to its starting position if it moves too far.
*   **`AI_NpcAtacaNpc(NpcIndex As Integer, Optional ByVal ChangeTargetMovement As Boolean = True)`**: Manages the combat logic when an NPC attacks another NPC.
*   **`SeguirAmo(NpcIndex As Integer)`**: Implements the behavior for NPCs (pets, summons) that are designated to follow a specific user (their "master").
*   **`GetAvailableSpellEffects(NpcIndex As Integer) As Long`**: Iterates through an NPC's known spells and returns a bitmask representing all the types of spell effects (e.g., heal, damage, paralyze) it can currently cast based on cooldowns.
*   **`TryCastHelpSpell(NpcIndex As Integer, ByVal AvailableSpellEffect As Long) As Boolean`**: Attempts to make the NPC cast a beneficial/support spell on a valid friendly target.
*   **`TryCastAttackSpell(NpcIndex As Integer, ByVal AvailableSpellEffect As Long) As Boolean`**: Attempts to make the NPC cast an offensive spell on a valid hostile target.
*   **`TrySupportThenAttackSpells(ByVal NpcIndex As Integer) As Boolean`**: A higher-level spellcasting routine that prioritizes casting support spells if needed, otherwise attempts to cast attack spells.
*   **`GoToNextWp(NpcIndex As Integer) As t_WorldPos`**: Retrieves the next waypoint coordinates for an NPC that is following a predefined path or battleground waypoints.
*   **`SetMovement(ByVal NpcIndex As Integer, ByVal NewMovement As e_TipoAI)`**: Changes the current AI movement type for a given NPC and updates its tile blocking state accordingly, especially for pets.

## 3. Notable Dependencies

The `AI_NPC.bas` module relies heavily on several other modules and global data structures:

*   **Global Data Structures:**
    *   `NpcList()`: An array or list containing detailed information about each NPC in the game (e.g., stats, position, AI state, target). This is the primary data source and target for most AI operations.
    *   `UserList()`: An array or list containing detailed information about each player character. Used for targeting and interaction checks.
    *   `MapData()`: Represents the game world's map, storing information about tiles, including which NPCs or users occupy them.
    *   `Hechizos()`: An array or list containing definitions and properties of all available spells.
    *   `Invasiones()`: Data structure holding information about invasion events.
*   **Modules (Implicit or Explicit):**
    *   `MODULO_NPCs` (or similar, often referred to as `NPCs`): Provides utility functions for NPC management, such as checking if an NPC can move (`NPCs.CanMove`), attack (`NPCs.CanAttack`), or interact (`NPCs.CanHelpUser`, `NPCs.CanHelpNpc`).
    *   `MODULO_USUARIOS` (or similar, often referred to as `UserMod`): Provides utility functions for user management, like checking visibility (`UserMod.IsVisible`).
    *   `SistemaCombate`: Likely handles the mechanics of combat, such as `SistemaCombate.NpcAtacaNpc`.
    *   `ModAreas`: Manages game world areas and the entities (users, NPCs) within them, used for proximity checks (`ModAreas.ConnGroups`).
    *   `ModReferenceUtils`: Provides utilities for working with game object references (e.g., `ModReferenceUtils.GetPosition`).
    *   Pathfinding Module (e.g., functions like `SeekPath`, `FollowPath`): Used for NPC movement along calculated paths.
    *   Networking/Protocol Module (e.g., functions like `SendData`, `PrepareMessageTextOverChar`, `PrepareCreateProjectile`): For sending updates to clients.
    *   Logging Module (e.g., `LogError`, `TraceError`): For error handling and debugging.
    *   Server Configuration (`SvrConfig`): To access server-side settings (e.g., `SvrConfig.GetValue("NPC_SPELL_RANGE_X")`).
*   **Global Functions & Constants:**
    *   Numerous functions for distance calculation (`Distancia`), random number generation (`RandomNumber`), heading calculation (`GetHeadingFromWorldPos`), character manipulation (`ChangeNPCChar`, `MoveNPCChar`), etc.
    *   Constants defining NPC types, vision ranges (e.g., `RANGO_VISION_X`), spell effects (`e_SpellEffects`), AI alignment (`e_Alineacion`), and behavior flags (`e_BehaviorFlags`).

This module is central to the dynamic behavior of NPCs and interacts with many core systems of the game server.
