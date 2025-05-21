# GameLogic Module (Extra.bas) Documentation

## 1. General Purpose

The `GameLogic.bas` module, internally referred to as `Extra`, serves as a collection of miscellaneous but essential utility functions and game logic routines for the Argentum 20 game server. It doesn't focus on a single system but rather provides a wide array of helper functions and procedures that are used by many other parts of the server.

Its responsibilities include:

*   **Positional Logic:** Validating coordinates, checking if positions are legal for players or NPCs to stand on or walk to, finding the closest legal positions, and determining map boundaries.
*   **Character Status & Faction Logic:** Checking player status (newbie, GM, alignment like citizen/criminal), faction affiliations (Armada, Caos), and interactions between factions (who can attack/help whom).
*   **Tile-Based Events:** Processing actions that occur when a player steps on a specific map tile, such as teleports, map exits, applying map restrictions (e.g., newbie-only maps, PK-only maps), and triggering healing/resurrection tiles.
*   **Information Display:** Providing detailed information to players when they "look at" a tile (identifying objects, NPCs, other players, and their statuses). It also handles displaying the server's help file.
*   **NPC Interaction:** Making NPCs deliver random expressions.
*   **Player Management:** Includes a critical function `resetPj` for completely resetting a player character's progress.
*   **Utility Functions:** Name-to-UserIndex lookups, IP-to-UserIndex lookups, heading calculations, byte array to string conversion, and debug packet ID to string conversion.
*   **Game Rules & Constraints:** Enforcing map entry restrictions, item usage requirements for spells, and interaction rules (e.g., why an attack or help action failed).

This module is critical for the day-to-day functioning of the game world, ensuring rules are followed and interactions behave as expected.

## 2. Main Public Subroutines and Functions

Key public functions and subroutines in this module include:

*   **`esCiudad(map As Integer) As Boolean`**: Checks if the given map index corresponds to a city map.
*   **`AgregarAConsola(Text As String)`**: Appends a line of text to the main server console GUI.
*   **`NameIndex(username As String) As t_UserReference`**: Efficiently retrieves a user's reference (index and version ID) by their username using an internal dictionary for quick lookups.
*   **`FindLegalPos(UserIndex As Integer, Map As Integer, ByRef X As Byte, ByRef Y As Byte)`**: Finds a valid, unoccupied tile near the given X, Y coordinates for a user, often used after teleportation to prevent stacking.
*   **`EsNewbie(UserIndex As Integer) As Boolean`**: Returns true if the user is considered a "newbie" (typically based on level).
*   **`esCiudadano`, `esCriminal`, `esArmada`, `esCaos`**: Functions to check a user's alignment or faction status.
*   **`FactionCanAttackFaction(...) As Boolean`**, **`FactionCanHelpFaction(...) As e_InteractionResult`**: Determine inter-faction relationship rules.
*   **`EsGM(UserIndex As Integer) As Boolean`**: Checks if a user has Game Master privileges.
*   **`DoTileEvents(UserIndex As Integer, Map As Integer, X As Integer, Y As Integer)`**: The core handler for actions triggered by a user moving onto a specific tile (teleports, map exits, traps, healing points, map restrictions).
*   **`ClearAttackerNpc(UserIndex As Integer)`**: Resets NPC aggression towards a user, typically when the user leaves the area or logs off.
*   **`InRangoVision(UserIndex As Integer, X As Integer, Y As Integer) As Boolean`**: Checks if coordinates are within a user's visual range.
*   **`InRangoVisionNPC(NpcIndex As Integer, x As Integer, y As Integer) As Boolean`**: Checks if coordinates are within an NPC's visual range.
*   **`InMapBounds(Map As Integer, X As Integer, Y As Integer) As Boolean`**: Validates if coordinates are within the defined map limits.
*   **`ClosestLegalPosNPC(...) As t_WorldPos`**: Finds the nearest valid spawn location for an NPC.
*   **`ClosestLegalPos(Pos As t_WorldPos, ByRef nPos As t_WorldPos, ...)`**: Finds the closest valid tile for a player, considering terrain like water/land.
*   **`ClosestStablePos(Pos As t_WorldPos, ByRef nPos As t_WorldPos)`**: Finds the closest valid tile that is not a teleporter.
*   **`HeadtoPos(Head As e_Heading, ByRef Pos As t_WorldPos)`**: Calculates the coordinates of the tile in front of a given position and heading.
*   **`LegalPos(...) As Boolean`**, **`LegalWalk(...) As Boolean`**, **`LegalPosNPC(...) As Boolean`**, **`LegalWalkNPC(...) As Boolean`**: A suite of functions to determine if a specific tile is valid for standing or moving onto, for players and NPCs, considering various conditions like terrain, occupants, and blockages.
*   **`SendHelp(Index As Integer)`**: Sends the contents of the server's help file (`Help.dat`) to the specified user.
*   **`Expresar(NpcIndex As Integer, UserIndex As Integer)`**: Makes the specified NPC say a random phrase from its list of expressions.
*   **`LookatTile(UserIndex As Integer, Map As Integer, X As Integer, Y As Integer)`**: A comprehensive function that provides a textual description to the user of what they see when they click on a map tile. This can include item names, player names and statuses, NPC names and statuses, and quest updates.
*   **`FindDirection(Pos As t_WorldPos, Target As t_WorldPos) As e_Heading`**: Calculates the general direction from a source position to a target position.
*   **`CargarMapasEspeciales()`**: Loads configurations for special map types (interdimensional, event maps) from `MapasEspeciales.dat`.
*   **`resetPj(UserIndex As Integer, Optional borrarHechizos As Boolean)`**: Performs a full character reset for the specified user, wiping stats, skills, inventory (then re-adding starting items), spells (optional), and faction status, then warps them to their starting location.
*   **`ResucitarOCurar(UserIndex As Integer)`**: Revives a dead player or fully heals an alive player if they are on a tile with an "AUTORESU" trigger.
*   **`TestRequiredEquipedItem(...) As e_SpellRequirementMask`**: Checks if a player has the necessary items equipped to cast a certain spell.
*   **`SendrequiredItemMessage(...)`**, **`SendHelpInteractionMessage(...)`**, **`SendAttackInteractionMessage(...)`**: Send predefined messages to users explaining why certain actions (spell casting, helping, attacking) failed.
*   **`PreferedTileForDirection(...) As t_WorldPos`**: Calculates a suitable tile to move to when an entity is trying to move away from a target while avoiding obstacles.

## 3. Notable Dependencies

The `GameLogic` (Extra) module is highly interconnected and relies on many other parts of the server:

*   **Global Data Structures:**
    *   `UserList()`: Primary source for all user-related data (position, stats, inventory, flags, factions, quests).
    *   `NpcList()`: Primary source for NPC data (position, stats, expressions, movement type).
    *   `MapData()`: Contains all tile-specific information (blockages, occupants, triggers, TileExit for portals).
    *   `ObjData()`: Contains definitions for all game items (type, properties).
    *   `MapInfo()`: Contains map-level properties (restrictions like Newbie/NoPK, map names, transport networks).
    *   `QuestList()`: Contains definitions for quests.
    *   `m_NameIndex As Dictionary`: Internal cache for quick username to UserIndex lookups.
    *   `TotalMapasCiudades()`, `MapasInterdimensionales()`, `MapasEventos()`: Arrays storing lists of special map numbers.
*   **Modules (Implicit or Explicit):**
    *   **Networking System:** Extensive use of functions like `SendData`, `WriteConsoleMsg`, `WriteLocaleMsg`, `PrepareMessageChatOverHead`, `PrepareMessagePlayWave`, `PrepareMessageCreateFX`, `WriteUpdateHP`, `WriteUpdateUserStats`, etc. (likely from `Protocol.bas` or a similar networking module).
    *   **Database/File I/O:**
        *   `GetVar` (for `Help.dat`).
        *   `clsIniManager` (for reading `MapasEspeciales.dat`).
        *   `Execute` (for direct SQL database updates in `resetPj`).
    *   **UI/Forms:** `frmMain.List1` for direct output to the server's main console.
    *   **Character Management:** Functions like `DarCuerpo`, `VaciarInventario`, `RellenarInventario`, `ResetUserSpells`, `UpdateUserHechizos`, `ResetFacciones`, `WarpUserChar`, `RevivirUsuario` (likely from modules managing character creation, inventory, spells, etc.).
    *   **NPC Management/AI:** `AI.SetMovement` for changing NPC behavior.
    *   **Guild System:** `modGuilds.GuildName` to get guild names.
    *   **Map System:** `ModMap.ActivateTrap` to trigger map traps.
    *   **Commerce System:** `FinComerciarUsu` to end secure trade if a player is moved.
    *   **Error Logging:** `TraceError`, `LogError`.
*   **Global Constants & Types:**
    *   Numerous constants defining game limits (`LimiteNewbie`), alignments (`Ciudadano`), vision ranges (`RANGO_VISION_X`), map boundaries (`MinXBorder`), item/NPC/player types (`e_OBJType`, `e_NPCType`, `e_PlayerType`), factions (`e_Facciones`), skills (`e_Skill`), attributes (`e_Atributos`), triggers (`e_Trigger`), and various status flags.

This module acts as a foundational layer providing many general-purpose game mechanics and checks.Okay, I've created the Markdown documentation. Now I'll create the Mermaid component diagram in `Documentation/GameLogic.mmd`.

This module is quite central, so it will interact with many components.
