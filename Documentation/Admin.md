# Admin Module Documentation

## 1. General Purpose

The `Admin.bas` module serves as a central hub for server administration and configuration within the Argentum 20 game server. It manages a wide range of functionalities, including:

*   **Server State and Configuration:** Storing and managing global server settings like the Message of the Day (MOTD), current weather conditions (rain, snow), network port, and IP.
*   **Game Mechanics Timing:** Defining various time intervals that govern core game mechanics such as health and stamina regeneration, hunger/thirst rates, poison effects, spell cooldowns, and AI processing rates.
*   **Game Balance Parameters:** Holding critical balance values for experience progression, skill gain difficulty, combat modifiers (critical hit damage, stun times), and resource gathering.
*   **Player Administration:** Providing tools for Game Masters (GMs) and administrators to manage players, including jailing, banning (temporary and permanent checks), unbanning, and checking character existence.
*   **Privilege Management:** Determining and comparing privilege levels of users (e.g., regular user, Counselor, Demigod, God, Admin).
*   **Server Maintenance:** Including functions for periodic world saves and purging jail sentences.
*   **Data Persistence:** Loading and saving administrative data, particularly ban lists, to and from files.

The module also declares many global variables that are accessed by other parts of the server to control their behavior according to these central settings.

## 2. Main Public Subroutines and Functions

Here's a list of the primary public subroutines and functions provided by the `Admin` module:

*   **`ReSpawnOrigPosNpcs()`**: Respawns all NPCs of type "Guardias" (Guards) to their original designated positions on their respective maps.
*   **`WorldSave()`**: Initiates a server-wide world save process. This includes notifying all players, respawning guards, and potentially backing up map data (details depend on `MapInfo(LoopX).backup_mode`).
*   **`PurgarPenas()`**: Iterates through all logged-in users and decrements their `Pena` (jail time) counter. If a user's sentence reaches zero, they are warped to a designated "liberation" spot.
*   **`Encarcelar(UserIndex As Integer, minutos As Long, Optional GmName As String)`**: Sends a specified user to jail for a given number of minutes. Notifies the user of their sentence and who jailed them if `GmName` is provided.
*   **`BANCheck(Name As String) As Boolean`**: Checks if a given character name is currently banned by querying a database (`BANCheckDatabase`).
*   **`PersonajeExiste(Name As String) As Boolean`**: Verifies if a character with the given name exists by checking a database (`GetUserValue`).
*   **`IsValidUserId(UserId As Long) As Boolean`**: Checks if a given numeric User ID is valid by querying the database.
*   **`UnBan(Name As String) As Boolean`**: Removes a ban for the specified character name by updating a database (`UnBanDatabase`) and clearing ban details from a log file (`BanDetail.dat`).
*   **`UserDarPrivilegioLevel(Name As String) As e_PlayerType`**: Determines and returns the administrative privilege level (e.g., Admin, Dios, User) of a character based on their name by checking against lists/functions like `EsAdmin`, `EsDios`, etc.
*   **`BanTemporal(nombre As String, dias As Integer, Causa As String, Baneador As String)`**: Implements a temporary ban for a user. Adds the ban details to an in-memory `Baneos` collection, saves this specific ban to persistent storage (`SaveBan`), and sends a notification to admins.
*   **`SaveBans()`**: Saves the entire list of current temporary bans (stored in the `Baneos` collection) to the `baneos.dat` file.
*   **`SaveBan(num As Integer)`**: Saves a single ban entry (specified by its index in the `Baneos` collection) to `baneos.dat` and also updates the ban status in a database via `SaveBanDatabase`.
*   **`LoadBans()`**: Loads all temporary ban records from the `baneos.dat` file into the in-memory `Baneos` collection when the server starts.
*   **`CompararUserPrivilegios(Personaje_1 As Integer, Personaje_2 As Integer) As Integer`**: Compares the administrative privilege levels of two online users (identified by their `UserIndex`). Returns 1 if P1 > P2, 0 if P1 = P2, -1 if P1 < P2.
*   **`CompararPrivilegiosUser(Personaje_1 As Integer, Personaje_2 As Integer) As Integer`**: Appears to be an identical alias for `CompararUserPrivilegios`.
*   **`CompararPrivilegios(Izquierda As e_PlayerType, Derecha As e_PlayerType) As Integer`**: A core function that compares two `e_PlayerType` privilege values directly. Returns 1 if left > right, 0 if equal, -1 if left < right.

## 3. Notable Dependencies

The `Admin` module interacts with various other parts of the server and relies on several global structures and external resources:

*   **Global Data Structures:**
    *   `UserList()`: Array/collection holding data for all online users. Used for jailing, checking privileges, etc.
    *   `NpcList()`: Array/collection holding NPC data. Used by `ReSpawnOrigPosNpcs`.
    *   `MapInfo()`: Array/collection with map-specific information, used by `WorldSave`.
    *   `AdministratorAccounts As Dictionary`: Stores details about administrator accounts.
    *   `MOTD() As t_Motd`: Stores the Message of the Day.
    *   `Baneos As Collection`: In-memory list of `tBaneo` objects for temporary bans.
    *   `Libertad As t_WorldPos`, `Prision As t_WorldPos`: Coordinates for liberation and prison locations.
*   **External Modules/Functions (Implicit):**
    *   **Database Interaction Layer:** Functions like `BANCheckDatabase`, `GetUserValue`, `UnBanDatabase`, `SaveBanDatabase`, `GetUserValueById` indicate a module or set of functions responsible for database operations (player accounts, bans).
    *   **File I/O System:** Functions like `WriteVar`, `GetVar`, `FileExist` for reading and writing data to `.dat` files (e.g., `baneos.dat`, `BanDetail.dat`). This implies a utility module for structured file access.
    *   **Networking System:** Functions like `SendData`, `PrepareMessageLocaleMsg`, `WriteLocaleMsg` for sending messages and updates to game clients.
    *   **NPC Management System:** `QuitarNPC`, `ReSpawnNpc` used for managing NPC instances.
    *   **Player Action System:** `WarpUserChar` for teleporting players.
    *   **Privilege Checking System:** Functions like `EsGM`, `EsAdmin`, `EsDios`, `EsSemiDios`, `EsConsejero` (which might use `AdministratorAccounts`).
    *   **Logging System:** `TraceError` for reporting errors.
*   **Global Variables & Constants:**
    *   Numerous global variables defined within this module itself store intervals, balance settings, and server status (e.g., `SanaIntervaloSinDescansar`, `PorcentajeRecuperoMana`, `Lloviendo`).
    *   `DatPath`: A global constant or variable specifying the path to data files.
    *   `Guardias`: A constant likely defining an NPC ID or type for guards.
*   **Forms:**
    *   `FrmStat`: A Visual Basic form, likely used to display server statistics or progress during operations like `WorldSave`.

This module is critical for the overall operation, management, and stability of the game server.
