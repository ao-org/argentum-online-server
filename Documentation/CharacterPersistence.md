# CharacterPersistence Module Documentation

## 1. General Purpose

The `CharacterPersistence.bas` module is responsible for the critical task of managing the saving (persisting) and loading of player character data to and from the game's database. It acts as the bridge between the server's in-memory representation of a player (`UserList(UserIndex) As t_User`) and the backend database where all permanent character information is stored.

This module handles a comprehensive set of character data, including:

*   Basic character attributes (name, level, experience, class, race, gender, appearance).
*   Stats (HP, mana, stamina, hunger, thirst, gold, bank gold).
*   Positional data (current map and coordinates).
*   Flags indicating various states (dead, poisoned, sailing, GM privileges, etc.).
*   Inventory items (both main inventory and bank inventory), including equipped status.
*   Learned spells.
*   Owned pets.
*   Acquired skills and skill points.
*   Active and completed quests.
*   Faction information and progress.
*   Patron tier status and associated benefits (like house keys, extra inventory slots, character slots).

It ensures that when a player logs in, their character is accurately restored to its last saved state, and when they log out or the server performs a save operation, all changes are written back to the database.

## 2. Main Public Subroutines and Functions

Here's a list of the primary public subroutines and functions provided by the `CharacterPersistence` module:

*   **`GetCharacterName(UserId As Long) As String`**: Retrieves a character's name from the database given their unique `UserId`.
*   **`LoadCharacterBank(UserIndex As Integer) As Boolean`**: Loads the items stored in the character's bank from the database and populates the `UserList(UserIndex).BancoInvent` structure.
*   **`get_num_inv_slots_from_tier(t As e_TipoUsuario) As Integer`**: Returns the total number of inventory slots a character is entitled to based on their patron tier (`e_TipoUsuario`).
*   **`LoadCharacterInventory(UserIndex As Integer) As Boolean`**: Loads the character's main inventory items from the database into `UserList(UserIndex).Invent`. It also handles equipping items that were saved as equipped.
*   **`LoadCharacterFromDB(userIndex As Integer) As Boolean`**: This is the primary function for loading a character's complete data set from the database when a player logs in. It fetches and populates all aspects of the `UserList(userIndex)` structure, including basic info, stats, flags, inventory, bank, spells, skills, pets, quests, and faction details. It also performs checks for bans or account locks and loads patron-specific data like house keys.
*   **`MaxCharacterForTier(tier As e_TipoUsuario) As Integer`**: Determines the maximum number of character slots an account is allowed based on its patron tier.
*   **`GetPatronTierFromAccountID(account_id As Long) As e_TipoUsuario`**: Queries the database to find the patron tier (e.g., Normal, Aventurero, Heroe, Leyenda) associated with a given account ID.
*   **`LoadPatronCreditsFromDB(UserIndex As Integer)`**: Loads any available patron credits for the user's account from the database.
*   **`SaveCharacterDB(userIndex As Integer)`**: The main function responsible for saving an existing character's current state from the `UserList(userIndex)` structure back to the database. It updates all relevant tables for the character's attributes, inventory, bank, spells, skills, pets, and quests using a series of SQL `REPLACE INTO` or `UPDATE` statements. This is typically called when a player logs out or during periodic server saves.
*   **`SaveNewCharacterDB(userIndex As Integer)`**: Handles the initial saving of a newly created character to the database. It first inserts the basic character data to obtain a new character ID (using `last_insert_rowid()`) and then saves all associated data (inventory, spells, skills, etc.) similar to `SaveCharacterDB`.

## 3. Notable Dependencies

The `CharacterPersistence` module is heavily reliant on database interactions and several other systems/data structures:

*   **Global Data Structures:**
    *   `UserList()`: The global array/collection holding the in-memory `t_User` structures for all online players. This module reads from and writes to these structures.
    *   `ObjData()`: Array containing definitions for all game items, used to validate item data during loading.
    *   `QuestList()`: Array containing definitions for quests, used when loading quest progress.
*   **Database System (Implicit, likely `Database.bas` or similar):**
    *   `Query(SQL As String, Optional Params As Variant) As ADODB.Recordset`: A function that executes a given SQL SELECT query (often with parameters) and returns an ADODB Recordset.
    *   `Execute(SQL As String, Optional Params As Variant)`: A subroutine that executes a given SQL action query (INSERT, UPDATE, REPLACE, DELETE).
    *   `ADODB.Recordset`: Object used to work with the results of database queries.
    *   `clsSQLBuilder` (aliased as `Builder`): A class used to dynamically construct complex SQL queries, particularly for batch operations.
*   **Database Query Definitions (Implicit, likely `Database_Queries.bas` or a constants module):**
    *   Constants holding predefined SQL query strings (e.g., `QUERY_LOAD_MAINPJ`, `QUERY_UPDATE_MAINPJ`, `QUERY_UPSERT_SPELLS`, `QUERY_SAVE_INVENTORY`).
*   **Error Logging Module (Implicit):**
    *   `LogDatabaseError(Message As String)`: Function to log database-specific errors.
*   **Networking System (Implicit, likely `Protocol.bas` or `Network.bas`):**
    *   `CloseSocket(UserIndex As Integer)`: To disconnect a user (e.g., if banned or account locked).
    *   `WriteShowMessageBox(UserIndex As Integer, MessageId As Integer, Optional ReplacementString As String)`: To send system messages to the client.
*   **Character Utility Functions (Implicit, from various modules):**
    *   `EquiparInvItem(UserIndex As Integer, Slot As Integer)`: To equip an item loaded from the database.
    *   `RegisterUserName(UserId As Long, UserName As String)`: Likely to update a global name cache or similar structure.
    *   `SanitizeNullValue(Value As Variant, DefaultValue As Variant) As Variant`: Utility to handle potential Null values from database fields.
*   **Performance Monitoring (Implicit):**
    *   `PerformanceTestStart(ByRef TimerVar As Long)`
    *   `PerformTimeLimitCheck(TimerVar As Long, TaskName As String, ThresholdMS As Long)`
*   **Global Constants & Types:**
    *   `MAX_USERINVENTORY_SLOTS`, `MAX_BANCOINVENTORY_SLOTS`, `MAXUSERHECHIZOS`, `NUMSKILLS`, `MAXMASCOTAS`, `MAXUSERQUESTS`: Define limits for various character attributes.
    *   `e_TipoUsuario`: Enum for patron tiers.
    *   `t_User`: The core data type defining a player character's structure.

This module is fundamental for maintaining player progression and data integrity in the game.Okay, I've created the Markdown documentation. Now I'll create the Mermaid component diagram in `Documentation/CharacterPersistence.mmd`.

This module's primary interaction is with the database system and the in-memory `UserList` data structure.
