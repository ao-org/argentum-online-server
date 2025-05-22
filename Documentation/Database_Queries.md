# Database_Queries Module Documentation

## 1. General Purpose

The `Database_Queries.bas` module serves as a centralized repository for constructing and storing SQL query strings used throughout the Argentum 20 game server. Its primary purpose is to define the exact SQL statements needed to interact with the game's database for operations related to character data. This includes queries for loading existing characters, creating new characters, and saving updates to existing characters.

By centralizing these query definitions, the module promotes:
*   **Maintainability:** SQL queries are located in one place, making them easier to find, understand, and modify if the database schema changes.
*   **Readability:** Other modules that perform database operations can use these clearly named constants instead of having large, embedded SQL strings.
*   **Efficiency:** It utilizes a `cStringBuilder` class (presumably custom) for efficient construction of these query strings, which is particularly beneficial for complex or batch queries.

The module itself does not execute these queries; rather, it makes them available as global string constants to be used by other modules that handle the actual database communication (e.g., `CharacterPersistence.bas` in conjunction with a database access module).

## 2. Main Public Subroutines and Global Query Constants

### Public Subroutine:

*   **`Contruir_Querys()`**:
    *   This is the main public entry point of the module.
    *   When called (typically at server startup), it invokes private subroutines to dynamically build all the SQL query strings and store them in their respective global string constants.

### Global Query String Constants:

The module defines several public string constants that hold the constructed SQL queries. These are grouped by their purpose:

**Loading Character Data:**
*   **`QUERY_LOAD_MAINPJ`**: A comprehensive `SELECT` query to retrieve all core data for a specified character from the `user` table and other related tables through JOINs (implied by the extensive list of fields).

**Creating New Character Data (primarily `INSERT` statements):**
*   **`QUERY_SAVE_MAINPJ`**: An `INSERT` query to create a new record in the `user` table with basic character information.
*   **`QUERY_SAVE_SPELLS`**: An `INSERT` query for batch inserting initial spell data for a new character into the `spell` table.
*   **`QUERY_SAVE_INVENTORY`**: An `INSERT` query for batch inserting initial inventory items for a new character into the `inventory_item` table.
*   **`QUERY_SAVE_SKILLS`**: An `INSERT` query for batch inserting initial skill point allocations for a new character into the `skillpoint` table.
*   **`QUERY_SAVE_QUESTS`**: An `INSERT` query for initializing quest data for a new character in the `quest` table.
*   **`QUERY_SAVE_PETS`**: An `INSERT` query for initializing pet data for a new character in the `pet` table.

**Saving Existing Character Data (primarily `UPDATE` or `REPLACE INTO` statements):**
*   **`QUERY_UPDATE_MAINPJ`**: An `UPDATE` query to save changes to an existing character's core data in the `user` table.
*   **`QUERY_SAVE_BANCOINV`**: A `REPLACE INTO` query for saving the character's bank inventory to the `bank_item` table.
*   **`QUERY_UPSERT_SPELLS`**: A `REPLACE INTO` query for saving the character's current spells to the `spell` table (handles both insert and update).
*   **`QUERY_UPSERT_INVENTORY`**: A `REPLACE INTO` query for saving the character's current inventory to the `inventory_item` table.
*   **`QUERY_UPSERT_SKILLS`**: A `REPLACE INTO` query for saving the character's current skill points to the `skillpoint` table.
*   **`QUERY_UPSERT_PETS`**: A `REPLACE INTO` query for saving the character's current pet information to the `pet` table.

## 3. Private Subroutines

*   **`ConstruirQuery_CargarPersonaje()`**: Responsible for building the `QUERY_LOAD_MAINPJ` string.
*   **`ConstruirQuery_CrearPersonaje()`**: Responsible for building all `QUERY_SAVE_*` strings used when a new character is created.
*   **`ConstruirQuery_GuardarPersonaje()`**: Responsible for building `QUERY_UPDATE_MAINPJ` and all `QUERY_UPSERT_*` (and `QUERY_SAVE_BANCOINV`) strings used when saving an existing character.

## 4. Notable Dependencies

*   **`cStringBuilder` Class:** This module relies on a custom class named `cStringBuilder` for efficient concatenation of string parts to form the final SQL queries. This class is instantiated locally within the construction subroutines.
*   **Database Schema:** The structure of all queries (table names, column names) is directly tied to the underlying database schema of the game. The primary tables referenced are `user`, `spell`, `inventory_item`, `bank_item`, `skillpoint`, `quest`, and `pet`.
*   **Global Game Constants:** The construction of batch queries (e.g., for spells, inventory items) uses global constants that define the maximum number of these items (e.g., `MAXUSERHECHIZOS`, `MAX_INVENTORY_SLOTS`, `NUMSKILLS`, `MAXMASCOTAS`, `MAX_BANCOINVENTORY_SLOTS`). These constants are expected to be defined elsewhere.
*   **Modules Utilizing These Queries:**
    *   `CharacterPersistence.bas`: This is the primary consumer of the query strings defined in this module. It uses these constants when calling database access functions to load or save character data.
    *   Any other module that might need to perform direct database operations related to the specific data covered by these queries.

This module plays a crucial role in decoupling the raw SQL from the logic that executes database operations, contributing to a cleaner and more organized codebase.
