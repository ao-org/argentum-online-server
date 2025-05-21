# API Module Documentation

## 1. General Purpose

The `API.bas` module serves as an interface layer for communication between the Argentum 20 game server and an external Application Programming Interface (API) or service. Its primary roles are:

1.  **Data Serialization and Transmission:** To gather detailed game data (specifically player character information) from the server's internal structures, serialize it into JSON format, and transmit it to the external API.
2.  **Incoming Command Handling:** To receive commands (also in JSON format) from the external API, parse them, and trigger corresponding actions or processes within the game server.
3.  **Connection Management (Basic):** It includes a basic mechanism to queue outgoing data if the connection to the external API is temporarily unavailable, attempting to resend it later.

This module facilitates integration with external systems, potentially for services like website user profiles, data analytics, or external game management tools.

## 2. Main Public Subroutines and Functions

*   **`HandleIncomingAPIData(data As String)`**:
    *   Processes incoming data strings received from the external API.
    *   It parses the `data` string as JSON (using `mod_JSON.parse`).
    *   It extracts an `action` command from the JSON header.
    *   Currently, it includes placeholder `Select Case` statements for commands like `"user_load"` and `"recursos_reload"`. Any unrecognized command results in an error log.
    *   If the debug form `frmAPISocket` is visible, it displays the raw incoming data.

*   **`SendDataAPI(data As String)`**:
    *   Sends the provided `data` string (expected to be JSON) to the external API.
    *   It uses the socket object located on the `frmAPISocket` form (`frmAPISocket.Socket.SendData`).
    *   If the socket is not currently connected (`sckConnected`), the data is pushed into a queue (`packetResend`) for a later attempt at sending.

*   **`SaveUserAPI(UserIndex As Integer, Optional Logout As Boolean)`**:
    *   Gathers comprehensive data for a specific player character (`UserIndex`).
    *   It constructs a JSON object with a "header" (action: "user_save", expectsResponse: false) and a "body".
    *   The "body" contains detailed information about the user, broken down into several sections, each fetched by functions from the `API_User` module/class:
        *   `user`: Principal character data (from `API_User.Principal`).
        *   `attribute`: Character attributes (from `API_User.Atributos`).
        *   `spell`: Character's learned spells (from `API_User.Hechizo`).
        *   `inventory_item`: Character's inventory (from `API_User.Inventario`).
        *   `bank_item`: Character's bank inventory (from `API_User.InventarioBanco`).
        *   `skillpoint`: Character's skills (from `API_User.Habilidades`).
        *   `pet`: Character's pets (from `API_User.Mascotas`).
        *   `quest`: Character's active quests (from `API_User.Quest`).
        *   `quest_done`: Character's completed quests (from `API_User.QuestTerminadas`), if any.
    *   The fully constructed JSON string is then sent to the external API using `SendDataAPI`.
    *   If the debug form `frmAPISocket` is visible, it displays the outgoing JSON data.

## 3. Notable Dependencies

The `API.bas` module interacts with several other components and systems:

*   **`frmAPISocket` (Form):**
    *   This form manages the underlying Winsock TCP/IP socket connection used to communicate with the external API.
    *   It's used by `SendDataAPI` to send data and check connection status.
    *   It also contains UI elements (`txtResponse`, `txtSend`) that can be made visible for debugging API communication.

*   **`mod_JSON` (Module/Library):**
    *   Used for parsing incoming JSON strings (`mod_JSON.parse`).
    *   Likely provides the `JS_Object` class (or similar) used for constructing the JSON objects that are sent to the API.

*   **`API_User` (Module/Class):**
    *   This is a crucial helper component that `SaveUserAPI` relies on extensively.
    *   It encapsulates the logic for retrieving and formatting various aspects of a user's data (e.g., `API_User.Principal`, `API_User.Atributos`, `API_User.Inventario`). It prepares these data segments for inclusion in the final JSON payload.

*   **`UserList()` (Global Data Structure):**
    *   The global array or collection holding all online player data. `SaveUserAPI` accesses `UserList(UserIndex)` to retrieve the raw data that `API_User` then processes.

*   **`CColaArray` (Class):**
    *   The `packetResend` queue, used to store data that couldn't be sent due to a disconnected socket, is an instance of this class. This is likely a custom queue implementation within the project.

*   **Error Logging System:**
    *   `RegistrarError`: Function used to log errors encountered during API operations.

This module acts as a dedicated interface for structured data exchange with an external system, primarily focused on detailed player data synchronization.
