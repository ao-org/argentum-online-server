# Comercio Module (modSistemaComercio) Documentation

## 1. General Purpose

The `Comercio.bas` module (internally named `modSistemaComercio`) is responsible for managing all player-NPC (Non-Player Character) trading interactions within the Argentum 20 game server. It handles the mechanics of players buying items from NPCs and selling items to NPCs. This includes validating trade conditions, calculating prices based on item values and player skills, updating player and NPC inventories, managing player gold, and triggering relevant skill increases.

The module differentiates between buying (`eModoComercio.Compra`) and selling (`eModoComercio.Venta`) operations, applying specific rules and calculations for each.

## 2. Main Public Subroutines and Functions

Here's a list of the primary public subroutines and functions provided by the `Comercio` module:

*   **`Comercio(Modo As eModoComercio, UserIndex As Integer, NpcIndex As Integer, Slot As Integer, Cantidad As Integer)`**:
    This is the core function for executing a trade action.
    *   **If `Modo = eModoComercio.Compra` (Buying):**
        *   The player (`UserIndex`) attempts to buy `Cantidad` of the item located in the NPC's (`NpcIndex`) inventory `Slot`.
        *   It checks if the player has enough gold and inventory space.
        *   Calculates the price, applying a discount based on the player's "Comerciar" (Trading) skill using the private `Descuento` function. Prices are rounded up (`Ceil`).
        *   If successful, the item is added to the player's inventory, gold is deducted, and the item is removed from the NPC's inventory.
        *   Includes an anti-cheat mechanism to ban players attempting to buy excessively large quantities.
        *   Special handling for key items: logs the sale and updates the NPC's persistent data (`NPCs.dat`) to prevent re-selling of unique keys.
    *   **If `Modo = eModoComercio.Venta` (Selling):**
        *   The player (`UserIndex`) attempts to sell `Cantidad` of an item from their own inventory `Slot` to the NPC (`NpcIndex`).
        *   It verifies that the item is sellable (not newbie, not marked as "destruye" or "instransferible") and that the NPC is configured to buy that type of item (or if the NPC already sells that specific item).
        *   Checks if the player has the necessary privileges to sell items (GMs might be restricted).
        *   Calculates the sale price using `SalePrice` function (typically item value / `REDUCTOR_PRECIOVENTA`), rounded down (`Fix`).
        *   If successful, the item is removed from the player's inventory, gold is added (up to `MAXORO`), and the item is added to the NPC's inventory using `SlotEnNPCInv` to find a suitable slot.
        *   The player's "Comerciar" (Trading) skill is potentially increased via `SubirSkill`.

*   **`IniciarComercioNPC(UserIndex As Integer)`**:
    *   Called when a player initiates a trade session with an NPC.
    *   It validates the targeted NPC reference.
    *   Sends the NPC's current inventory to the player (using `UpdateNpcInv`), including prices adjusted for the player's trading skill.
    *   Plays a sound effect if the NPC has one configured for opening trade.
    *   Sets the player's `Comerciando` flag to `True`.
    *   Sends a "CommerceInit" message to the client to open the trade window.

*   **`UpdateNpcInvToAll(UpdateAll As Boolean, NpcIndex As Integer, Slot As Byte)`**:
    *   Updates the inventory view of a specific NPC (`NpcIndex`) for all players who are currently trading with that NPC.
    *   If `UpdateAll` is true, all slots are refreshed. Otherwise, only the specified `Slot` is updated.
    *   This ensures that changes to an NPC's inventory (e.g., an item being bought) are reflected in real-time for all interacting players.

*   **`SalePrice(ObjIndex As Integer) As Single`**:
    *   Calculates and returns the price an NPC will pay for a given item (`ObjIndex`).
    *   The price is typically the item's base value (`ObjData(ObjIndex).Valor`) divided by the `REDUCTOR_PRECIOVENTA` constant.
    *   Items marked as "Newbie" cannot be sold and will result in a price of 0.

## 3. Notable Dependencies

The `Comercio` module interacts with several other modules and global data structures:

*   **Global Data Structures:**
    *   `UserList()`: Array/collection holding all online user data, including their inventory (`Invent`), stats (`Stats.GLD`, `Stats.UserSkills`), and flags (`flags.Comerciando`, `flags.TargetNPC`).
    *   `NpcList()`: Array/collection holding NPC data, including their inventory (`Invent`), configuration (`TipoItems`, `Numero`), and sound events (`SoundOpen`).
    *   `ObjData()`: Array/collection containing definitions for all game items, such as their value (`Valor`), type (`OBJType`), and properties (`Newbie`, `Crucial`, `Destruye`, `Instransferible`).
*   **Modules (Implicit or Explicit):**
    *   **Inventory Management System:**
        *   `MeterItemEnInventario(UserIndex, Objeto)`: Adds an item to a user's inventory.
        *   `QuitarUserInvItem(UserIndex, Slot, Cantidad)`: Removes items from a user's inventory slot.
        *   `QuitarNpcInvItem(NpcIndex, Slot, Cantidad)`: Removes items from an NPC's inventory slot.
        *   `UpdateUserInv(UpdateAll, UserIndex, Slot)`: Sends updates about a user's inventory to the client.
        (These are likely located in a module like `InvUsuario.bas` or a general inventory handling module).
    *   **Networking/Protocol System:**
        *   `SendData(Target, UserIndex, Message)`: Core function to send data packets.
        *   `WriteLocaleMsg`, `WriteUpdateGold`, `WriteChangeNPCInventorySlot`, `WriteCommerceInit`, `WritePlayWave`, `WriteShowMessageBox`: Specific functions to prepare and send formatted messages to clients. (Likely from `Protocol.bas` or similar).
    *   **Skill System:**
        *   `SubirSkill(UserIndex, e_Skill.Comerciar)`: Function to grant experience or level up a skill. (Likely from a `Skills.bas` or `SistemaSkills.bas` module).
    *   **Data Persistence/Logging:**
        *   `WriteVar(FilePath, Section, Key, Value)`: Used to write to `.dat` files, specifically `NPCs.dat` for key sales.
        *   `logVentaCasa(Message As String)`: Logs the sale of house keys.
        (Suggests a utility module for file I/O or a more specific data logging module).
    *   **NPC Data System:**
        *   `NpcSellsItem(NpcNumero, ObjIndex)`: Checks if an NPC type is configured to sell a particular item.
    *   **Error Handling & Admin:**
        *   `TraceError(ErrNum, ErrDesc, Location, Line)`: Logs errors.
        *   `Ban(UserName, Reason, AdminName)`: Bans a user.
        *   `CloseSocket(UserIndex)`: Disconnects a user.
*   **Global Constants & Types:**
    *   `eModoComercio`: Enum defining Buy/Sell modes.
    *   `REDUCTOR_PRECIOVENTA`: Constant determining how much less NPCs pay for items compared to their base value.
    *   `MAX_INVENTORY_SLOTS`, `MAX_INVENTORY_OBJS`, `MAXORO`: Constants defining game limits.
    *   `iORO`: Constant for the gold item index.
    *   `DatPath`: Path to data files.
    *   `t_Obj`, `t_UserOBJ`: Types defining item structures.
    *   `e_Skill`, `e_OBJType`, `e_PlayerType`, `e_FontTypeNames`: Various enumerated types.

This module is central to the in-game economy involving NPC merchants.Okay, I've created the Markdown documentation. Now I'll create the Mermaid component diagram in `Documentation/Comercio.mmd`.

Based on the analysis, the `Comercio` module interacts significantly with Inventory Management, User/NPC Data, Item Data, Networking, and potentially Skill and Data Persistence systems.
