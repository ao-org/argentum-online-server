# Project Modules Overview

This document provides an overview and links to the detailed documentation for various modules within the Argentum 20 game server project. Each module's documentation includes a general description, a list of its main public functions, and its notable dependencies. Component diagrams are also provided to illustrate their interactions.

## Documented Modules

Below is a list of modules for which documentation has been generated:

-   **AI_NPC**: Handles Artificial Intelligence for Non-Player Characters.
    -   [Detailed Documentation](./AI_NPC.md)
    -   [Component Diagram](./AI_NPC.mmd)

-   **AO20CryptoSysWrapper**: Wrapper for cryptographic operations, likely interfacing with an external library like CryptoSys.
    -   [Detailed Documentation](./AO20CryptoSysWrapper.md)
    -   [Component Diagram](./AO20CryptoSysWrapper.mmd)

-   **Admin**: Manages server administration, global settings, game balance parameters, player administration (banning, jailing), and server maintenance tasks.
    -   [Detailed Documentation](./Admin.md)
    -   [Component Diagram](./Admin.mmd)

-   **API**: Facilitates communication between the game server and an external API or service, handling incoming commands and sending out game data (e.g., player information) in JSON format.
    -   [Detailed Documentation](./API.md)
    -   (Component Diagram for API is not available)

-   **CharacterPersistence**: Responsible for saving and loading player character data to and from the game's database.
    -   [Detailed Documentation](./CharacterPersistence.md)
    -   [Component Diagram](./CharacterPersistence.mmd)

-   **Comercio**: Manages the trading system between players (users) and NPCs, including buying and selling items.
    -   [Detailed Documentation](./Comercio.md)
    -   [Component Diagram](./Comercio.mmd)

-   **Database_Queries**: Constructs and stores global string constants that represent SQL queries used for database interactions, primarily for character data.
    -   [Detailed Documentation](./Database_Queries.md)
    -   [Component Diagram](./Database_Queries.mmd)

-   **EffectsOverTime**: Manages dynamic status effects, buffs, debuffs, and other game mechanics that have a duration or trigger under specific conditions.
    -   [Detailed Documentation](./EffectsOverTime.md)
    -   [Component Diagram](./EffectsOverTime.mmd)

-   **GameLogic (Extra)**: A collection of various utility functions and core game logic routines, including positional logic, character status checks, tile event processing, and player information display.
    -   [Detailed Documentation](./GameLogic.md)
    -   [Component Diagram](./GameLogic.mmd)
