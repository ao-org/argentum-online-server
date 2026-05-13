-- Drop the legacy all_housekeys table.
-- This table is no longer used by the game server or the gameapi.
-- All house data is now stored in the house_key table with account_id-based ownership.
-- The gameapi migration (consolidate-house-keys.js) must be run BEFORE this migration.

DROP TABLE IF EXISTS all_housekeys;
DROP TABLE IF EXISTS all_housekeys_backup;
