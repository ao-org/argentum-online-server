/* Requires SQLite 3.35 or newer. Keep newline-independent block comments for RunScriptInFile. */
BEGIN IMMEDIATE;

CREATE TABLE castle_legacy_trigger_map (
    legacy_trigger INTEGER PRIMARY KEY CHECK (legacy_trigger BETWEEN 21 AND 40),
    castle_id INTEGER NOT NULL UNIQUE REFERENCES castle(id) ON DELETE CASCADE ON UPDATE CASCADE
);

INSERT INTO castle_legacy_trigger_map (legacy_trigger, castle_id)
SELECT trigger, id FROM castle;

ALTER TABLE castle DROP COLUMN trigger;

/* Replace only the placement columns, preserving row IDs, incoming references and sequences. */
ALTER TABLE castle_coordinates ADD COLUMN migrated_outside_map INTEGER DEFAULT NULL;
ALTER TABLE castle_coordinates ADD COLUMN migrated_outside_x INTEGER DEFAULT NULL;
ALTER TABLE castle_coordinates ADD COLUMN migrated_outside_y INTEGER DEFAULT NULL
    CHECK ((migrated_outside_map IS NULL AND migrated_outside_x IS NULL AND migrated_outside_y IS NULL)
        OR (migrated_outside_map IS NOT NULL AND migrated_outside_x IS NOT NULL AND migrated_outside_y IS NOT NULL
            AND migrated_outside_map > 0 AND migrated_outside_x > 0 AND migrated_outside_y > 0));

UPDATE castle_coordinates SET
    migrated_outside_map = CASE WHEN outside_map = 0 AND outside_x = 0 AND outside_y = 0 THEN NULL ELSE outside_map END,
    migrated_outside_x = CASE WHEN outside_map = 0 AND outside_x = 0 AND outside_y = 0 THEN NULL ELSE outside_x END,
    migrated_outside_y = CASE WHEN outside_map = 0 AND outside_x = 0 AND outside_y = 0 THEN NULL ELSE outside_y END;

ALTER TABLE castle_coordinates DROP COLUMN outside_map;
ALTER TABLE castle_coordinates DROP COLUMN outside_x;
ALTER TABLE castle_coordinates DROP COLUMN outside_y;
ALTER TABLE castle_coordinates RENAME COLUMN migrated_outside_map TO outside_map;
ALTER TABLE castle_coordinates RENAME COLUMN migrated_outside_x TO outside_x;
ALTER TABLE castle_coordinates RENAME COLUMN migrated_outside_y TO outside_y;

COMMIT;
