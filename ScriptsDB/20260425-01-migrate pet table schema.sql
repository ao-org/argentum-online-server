-- Backup existing pet data
CREATE TABLE IF NOT EXISTS "pet_backup" AS SELECT * FROM "pet";

-- Drop old pet table
DROP TABLE IF EXISTS "pet";

-- Create new optimized pet table
CREATE TABLE IF NOT EXISTS "pet" (
	"user_id"	integer NOT NULL,
	"pet_id1"	integer DEFAULT NULL,
	"pet_id2"	integer DEFAULT NULL,
	"pet_id3"	integer DEFAULT NULL,
	PRIMARY KEY("user_id"),
	CONSTRAINT "fk_pet_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- Migrate existing data from backup to new structure
INSERT INTO "pet" ("user_id", "pet_id1", "pet_id2", "pet_id3")
SELECT
    "user_id",
    MAX(CASE WHEN "number" = 1 THEN "pet_id" ELSE NULL END),
    MAX(CASE WHEN "number" = 2 THEN "pet_id" ELSE NULL END),
    MAX(CASE WHEN "number" = 3 THEN "pet_id" ELSE NULL END)
FROM "pet_backup"
GROUP BY "user_id";

-- Clean up backup table
DROP TABLE IF EXISTS "pet_backup";
