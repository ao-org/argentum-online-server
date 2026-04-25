-- Create new optimized pet table
CREATE TABLE IF NOT EXISTS "pet_new" (
	"user_id"	integer NOT NULL,
	"pet_id1"	integer DEFAULT 0,
	"pet_id2"	integer DEFAULT 0,
	"pet_id3"	integer DEFAULT 0,
	PRIMARY KEY("user_id"),
	CONSTRAINT "fk_pet_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- Migrate existing data from old structure to new structure
INSERT INTO "pet_new" ("user_id", "pet_id1", "pet_id2", "pet_id3")
SELECT
    "user_id",
    MAX(CASE WHEN "number" = 1 THEN "pet_id" ELSE 0 END),
    MAX(CASE WHEN "number" = 2 THEN "pet_id" ELSE 0 END),
    MAX(CASE WHEN "number" = 3 THEN "pet_id" ELSE 0 END)
FROM "pet"
GROUP BY "user_id";

-- Drop old pet table and rename new one
DROP TABLE IF EXISTS "pet";
ALTER TABLE "pet_new" RENAME TO "pet";

-- Create index on user_id
CREATE INDEX "pet_index" ON "pet" ("user_id");
