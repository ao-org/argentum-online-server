-- Create new optimized pet table
CREATE TABLE IF NOT EXISTS "pet" (
	"user_id"	integer NOT NULL,
	"pet_id1"	integer DEFAULT NULL,
	"pet_id2"	integer DEFAULT NULL,
	"pet_id3"	integer DEFAULT NULL,
	PRIMARY KEY("user_id"),
	CONSTRAINT "fk_pet_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
-- Create index on user_id
CREATE INDEX "pet_index" ON "pet" ("user_id");
