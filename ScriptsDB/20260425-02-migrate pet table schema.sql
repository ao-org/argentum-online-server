CREATE TABLE IF NOT EXISTS "pet" (
	user_id	integer PRIMARY KEY NOT NULL,
	pet_id1	integer DEFAULT 0,
	pet_id2	integer DEFAULT 0,
	pet_id3	integer DEFAULT 0,
	CONSTRAINT "fk_pet_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX "pet_index" ON "pet" ("user_id");
