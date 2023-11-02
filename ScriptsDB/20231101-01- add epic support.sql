CREATE TABLE IF NOT EXISTS "epic_id_mapping" (
	"epic_id"	varchar(64) NOT NULL,
    "user_id" integer NOT NULL,
	"last_login" integer NOT NULL,
	CONSTRAINT "fk_epic_id_mapping" FOREIGN KEY("user_id") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE,
	CREATE UNIQUE INDEX epic_id_mapping_idx ON data(epic_id, user_id);
);