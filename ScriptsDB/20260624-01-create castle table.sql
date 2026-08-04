CREATE TABLE IF NOT EXISTS "castle" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"trigger" integer NOT NULL,
	"owner_account_id"  integer NULL UNIQUE,
	"owner_character_id" integer NULL UNIQUE,
	"spawner_obj_id" integer NOT NULL UNIQUE,
	"portal_obj_id" integer NOT NULL UNIQUE,
	"foundation_date" timestamp DEFAULT NULL,
	"is_active" integer DEFAULT 0,
	FOREIGN KEY (owner_account_id) REFERENCES user(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX idx_castle_owner_account_id ON castle(owner_account_id);