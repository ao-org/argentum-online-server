CREATE TABLE IF NOT EXISTS "castle" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"trigger" integer NOT NULL,
	"owner_account_id"  integer NULL UNIQUE,
	"owner_character_id" integer NULL UNIQUE,
	"spawner_obj_id" integer NOT NULL UNIQUE,
	"portal_obj_id" integer NOT NULL UNIQUE,
	"foundation_date" timestamp DEFAULT NULL,
	"end_date" timestamp DEFAULT NULL,
	"is_active" integer DEFAULT 0,

	"outside_map" integer DEFAULT 0,
	"outside_x" integer DEFAULT 0,
	"outside_y" integer DEFAULT 0,
	"inside_map" integer NOT NULL UNIQUE,
	"inside_x" integer NOT NULL,
	"inside_y" integer NOT NULL,
	FOREIGN KEY (owner_account_id) REFERENCES user(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX idx_owner_account_id ON castle(owner_account_id);