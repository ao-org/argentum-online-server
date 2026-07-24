DROP table IF EXISTS castle;

CREATE TABLE "castle" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"trigger" integer NOT NULL,
	"owner_account_id"  integer NULL UNIQUE,
	"owner_character_id" integer NULL UNIQUE,
	"spawner_obj_id" integer NOT NULL UNIQUE,
	"inside_key_obj_id" integer NOT NULL UNIQUE,
	"foundation_date" timestamp DEFAULT NULL,
	"is_active" integer DEFAULT 0,
	"name" varchar(255) DEFAULT NULL,
	FOREIGN KEY (owner_account_id) REFERENCES account(id) ON DELETE CASCADE ON UPDATE CASCADE,
	FOREIGN KEY (owner_character_id) REFERENCES user(id) ON DELETE CASCADE ON UPDATE CASCADE
);