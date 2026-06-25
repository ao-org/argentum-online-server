CREATE TABLE IF NOT EXISTS "castle" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"trigger" integer NOT NULL UNIQUE,
	"owner_account_id"  integer NOT NULL UNIQUE,
	"obj_id" integer NOT NULL,
	"white_list" varchar(400) DEFAULT NULL,
	"foundation_date" timestamp DEFAULT NULL,
	"end_date" timestamp DEFAULT NULL,
	"is_active" integer DEFAULT 1,
	FOREIGN KEY (owner_account_id) REFERENCES user(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX idx_owner_account_id ON castle(owner_account_id);