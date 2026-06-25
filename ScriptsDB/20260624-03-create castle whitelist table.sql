CREATE TABLE IF NOT EXISTS "castle_whitelist" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"character_name" varchar(50) NOT NULL UNIQUE,
	"castle_id" integer NOT NULL UNIQUE,
	FOREIGN KEY (castle_id) REFERENCES castle(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX idx_castle_id ON castle_whitelist(castle_id);