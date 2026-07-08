DROP TABLE "castle_whitelist";

CREATE TABLE "castle_whitelist" (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "character_name" VARCHAR(50) NOT NULL,
    "castle_id" INTEGER NOT NULL,
    FOREIGN KEY (castle_id) REFERENCES castle(id) ON DELETE CASCADE ON UPDATE CASCADE,
    UNIQUE (character_name, castle_id)
);
