CREATE TABLE IF NOT EXISTS "collectible_cards" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "description" text NOT NULL,
    "rarity" integer NOT NULL,
    "tags" integer NOT NULL
);