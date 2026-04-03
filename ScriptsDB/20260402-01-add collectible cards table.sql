CREATE TABLE IF NOT EXISTS "account_collectible_cards" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "account_id" integer NOT NULL,
    "card_id" integer NOT NULL,
    "rarity" integer NOT NULL,
    "timestamp" timestamp NOT NULL DEFAULT current_timestamp,
    FOREIGN KEY (account_id) REFERENCES account(id) ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX idx_account ON account_collectible_cards(account_id);