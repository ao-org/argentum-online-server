CREATE TABLE IF NOT EXISTS "account_collectible_cards" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "account_id" integer NOT NULL,
    "card_bit_array" blob NOT NULL,
);
CREATE INDEX idx_account ON account_collectible_cards(account_id);