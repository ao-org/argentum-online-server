CREATE TABLE IF NOT EXISTS "account_collectible_cards" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "account_id" integer NOT NULL,
    "card_bit_array" blob NOT NULL,
    FOREIGN KEY (account_id) REFERENCES account(id) ON DELETE CASCADE ON UPDATE CASCADE,
    UNIQUE (account_id)
);
CREATE INDEX idx_account ON account_collectible_cards(account_id);
