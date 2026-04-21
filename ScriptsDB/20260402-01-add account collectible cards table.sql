CREATE TABLE IF NOT EXISTS "account_collectible_cards" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "account_id" integer NOT NULL,
    "card_id" integer NOT NULL,
    "first_added" timestamp NOT NULL DEFAULT current_timestamp,
    "last_updated" timestamp NOT NULL DEFAULT current_timestamp,
    "quantity" integer NOT NULL DEFAULT 1,
    FOREIGN KEY (account_id) REFERENCES account(id) ON DELETE CASCADE ON UPDATE CASCADE
    FOREIGN KEY (card_id) REFERENCES collectible_cards(id) ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX idx_account ON account_collectible_cards(account_id);
CREATE INDEX idx_card ON account_collectible_cards(card_id);