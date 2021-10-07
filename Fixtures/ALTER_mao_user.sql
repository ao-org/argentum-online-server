CREATE TABLE IF NOT EXISTS "mao_user_new" (
	"id"	INTEGER NOT NULL,
	"user_id"	INTEGER NOT NULL,
	"eth_transaction_id"	TEXT,
    "created_at"    timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at"    timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
	"deleted_at"	datetime,
    "price_in_tokens"    INTEGER NOT NULL DEFAULT 1,
	"status"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT)
);

INSERT INTO mao_user_new (user_id, eth_transaction_id, created_at, updated_at, deleted_at, status) 
SELECT user_id, eth_transaction_id, created_at, updated_at, deleted_at, status FROM mao_user;

DROP TABLE mao_user;
ALTER TABLE mao_user_new RENAME TO mao_user;


