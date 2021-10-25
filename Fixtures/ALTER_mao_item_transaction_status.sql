CREATE TABLE IF NOT EXISTS "mao_item_transaction_status" (
	"user_id" INTEGER NOT NULL,
	"eth_transaction_id" TEXT NULL,
	"status" TEXT NOT NULL,
	"created_at" DATETIME NULL,
	"updated_at" DATETIME NULL,
	"deleted_at" DATETIME NULL,
	"item_id" INTEGER NOT NULL,
	"item_quantity" INTEGER NOT NULL
);


INSERT INTO mao_item_on_transaction_status_new (user_id, eth_transaction_id, status, created_at, updated_at, deleted_at, item_id)
SELECT user_id, eth_transaction_id, status, created_at, updated_at, deleted_at, item_id FROM mao_item_on_transaction_status;

DROP TABLE mao_item_on_transaction_status;
ALTER TABLE mao_item_on_transaction_status_new RENAME TO mao_item_on_transaction_status;



