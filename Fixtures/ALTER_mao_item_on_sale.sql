CREATE TABLE IF NOT EXISTS "mao_item_on_sale_new" (
    "id"    INTEGER NOT NULL,
    "created_at"    timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at"    timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "deleted_at"    timestamp DEFAULT NULL,
    "item_id"    INTEGER NOT NULL,
    "price_in_tokens"    INTEGER NOT NULL DEFAULT 1,
    "item_quantity"    INTEGER NOT NULL DEFAULT 1,
    PRIMARY KEY("id" AUTOINCREMENT)
);

INSERT INTO mao_item_on_sale_new (created_at,updated_at,deleted_at,item_id,price_in_tokens,item_quantity)
SELECT created_at,updated_at,deleted_at,item_id,price_in_tokens,item_quantity FROM mao_item_on_sale;

DROP TABLE mao_item_on_sale;
ALTER TABLE mao_item_on_sale_new RENAME TO mao_item_on_sale;



