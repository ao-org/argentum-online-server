CREATE TABLE IF NOT EXISTS "mao_items_on_sale" (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "user_id" integer NOT NULL,
    "account_id" integer NOT NULL,
    "item_id" integer NOT NULL,
    "item_qty" integer NOT NULL DEFAULT 1,
    "price_in_fiat" integer NOT NULL DEFAULT 0,
    "date_created" integer NOT NULL DEFAULT (strftime('%s', 'now')),
    "date_updated" integer NOT NULL DEFAULT (strftime('%s', 'now')));