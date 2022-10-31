CREATE TABLE patreon_shop_audit (
	id INTEGER PRIMARY KEY,
	acc_id integer NOT NULL,
	char_id integer NOT NULL,
	item_id integer NOT NULL,
	price integer NOT NULL,
	credit_left integer NOT NULL,
	time integer not null
);
