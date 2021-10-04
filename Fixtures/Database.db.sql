BEGIN TRANSACTION;
CREATE TABLE IF NOT EXISTS "account" (
	"id"	integer NOT NULL,
	"email"	varchar(320) NOT NULL,
	"password"	char(64) NOT NULL,
	"salt"	char(32) NOT NULL,
	"date_created"	timestamp NOT NULL DEFAULT current_timestamp,
	"deleted"	integer DEFAULT '0',
	"validated"	integer NOT NULL DEFAULT '0',
	"validate_code"	char(32) NOT NULL,
	"recovery_code"	varchar(32) NOT NULL DEFAULT '',
	"is_banned"	integer DEFAULT '0',
	"banned_by"	varchar(30) NOT NULL DEFAULT '',
	"ban_reason"	varchar(255) DEFAULT '',
	"credits"	integer DEFAULT '0',
	"is_donor"	integer DEFAULT '0',
	"donor_expire"	timestamp NOT NULL DEFAULT current_timestamp,
	"credits_used"	integer DEFAULT '0',
	"donor_purchases"	integer DEFAULT '0',
	"last_access"	timestamp NOT NULL DEFAULT current_timestamp,
	"last_ip"	varchar(16) DEFAULT '',
	PRIMARY KEY("id" AUTOINCREMENT)
);
CREATE TABLE IF NOT EXISTS "attribute" (
	"user_id"	integer NOT NULL,
	"strength"	integer NOT NULL,
	"agility"	integer NOT NULL,
	"intelligence"	integer NOT NULL,
	"constitution"	integer NOT NULL,
	"charisma"	integer NOT NULL,
	PRIMARY KEY("user_id")
);
CREATE TABLE IF NOT EXISTS "bank_item" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL,
	"item_id"	integer DEFAULT NULL,
	"amount"	integer DEFAULT NULL,
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_bank_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "house_key" (
	"key_obj"	integer NOT NULL,
	"account_id"	integer NOT NULL,
	"assigned_at"	timestamp NOT NULL DEFAULT current_timestamp,
	PRIMARY KEY("key_obj"),
	CONSTRAINT "fk_account" FOREIGN KEY("account_id") REFERENCES "account"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "inventory_item" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL,
	"item_id"	integer DEFAULT NULL,
	"amount"	integer DEFAULT NULL,
	"is_equipped"	integer DEFAULT NULL,
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_inventory_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "object" (
	"number"	integer DEFAULT NULL,
	"name"	varchar(45) DEFAULT NULL
);
CREATE TABLE IF NOT EXISTS "pet" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL,
	"pet_id"	integer DEFAULT NULL,
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_pet_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "punishment" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL DEFAULT '0',
	"reason"	varchar(255) NOT NULL,
	"created_at"	timestamp DEFAULT current_timestamp,
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_punishment_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "quest" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL,
	"quest_id"	integer NOT NULL DEFAULT '0',
	"npcs"	varchar(64) NOT NULL DEFAULT '',
	"npcstarget"	varchar(64) NOT NULL DEFAULT '',
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_quest_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "quest_done" (
	"user_id"	integer NOT NULL,
	"quest_id"	integer NOT NULL,
	PRIMARY KEY("user_id","quest_id"),
	CONSTRAINT "fk_quest_done_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "skillpoint" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL,
	"value"	integer NOT NULL,
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_skillpoint_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "spell" (
	"user_id"	integer NOT NULL,
	"number"	integer NOT NULL,
	"spell_id"	integer DEFAULT NULL,
	PRIMARY KEY("user_id","number"),
	CONSTRAINT "fk_spell_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "statistics" (
	"name"	varchar(50) NOT NULL,
	"value"	varchar(50) DEFAULT NULL,
	PRIMARY KEY("name")
);
CREATE TABLE IF NOT EXISTS "user" (
	"id"	integer NOT NULL,
	"account_id"	integer NOT NULL,
	"deleted"	integer NOT NULL DEFAULT '0',
	"name"	varchar(30) NOT NULL,
	"level"	integer NOT NULL,
	"exp"	integer NOT NULL,
	"genre_id"	integer NOT NULL,
	"race_id"	integer NOT NULL,
	"class_id"	integer NOT NULL,
	"home_id"	integer NOT NULL,
	"description"	varchar(255) DEFAULT NULL,
	"gold"	integer NOT NULL,
	"bank_gold"	integer NOT NULL DEFAULT '0',
	"free_skillpoints"	integer NOT NULL,
	"pets_saved"	integer NOT NULL DEFAULT '0',
	"votes_amount"	integer DEFAULT '0',
	"spouse"	varchar(30) NOT NULL DEFAULT '',
	"message_info"	varchar(512) DEFAULT '',
	"pos_map"	integer NOT NULL,
	"pos_x"	integer NOT NULL,
	"pos_y"	integer NOT NULL,
	"body_id"	integer NOT NULL,
	"head_id"	integer NOT NULL,
	"weapon_id"	integer NOT NULL,
	"helmet_id"	integer NOT NULL,
	"shield_id"	integer NOT NULL,
	"heading"	integer NOT NULL DEFAULT '3',
	"slot_armour"	integer DEFAULT NULL,
	"slot_weapon"	integer DEFAULT NULL,
	"slot_helmet"	integer DEFAULT NULL,
	"slot_shield"	integer DEFAULT NULL,
	"slot_ammo"	integer DEFAULT NULL,
	"slot_ship"	integer DEFAULT NULL,
	"slot_mount"	integer DEFAULT NULL,
	"slot_dm"	integer DEFAULT NULL,
	"slot_rm"	integer DEFAULT NULL,
	"slot_magic"	integer DEFAULT NULL,
	"slot_knuckles"	integer DEFAULT NULL,
	"slot_tool"	integer DEFAULT NULL,
	"min_hp"	integer NOT NULL,
	"max_hp"	integer NOT NULL,
	"min_man"	integer NOT NULL,
	"max_man"	integer NOT NULL,
	"min_sta"	integer NOT NULL,
	"max_sta"	integer NOT NULL,
	"min_ham"	integer NOT NULL,
	"max_ham"	integer NOT NULL,
	"min_sed"	integer NOT NULL,
	"max_sed"	integer NOT NULL,
	"min_hit"	integer NOT NULL,
	"max_hit"	integer NOT NULL,
	"killed_npcs"	integer NOT NULL DEFAULT '0',
	"killed_users"	integer NOT NULL DEFAULT '0',
	"invent_level"	integer NOT NULL DEFAULT '0',
	"is_naked"	integer NOT NULL DEFAULT '0',
	"is_poisoned"	integer NOT NULL DEFAULT '0',
	"is_incinerated"	integer NOT NULL DEFAULT '0',
	"is_dead"	integer NOT NULL DEFAULT '0',
	"is_sailing"	integer NOT NULL DEFAULT '0',
	"is_paralyzed"	integer NOT NULL DEFAULT '0',
	"is_silenced"	integer NOT NULL DEFAULT '0',
	"silence_minutes_left"	integer DEFAULT '0',
	"silence_elapsed_seconds"	integer DEFAULT '0',
	"is_mounted"	integer NOT NULL DEFAULT '0',
	"is_banned"	integer DEFAULT '0',
	"banned_by"	varchar(30) NOT NULL DEFAULT '',
	"ban_reason"	varchar(255) DEFAULT '',
	"counter_pena"	integer NOT NULL DEFAULT '0',
	"deaths"	integer NOT NULL DEFAULT '0',
	"pertenece_consejo_real"	integer NOT NULL DEFAULT '0',
	"pertenece_consejo_caos"	integer NOT NULL DEFAULT '0',
	"pertenece_real"	integer NOT NULL DEFAULT '0',
	"pertenece_caos"	integer NOT NULL DEFAULT '0',
	"ciudadanos_matados"	integer NOT NULL DEFAULT '0',
	"criminales_matados"	integer NOT NULL DEFAULT '0',
	"recibio_armadura_real"	integer NOT NULL DEFAULT '0',
	"recibio_armadura_caos"	integer NOT NULL DEFAULT '0',
	"recibio_exp_real"	integer NOT NULL DEFAULT '0',
	"recibio_exp_caos"	integer NOT NULL DEFAULT '0',
	"recompensas_real"	integer DEFAULT '0',
	"recompensas_caos"	integer DEFAULT '0',
	"reenlistadas"	integer NOT NULL DEFAULT '0',
	"fecha_ingreso"	timestamp NOT NULL DEFAULT current_timestamp,
	"nivel_ingreso"	integer DEFAULT NULL,
	"matados_ingreso"	integer DEFAULT NULL,
	"siguiente_recompensa"	integer DEFAULT NULL,
	"status"	integer DEFAULT '0',
	"guild_index"	integer DEFAULT '0',
	"guild_aspirant_index"	integer DEFAULT NULL,
	"guild_member_history"	varchar(1024) DEFAULT NULL,
	"guild_requests_history"	varchar(1024) DEFAULT NULL,
	"guild_rejected_because"	varchar(255) DEFAULT NULL,
	"chat_global"	integer DEFAULT '1',
	"chat_combate"	integer DEFAULT '1',
	"warnings"	integer NOT NULL DEFAULT '0',
	"elo"	integer NOT NULL DEFAULT '1000',
	"return_map"	integer NOT NULL DEFAULT '0',
	"return_x"	integer NOT NULL DEFAULT '0',
	"return_y"	integer NOT NULL DEFAULT '0',
	"last_logout"	integer NOT NULL DEFAULT 0,
	"is_locked_in_mao"	boolean NOT NULL DEFAULT TRUE,
	"eth_wallet_id"	TEXT,
	UNIQUE("id","account_id","deleted"),
	PRIMARY KEY("id" AUTOINCREMENT),
	CONSTRAINT "fk_user_account" FOREIGN KEY("account_id") REFERENCES "account"("id") ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS "user_deleted" (
	"id"	integer NOT NULL,
	"account_id"	integer NOT NULL,
	"deleted"	timestamp NOT NULL DEFAULT current_timestamp,
	"name"	varchar(30) NOT NULL,
	"level"	integer NOT NULL,
	"exp"	integer NOT NULL,
	"genre_id"	integer NOT NULL,
	"race_id"	integer NOT NULL,
	"class_id"	integer NOT NULL,
	"home_id"	integer NOT NULL,
	"description"	varchar(255) DEFAULT NULL,
	"gold"	integer NOT NULL,
	"bank_gold"	integer NOT NULL DEFAULT '0',
	"free_skillpoints"	integer NOT NULL,
	"pets_saved"	integer NOT NULL DEFAULT '0',
	"votes_amount"	integer DEFAULT '0',
	"spouse"	varchar(30) NOT NULL DEFAULT '',
	"message_info"	varchar(512) DEFAULT '',
	"pos_map"	integer NOT NULL,
	"pos_x"	integer NOT NULL,
	"pos_y"	integer NOT NULL,
	"body_id"	integer NOT NULL,
	"head_id"	integer NOT NULL,
	"weapon_id"	integer NOT NULL,
	"helmet_id"	integer NOT NULL,
	"shield_id"	integer NOT NULL,
	"heading"	integer NOT NULL DEFAULT '3',
	"slot_armour"	integer DEFAULT NULL,
	"slot_weapon"	integer DEFAULT NULL,
	"slot_helmet"	integer DEFAULT NULL,
	"slot_shield"	integer DEFAULT NULL,
	"slot_ammo"	integer DEFAULT NULL,
	"slot_ship"	integer DEFAULT NULL,
	"slot_mount"	integer DEFAULT NULL,
	"slot_dm"	integer DEFAULT NULL,
	"slot_rm"	integer DEFAULT NULL,
	"slot_magic"	integer DEFAULT NULL,
	"slot_knuckles"	integer DEFAULT NULL,
	"slot_tool"	integer DEFAULT NULL,
	"min_hp"	integer NOT NULL,
	"max_hp"	integer NOT NULL,
	"min_man"	integer NOT NULL,
	"max_man"	integer NOT NULL,
	"min_sta"	integer NOT NULL,
	"max_sta"	integer NOT NULL,
	"min_ham"	integer NOT NULL,
	"max_ham"	integer NOT NULL,
	"min_sed"	integer NOT NULL,
	"max_sed"	integer NOT NULL,
	"min_hit"	integer NOT NULL,
	"max_hit"	integer NOT NULL,
	"killed_npcs"	integer NOT NULL DEFAULT '0',
	"killed_users"	integer NOT NULL DEFAULT '0',
	"invent_level"	integer NOT NULL DEFAULT '0',
	"is_naked"	integer NOT NULL DEFAULT '0',
	"is_poisoned"	integer NOT NULL DEFAULT '0',
	"is_incinerated"	integer NOT NULL DEFAULT '0',
	"is_dead"	integer NOT NULL DEFAULT '0',
	"is_sailing"	integer NOT NULL DEFAULT '0',
	"is_paralyzed"	integer NOT NULL DEFAULT '0',
	"is_logged"	integer NOT NULL DEFAULT '0',
	"is_silenced"	integer NOT NULL DEFAULT '0',
	"silence_minutes_left"	integer DEFAULT '0',
	"silence_elapsed_seconds"	integer DEFAULT '0',
	"is_mounted"	integer NOT NULL DEFAULT '0',
	"is_banned"	integer DEFAULT '0',
	"banned_by"	varchar(30) NOT NULL DEFAULT '',
	"ban_reason"	varchar(255) DEFAULT '',
	"counter_pena"	integer NOT NULL DEFAULT '0',
	"deaths"	integer NOT NULL DEFAULT '0',
	"pertenece_consejo_real"	integer NOT NULL DEFAULT '0',
	"pertenece_consejo_caos"	integer NOT NULL DEFAULT '0',
	"pertenece_real"	integer NOT NULL DEFAULT '0',
	"pertenece_caos"	integer NOT NULL DEFAULT '0',
	"ciudadanos_matados"	integer NOT NULL DEFAULT '0',
	"criminales_matados"	integer NOT NULL DEFAULT '0',
	"recibio_armadura_real"	integer NOT NULL DEFAULT '0',
	"recibio_armadura_caos"	integer NOT NULL DEFAULT '0',
	"recibio_exp_real"	integer NOT NULL DEFAULT '0',
	"recibio_exp_caos"	integer NOT NULL DEFAULT '0',
	"recompensas_real"	integer DEFAULT '0',
	"recompensas_caos"	integer DEFAULT '0',
	"reenlistadas"	integer NOT NULL DEFAULT '0',
	"fecha_ingreso"	timestamp NOT NULL DEFAULT current_timestamp,
	"nivel_ingreso"	integer DEFAULT NULL,
	"matados_ingreso"	integer DEFAULT NULL,
	"siguiente_recompensa"	integer DEFAULT NULL,
	"status"	integer DEFAULT '0',
	"guild_index"	integer DEFAULT '0',
	"guild_aspirant_index"	integer DEFAULT NULL,
	"guild_member_history"	varchar(1024) DEFAULT NULL,
	"guild_requests_history"	varchar(1024) DEFAULT NULL,
	"guild_rejected_because"	varchar(255) DEFAULT NULL,
	"chat_global"	integer DEFAULT '1',
	"chat_combate"	integer DEFAULT '1',
	"warnings"	integer NOT NULL DEFAULT '0',
	"elo"	integer NOT NULL DEFAULT '1000',
	"return_map"	integer NOT NULL DEFAULT '0',
	"return_x"	integer NOT NULL DEFAULT '0',
	"return_y"	integer NOT NULL DEFAULT '0'
);
CREATE TABLE IF NOT EXISTS "`_item_on_sale" (
	"id"	INTEGER NOT NULL,
	"created_at"	timestamp NOT NULL,
	"updated_at"	timestamp NOT NULL,
	"deleted_at"	timestamp DEFAULT NULL,
	"item_id"	INTEGER NOT NULL,
	"price_in_tokens"	INTEGER NOT NULL DEFAULT 1,
	"item_quantity"	INTEGER NOT NULL DEFAULT 1,
	PRIMARY KEY("id" AUTOINCREMENT)
);
CREATE TABLE IF NOT EXISTS "mao_item_transaction_status" (
	"user_id" INTEGER NOT NULL,
	"eth_transaction_id"	TEXT,
	"status"	TEXT NOT NULL DEFAULT "PENDING",
	"created_at"	datetime,
	"updated_at"	datetime,
	"deleted_at"	datetime,
	"item_id"	INTEGER NOT NULL,
	"item_quantity"	INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS "mao_user" (
	"id"	INTEGER NOT NULL,
	"user_id"	INTEGER NOT NULL,
	"eth_transaction_id"	TEXT,
	"created_at"	datetime NOT NULL,
	"updated_at"	datetime,
	"deleted_at"	datetime,
	"status"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT)
);
CREATE TABLE IF NOT EXISTS "mao_user_transaction_status" (
	"user_id"	INTEGER NOT NULL,
	"eth_transaction_status"	TEXT,
	"status"	TEXT NOT NULL,
	"created_at"	datetime NOT NULL,
	"updated_at"	datetime NOT NULL,
	"deleted_at"	datetime,
	"account_id"	INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS "idx_user_fk_user_account" ON "user" (
	"account_id"
);
CREATE INDEX IF NOT EXISTS "idx_user_name" ON "user" (
	"name"
);
CREATE INDEX IF NOT EXISTS "idx_house_key_fk_account" ON "house_key" (
	"account_id"
);
CREATE INDEX IF NOT EXISTS "idx_account_email" ON "account" (
	"email"
);
CREATE INDEX IF NOT EXISTS "idx_delete_account" ON "account" (
	"deleted"
);
CREATE VIEW `ranking_users` AS select `u`.`id` AS `id`,`u`.`name` AS `character_name`,`u`.`class_id` AS `class_id`,`u`.`race_id` AS `race_id`,`u`.`genre_id` AS `genre_id`,`u`.`head_id` AS `head_id`,`u`.`level` AS `level`,`u`.`exp` AS `exp`,`u`.`gold` + `u`.`bank_gold` AS `total_gold`,`u`.`criminales_matados` + `u`.`ciudadanos_matados` AS `total_kills` from `user` `u` where `u`.`deleted` <> true;

INSERT INTO "main"."statistics" ("name", "value") VALUES ('online', '1');
INSERT INTO "main"."statistics" ("name", "value") VALUES ('record', '1');

COMMIT;
