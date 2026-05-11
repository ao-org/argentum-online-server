CREATE TABLE IF NOT EXISTS "bank_item_backup" AS SELECT * FROM "bank_item";
CREATE TABLE IF NOT EXISTS "global_quest_user_contribution_backup" AS SELECT * FROM "global_quest_user_contribution";
CREATE TABLE IF NOT EXISTS "guild_member_history_backup" AS SELECT * FROM "guild_member_history";
CREATE TABLE IF NOT EXISTS "guild_members_backup" AS SELECT * FROM "guild_members";
CREATE TABLE IF NOT EXISTS "guild_request_backup" AS SELECT * FROM "guild_request";
CREATE TABLE IF NOT EXISTS "guild_request_history_backup" AS SELECT * FROM "guild_request_history";
CREATE TABLE IF NOT EXISTS "inventory_item_backup" AS SELECT * FROM "inventory_item";
CREATE TABLE IF NOT EXISTS "inventory_item_skins_backup" AS SELECT * FROM "inventory_item_skins";
CREATE TABLE IF NOT EXISTS "punishment_backup" AS SELECT * FROM "punishment";
CREATE TABLE IF NOT EXISTS "quest_backup" AS SELECT * FROM "quest";
CREATE TABLE IF NOT EXISTS "quest_done_backup" AS SELECT * FROM "quest_done";
CREATE TABLE IF NOT EXISTS "skillpoint_backup" AS SELECT * FROM "skillpoint";
CREATE TABLE IF NOT EXISTS "spell_backup" AS SELECT * FROM "spell";
CREATE TABLE IF NOT EXISTS "epic_id_mapping_backup" AS SELECT * FROM "epic_id_mapping";

CREATE TABLE "bank_item_new" ("user_id" integer NOT NULL, "number" integer NOT NULL, "item_id" integer DEFAULT NULL, "amount" integer DEFAULT NULL, "elemental_tags" integer NOT NULL DEFAULT 0, PRIMARY KEY("user_id","number"), CONSTRAINT "fk_bank_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "bank_item_new" SELECT * FROM "bank_item";
DROP TABLE "bank_item";
ALTER TABLE "bank_item_new" RENAME TO "bank_item";

CREATE TABLE "global_quest_user_contribution_new" ("id" INTEGER, "event_id" integer NOT NULL, "user_id" integer NOT NULL, "timestamp" timestamp NOT NULL DEFAULT current_timestamp, "amount" integer NOT NULL, PRIMARY KEY("id" AUTOINCREMENT), FOREIGN KEY("event_id") REFERENCES "global_quest_desc"("event_id") ON DELETE CASCADE ON UPDATE CASCADE, FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "global_quest_user_contribution_new" SELECT * FROM "global_quest_user_contribution";
DROP TABLE "global_quest_user_contribution";
ALTER TABLE "global_quest_user_contribution_new" RENAME TO "global_quest_user_contribution";

CREATE TABLE "guild_member_history_new" ("id" integer NOT NULL, "user_id" integer NOT NULL, "guild_name" varchar(20) NOT NULL, "request_time" integer NOT NULL DEFAULT (strftime('%s', 'now')), PRIMARY KEY("id" AUTOINCREMENT), CONSTRAINT "fk_guild_member_history" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "guild_member_history_new" SELECT * FROM "guild_member_history";
DROP TABLE "guild_member_history";
ALTER TABLE "guild_member_history_new" RENAME TO "guild_member_history";

CREATE TABLE "guild_members_new" ("id" integer NOT NULL, "guild_id" integer NOT NULL, "user_id" integer NOT NULL, PRIMARY KEY("id" AUTOINCREMENT), CONSTRAINT "fk_guild_members" FOREIGN KEY("guild_id") REFERENCES "guilds"("id") ON DELETE CASCADE ON UPDATE CASCADE, CONSTRAINT "fk_user_guild_members" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "guild_members_new" SELECT * FROM "guild_members";
DROP TABLE "guild_members";
ALTER TABLE "guild_members_new" RENAME TO "guild_members";

CREATE TABLE "guild_request_new" ("id" integer NOT NULL, "guild_id" integer NOT NULL, "user_id" integer NOT NULL, "description" varchar(512) NOT NULL, PRIMARY KEY("id" AUTOINCREMENT), CONSTRAINT "fk_guild_request_guild" FOREIGN KEY("guild_id") REFERENCES "guilds"("id") ON DELETE CASCADE ON UPDATE CASCADE, CONSTRAINT "fk_guild_request_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "guild_request_new" SELECT * FROM "guild_request";
DROP TABLE "guild_request";
ALTER TABLE "guild_request_new" RENAME TO "guild_request";

CREATE TABLE "guild_request_history_new" ("id" integer NOT NULL, "user_id" integer NOT NULL, "guild_name" varchar(20) NOT NULL, "request_time" integer NOT NULL DEFAULT (strftime('%s', 'now')), PRIMARY KEY("id" AUTOINCREMENT), CONSTRAINT "fk_guild_request_history" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "guild_request_history_new" SELECT * FROM "guild_request_history";
DROP TABLE "guild_request_history";
ALTER TABLE "guild_request_history_new" RENAME TO "guild_request_history";

CREATE TABLE "inventory_item_new" ("user_id" integer NOT NULL, "number" integer NOT NULL, "item_id" integer DEFAULT NULL, "amount" integer DEFAULT NULL, "is_equipped" integer DEFAULT NULL, "elemental_tags" integer NOT NULL DEFAULT 0, PRIMARY KEY("user_id","number"), CONSTRAINT "fk_inventory_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "inventory_item_new" SELECT * FROM "inventory_item";
DROP TABLE "inventory_item";
ALTER TABLE "inventory_item_new" RENAME TO "inventory_item";

CREATE TABLE "inventory_item_skins_new" ("id" INTEGER, "user_id" INTEGER NOT NULL, "type_skin" INTEGER NOT NULL, "skin_id" INTEGER NOT NULL, "skin_equipped" TINYINT NOT NULL, "date_created" TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, PRIMARY KEY("id" AUTOINCREMENT), FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "inventory_item_skins_new" SELECT * FROM "inventory_item_skins";
DROP TABLE "inventory_item_skins";
ALTER TABLE "inventory_item_skins_new" RENAME TO "inventory_item_skins";

CREATE TABLE "punishment_new" ("user_id" integer NOT NULL, "number" integer NOT NULL DEFAULT '0', "reason" varchar(255) NOT NULL, "created_at" timestamp DEFAULT current_timestamp, PRIMARY KEY("user_id","number"), CONSTRAINT "fk_punishment_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "punishment_new" SELECT * FROM "punishment";
DROP TABLE "punishment";
ALTER TABLE "punishment_new" RENAME TO "punishment";

CREATE TABLE "quest_new" ("user_id" integer NOT NULL, "number" integer NOT NULL, "quest_id" integer NOT NULL DEFAULT '0', "npcs" varchar(64) NOT NULL DEFAULT '', "npcstarget" varchar(64) NOT NULL DEFAULT '', PRIMARY KEY("user_id","number"), CONSTRAINT "fk_quest_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "quest_new" SELECT * FROM "quest";
DROP TABLE "quest";
ALTER TABLE "quest_new" RENAME TO "quest";

CREATE TABLE "quest_done_new" ("user_id" integer NOT NULL, "quest_id" integer NOT NULL, PRIMARY KEY("user_id","quest_id"), CONSTRAINT "fk_quest_done_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "quest_done_new" SELECT * FROM "quest_done";
DROP TABLE "quest_done";
ALTER TABLE "quest_done_new" RENAME TO "quest_done";

CREATE TABLE "skillpoint_new" ("user_id" integer NOT NULL, "number" integer NOT NULL, "value" integer NOT NULL, PRIMARY KEY("user_id","number"), CONSTRAINT "fk_skillpoint_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "skillpoint_new" SELECT * FROM "skillpoint";
DROP TABLE "skillpoint";
ALTER TABLE "skillpoint_new" RENAME TO "skillpoint";

CREATE TABLE "spell_new" ("user_id" integer NOT NULL, "number" integer NOT NULL, "spell_id" integer DEFAULT NULL, PRIMARY KEY("user_id","number"), CONSTRAINT "fk_spell_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "spell_new" SELECT * FROM "spell";
DROP TABLE "spell";
ALTER TABLE "spell_new" RENAME TO "spell";

CREATE TABLE "epic_id_mapping_new" ("epic_id" varchar(64) NOT NULL, "user_id" integer NOT NULL, "last_login" integer NOT NULL, CONSTRAINT "unique_epic_user_id" UNIQUE("epic_id","user_id"), CONSTRAINT "fk_epic_id_mapping" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "epic_id_mapping_new" SELECT * FROM "epic_id_mapping";
DROP TABLE "epic_id_mapping";
ALTER TABLE "epic_id_mapping_new" RENAME TO "epic_id_mapping";

CREATE INDEX IF NOT EXISTS "idx_pk_account_id" ON "account" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_account_collectible_cards_id" ON "account_collectible_cards" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_all_housekeys_id" ON "all_housekeys" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_char_transfer_record_id" ON "char_transfer_record" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_collectible_cards_id" ON "collectible_cards" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_global_quest_desc_id" ON "global_quest_desc" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_global_quest_user_contribution_id" ON "global_quest_user_contribution" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_guild_member_history_id" ON "guild_member_history" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_guild_members_id" ON "guild_members" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_guild_request_id" ON "guild_request" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_guild_request_history_id" ON "guild_request_history" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_guilds_id" ON "guilds" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_inventory_item_skins_id" ON "inventory_item_skins" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_migrations_id" ON "migrations" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_patreon_shop_audit_id" ON "patreon_shop_audit" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_tokens_id" ON "tokens" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_user_id" ON "user" ("id");
CREATE INDEX IF NOT EXISTS "idx_pk_bank_item" ON "bank_item" ("user_id", "number");
CREATE INDEX IF NOT EXISTS "idx_pk_inventory_item" ON "inventory_item" ("user_id", "number");
CREATE INDEX IF NOT EXISTS "idx_pk_punishment" ON "punishment" ("user_id", "number");
CREATE INDEX IF NOT EXISTS "idx_pk_quest" ON "quest" ("user_id", "number");
CREATE INDEX IF NOT EXISTS "idx_pk_quest_done" ON "quest_done" ("user_id", "quest_id");
CREATE INDEX IF NOT EXISTS "idx_pk_skillpoint" ON "skillpoint" ("user_id", "number");
CREATE INDEX IF NOT EXISTS "idx_pk_spell" ON "spell" ("user_id", "number");
CREATE INDEX IF NOT EXISTS "idx_fk_bank_item_user_id" ON "bank_item" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_global_quest_user_contribution_user_id" ON "global_quest_user_contribution" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_global_quest_user_contribution_event_id" ON "global_quest_user_contribution" ("event_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_member_history_user_id" ON "guild_member_history" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_members_user_id" ON "guild_members" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_members_guild_id" ON "guild_members" ("guild_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_request_user_id" ON "guild_request" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_request_guild_id" ON "guild_request" ("guild_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_request_history_user_id" ON "guild_request_history" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_inventory_item_user_id" ON "inventory_item" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_punishment_user_id" ON "punishment" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_quest_user_id" ON "quest" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_quest_done_user_id" ON "quest_done" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_skillpoint_user_id" ON "skillpoint" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_spell_user_id" ON "spell" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_epic_id_mapping_user_id" ON "epic_id_mapping" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_user_account_id" ON "user" ("account_id");
CREATE UNIQUE INDEX IF NOT EXISTS "uq_inventory_item_skins_user_skin" ON "inventory_item_skins" ("user_id", "skin_id");
CREATE INDEX IF NOT EXISTS "user_index" ON "user" ("id", "account_id", "deleted");