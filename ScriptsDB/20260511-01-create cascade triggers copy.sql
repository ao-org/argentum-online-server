CREATE TABLE IF NOT EXISTS "guild_request_backup" AS SELECT * FROM "guild_request";

CREATE TABLE "guild_request_new" ("id" integer NOT NULL, "guild_id" integer NOT NULL, "user_id" integer NOT NULL, "description" varchar(512) NOT NULL, PRIMARY KEY("id" AUTOINCREMENT), CONSTRAINT "fk_guild_request_guild" FOREIGN KEY("guild_id") REFERENCES "guilds"("id") ON DELETE CASCADE ON UPDATE CASCADE, CONSTRAINT "fk_guild_request_user" FOREIGN KEY("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO "guild_request_new" SELECT * FROM "guild_request";
DROP TABLE "guild_request";
ALTER TABLE "guild_request_new" RENAME TO "guild_request";

CREATE INDEX IF NOT EXISTS "idx_fk_guild_request_user_id" ON "guild_request" ("user_id");
CREATE INDEX IF NOT EXISTS "idx_fk_guild_request_guild_id" ON "guild_request" ("guild_id");