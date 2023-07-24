CREATE TABLE IF NOT EXISTS "guild_members" (
	"id" integer NOT NULL,
	"guild_id"	integer NOT NULL,
    "user_id" integer NOT NULL,
	CONSTRAINT "fk_guild_members" FOREIGN KEY("guild_id") REFERENCES "guilds"("id") ON DELETE CASCADE ON UPDATE CASCADE,
	CONSTRAINT "fk_user_guild_members" FOREIGN KEY("user_id") REFERENCES "user" ("id") ON DELETE CASCADE ON UPDATE CASCADE
	PRIMARY KEY ("id")
);