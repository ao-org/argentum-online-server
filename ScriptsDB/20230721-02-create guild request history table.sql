CREATE TABLE IF NOT EXISTS "guild_request_history" (
	"id" integer NOT NULL,
    "user_id" integer NOT NULL,
    "guild_name" varchar(20) NOT NULL,
    "request_time" integer not null default (strftime('%s','now')),
    PRIMARY KEY ("id"),
    CONSTRAINT "fk_guild_request_history" FOREIGN KEY ("user_id") REFERENCES "user" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);