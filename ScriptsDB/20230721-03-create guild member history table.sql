CREATE TABLE IF NOT EXISTS "guild_member_history" (
    "user_id" integer NOT NULL,
    "guild_name" varchar(20) NOT NULL,
    "request_time" integer not null default (strftime('%s','now')),
    PRIMARY KEY ("user_id"),
    CONSTRAINT "fk_guild_member_history" FOREIGN KEY ("user_id") REFERENCES "user" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);