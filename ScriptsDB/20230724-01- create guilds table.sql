CREATE TABLE IF NOT EXISTS "guilds" (
	"id" integer NOT NULL,
   "founder_id" integer NOT NULL,
	"guild_name" varchar(32) NOT NULL,
	"creation_date" integer not null default (strftime('%s', 'now')),
	"alignment" integer NOT NULL,
	"last_elections" integer not null default (strftime('%s', 'now')),
	"description" varchar(1024) NOT NULL default '',
	"news" varchar(1024) NOT NULL default '',
	"leader_id" integer NOT NULL DEFAULT 0,
	"level" integer NOT NULL  DEFAULT 1,
	"current_exp" integer NOT NULL  DEFAULT 0,
	"flag_file" integer NOT NULL  DEFAULT 0,
	"url" varchar(128) NOT NULL  DEFAULT '',
    PRIMARY KEY ("id")
);