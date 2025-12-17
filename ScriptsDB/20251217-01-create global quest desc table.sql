CREATE TABLE IF NOT EXISTS "global_quest_desc" (
	"id" INTEGER PRIMARY KEY,
	"name"  varchar(50) NOT NULL,
	"obj_id" integer NOT NULL,
	"counter" integer NOT NULL,
	"start_date" timestamp NOT NULL DEFAULT current_timestamp,
	"end_date" timestamp DEFAULT NULL
);