CREATE TABLE IF NOT EXISTS "global_quest_desc" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"event_id" integer NOT NULL UNIQUE,
	"name"  varchar(50) NOT NULL,
	"obj_id" integer NOT NULL,
	"threshold" integer NOT NULL,
	"start_date" timestamp NOT NULL DEFAULT current_timestamp,
	"end_date" timestamp DEFAULT NULL,
	"is_active" boolean NOT NULL DEFAULT 1
);