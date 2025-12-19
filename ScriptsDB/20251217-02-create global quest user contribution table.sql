CREATE TABLE IF NOT EXISTS "global_quest_user_contribution" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"event_id" integer NOT NULL,
	"user_id"  integer NOT NULL,
	"timestamp" timestamp NOT NULL DEFAULT current_timestamp,
	"amount" integer NOT NULL,
	FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE ON UPDATE CASCADE,
	FOREIGN KEY (event_id) REFERENCES global_quest_desc(event_id) ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX idx_user_id ON user(id);
CREATE INDEX idx_event_id ON global_quest_desc(event_id);