CREATE TABLE IF NOT EXISTS "castle_coordinates" (
	"id" INTEGER PRIMARY KEY AUTOINCREMENT,
	"castle_id" integer NOT NULL UNIQUE,
	"outside_map" integer DEFAULT 0,
	"outside_x" integer DEFAULT 0,
	"outside_y" integer DEFAULT 0,
	"inside_map" integer NOT NULL,
	"inside_x" integer NOT NULL,
	"inside_y" integer NOT NULL,
	FOREIGN KEY (castle_id) REFERENCES castle(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX idx_castle_coordinates_castle_id ON castle_coordinates(castle_id);