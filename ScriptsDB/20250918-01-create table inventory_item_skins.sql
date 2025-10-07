DROP TABLE IF EXISTS inventory_item_skins;

CREATE TABLE inventory_item_skins (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    type_skin INTEGER NOT NULL,
    skin_id INTEGER NOT NULL,
    skin_equipped TINYINT NOT NULL,
    date_created TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX idx_inventory_item_skins_user_id ON inventory_item_skins(user_id);