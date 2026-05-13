DROP TABLE IF EXISTS epic_id_mapping_new;

CREATE TABLE epic_id_mapping_new (
    epic_id TEXT NOT NULL,
    user_id INTEGER NOT NULL,
    last_login INTEGER DEFAULT NULL,
    PRIMARY KEY(epic_id),
    CONSTRAINT fk_epic_id_mapping_user
        FOREIGN KEY (user_id)
        REFERENCES "user"(id)
        ON DELETE CASCADE
        ON UPDATE CASCADE
);

INSERT INTO epic_id_mapping_new (epic_id, user_id, last_login)
SELECT e.epic_id, e.user_id, e.last_login
FROM epic_id_mapping e
INNER JOIN "user" u ON u.id = e.user_id;

DROP TABLE IF EXISTS epic_id_mapping;

ALTER TABLE epic_id_mapping_new RENAME TO epic_id_mapping;

CREATE INDEX IF NOT EXISTS idx_epic_id_mapping_user_id
ON epic_id_mapping(user_id);
