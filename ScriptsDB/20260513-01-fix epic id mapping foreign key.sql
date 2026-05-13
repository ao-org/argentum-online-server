DROP TABLE IF EXISTS epic_id_mapping_new;

CREATE TABLE epic_id_mapping_new (epic_id varchar(64) NOT NULL, user_id integer NOT NULL, last_login integer NOT NULL, CONSTRAINT fk_epic_id_mapping FOREIGN KEY (user_id) REFERENCES "user"(id) ON DELETE CASCADE ON UPDATE CASCADE, CONSTRAINT unique_epic_user_id UNIQUE (epic_id, user_id));

INSERT INTO epic_id_mapping_new (epic_id, user_id, last_login) SELECT epic_id, user_id, last_login FROM epic_id_mapping WHERE user_id IN (SELECT id FROM "user");

DROP TABLE IF EXISTS epic_id_mapping;

ALTER TABLE epic_id_mapping_new RENAME TO epic_id_mapping;