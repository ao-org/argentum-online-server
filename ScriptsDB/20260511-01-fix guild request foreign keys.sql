DROP TABLE IF EXISTS guild_request_new;

CREATE TABLE guild_request_new (
    id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
    guild_id INTEGER NOT NULL,
    user_id INTEGER NOT NULL,
    description VARCHAR(512) NOT NULL,
    CONSTRAINT fk_guild_request_guild_id
        FOREIGN KEY (guild_id)
        REFERENCES guilds(id)
        ON DELETE CASCADE
        ON UPDATE CASCADE,
    CONSTRAINT fk_guild_request_user_id
        FOREIGN KEY (user_id)
        REFERENCES "user"(id)
        ON DELETE CASCADE
        ON UPDATE CASCADE
);

INSERT INTO guild_request_new (id, guild_id, user_id, description)
SELECT gr.id, gr.guild_id, gr.user_id, gr.description
    FROM guild_request gr
    INNER JOIN "user" u ON u.id = gr.user_id
    INNER JOIN guilds g ON g.id = gr.guild_id;

DROP TABLE IF EXISTS guild_request;

ALTER TABLE guild_request_new RENAME TO guild_request;

CREATE INDEX IF NOT EXISTS idx_fk_guild_request_user_id ON guild_request(user_id);

CREATE INDEX IF NOT EXISTS idx_fk_guild_request_guild_id ON guild_request(guild_id);
