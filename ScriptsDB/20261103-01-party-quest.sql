ALTER TABLE user ADD COLUMN party_quest_save_revision INTEGER NOT NULL DEFAULT 0;
-- Shared party-quest state is authoritative. Ordinary quest slots are projections.
CREATE TABLE party_quest_run (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    revision INTEGER NOT NULL CHECK(revision>0),
    quest_id INTEGER NOT NULL,
    leader_id INTEGER NOT NULL,
    target_npc INTEGER NOT NULL,
    required INTEGER NOT NULL CHECK(required BETWEEN 1 AND 32767),
    progress INTEGER NOT NULL DEFAULT 0 CHECK(progress >= 0 AND progress <= required),
    reward_gold INTEGER NOT NULL CHECK(reward_gold > 0 AND reward_gold <= 90000000)
);
CREATE TABLE party_quest_member (
    run_id INTEGER NOT NULL REFERENCES party_quest_run(id),
    ordinal INTEGER NOT NULL CHECK(ordinal BETWEEN 1 AND 6),
    character_id INTEGER NOT NULL,
    user_id INTEGER REFERENCES user(id) ON DELETE SET NULL,
    name TEXT NOT NULL,
    participated INTEGER NOT NULL DEFAULT 0 CHECK(participated IN (0,1)),
    status INTEGER NOT NULL DEFAULT 0 CHECK(status IN (0,1,2)),
    PRIMARY KEY(run_id, character_id),
    UNIQUE(run_id, ordinal)
);
CREATE UNIQUE INDEX party_quest_one_active ON party_quest_member(character_id) WHERE status=0;
CREATE INDEX party_quest_character ON party_quest_member(character_id,status);
CREATE TABLE party_quest_kill (
    run_id INTEGER NOT NULL REFERENCES party_quest_run(id),
    ordinal INTEGER NOT NULL,
    qualified_mask INTEGER NOT NULL CHECK(qualified_mask BETWEEN 1 AND 63),
    PRIMARY KEY(run_id, ordinal)
);
