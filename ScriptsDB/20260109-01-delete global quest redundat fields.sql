BEGIN TRANSACTION;

ALTER TABLE global_quest_desc DROP COLUMN threshold;
ALTER TABLE global_quest_desc DROP COLUMN start_date;
ALTER TABLE global_quest_desc DROP COLUMN end_date;

COMMIT;