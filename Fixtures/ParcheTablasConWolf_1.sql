ALTER TABLE `user`
	DROP COLUMN `battle_points`,
	DROP COLUMN `last_map`,
	DROP COLUMN `items_amount`,
	DROP COLUMN `is_hidden`,
	DROP COLUMN `is_hungry`,
	DROP COLUMN `is_thirsty`;

ALTER TABLE `attribute` RENAME TO `attribute_old`;

CREATE TABLE `attribute` (
	`user_id` MEDIUMINT(8) UNSIGNED NOT NULL,
	`strength` TINYINT(4) UNSIGNED NOT NULL,
	`agility` TINYINT(4) UNSIGNED NOT NULL,
	`intelligence` TINYINT(4) UNSIGNED NOT NULL,
	`constitution` TINYINT(4) UNSIGNED NOT NULL,
	`charisma` TINYINT(4) UNSIGNED NOT NULL
)
COLLATE='utf8mb4_general_ci'
ENGINE=InnoDB
;

ALTER TABLE `attribute`` ADD PRIMARY KEY(`user_id`);
