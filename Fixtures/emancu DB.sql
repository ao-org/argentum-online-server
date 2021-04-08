CREATE TABLE `guilds` (
	`id` MEDIUMINT(8) UNSIGNED NOT NULL AUTO_INCREMENT,
	`founder_id` MEDIUMINT(8) UNSIGNED NOT NULL,
	`leader_id` MEDIUMINT(8) UNSIGNED NOT NULL,
	`name` VARCHAR(255) NOT NULL COLLATE 'utf8mb4_general_ci',
	`description` VARCHAR(255) NOT NULL DEFAULT '""' COLLATE 'utf8mb4_general_ci',
	`level` TINYINT(3) UNSIGNED NOT NULL DEFAULT '1',
	`experience` INT(10) UNSIGNED NOT NULL DEFAULT '0',
	`alignment` TINYINT(3) UNSIGNED NOT NULL DEFAULT '0',
	`created_at` TIMESTAMP NOT NULL DEFAULT current_timestamp(),
	`updated_at` TIMESTAMP NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
	`deleted_at` TIMESTAMP NULL DEFAULT NULL,
	PRIMARY KEY (`id`) USING BTREE,
	UNIQUE INDEX `guild_name_unique_index` (`name`, `deleted_at`) USING BTREE,
	INDEX `fk_founder_id` (`founder_id`) USING BTREE,
	INDEX `fk_leader_id` (`leader_id`) USING BTREE,
	CONSTRAINT `fk_founder_id` FOREIGN KEY (`founder_id`) REFERENCES `ao_server_prod`.`user` (`id`) ON UPDATE RESTRICT ON DELETE RESTRICT,
	CONSTRAINT `fk_leader_id` FOREIGN KEY (`leader_id`) REFERENCES `ao_server_prod`.`user` (`id`) ON UPDATE RESTRICT ON DELETE RESTRICT
)
COLLATE='utf8mb4_general_ci'
ENGINE=InnoDB
;



CREATE TABLE `guild_news` (
	`id` MEDIUMINT UNSIGNED NOT NULL AUTO_INCREMENT,
	`guild_id` MEDIUMINT UNSIGNED NOT NULL,
	`author_id` MEDIUMINT UNSIGNED NOT NULL,
	`message` TEXT NOT NULL,
	`created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP(),
	`updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP() ON UPDATE CURRENT_TIMESTAMP(),
	`deleted_at` TIMESTAMP NULL DEFAULT NULL,
	PRIMARY KEY (`id`),
	CONSTRAINT `guild_news__guild_id_fk` FOREIGN KEY (`guild_id`) REFERENCES `guilds` (`id`) ON DELETE CASCADE,
	CONSTRAINT `guild_news__author_id_fk` FOREIGN KEY (`author_id`) REFERENCES `user` (`id`) ON DELETE CASCADE
)
COLLATE='utf8mb4_general_ci'
;

CREATE TABLE `guild_memberships` (
	`id` MEDIUMINT(8) UNSIGNED NOT NULL AUTO_INCREMENT,
	`guild_id` MEDIUMINT(8) UNSIGNED NOT NULL,
	`user_id` MEDIUMINT(8) UNSIGNED NOT NULL,
	`state` VARCHAR(50) NOT NULL COLLATE 'utf8mb4_general_ci',
	`petition_message` VARCHAR(255) NOT NULL COLLATE 'utf8mb4_general_ci',
	`state_explanation` VARCHAR(511) NULL DEFAULT NULL COLLATE 'utf8mb4_general_ci',
	`created_at` TIMESTAMP NOT NULL DEFAULT current_timestamp(),
	`updated_at` TIMESTAMP NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
	`deleted_at` TIMESTAMP NULL DEFAULT NULL,
	PRIMARY KEY (`id`) USING BTREE,
	INDEX `guild_memberships__guild_id__fk` (`guild_id`) USING BTREE,
	INDEX `guild_memberships__user_id__fk` (`user_id`) USING BTREE,
	CONSTRAINT `guild_memberships__guild_id__fk` FOREIGN KEY (`guild_id`) REFERENCES `ao_server_prod`.`guilds` (`id`) ON UPDATE RESTRICT ON DELETE CASCADE,
	CONSTRAINT `guild_memberships__user_id__fk` FOREIGN KEY (`user_id`) REFERENCES `ao_server_prod`.`user` (`id`) ON UPDATE RESTRICT ON DELETE RESTRICT
)
COLLATE='utf8mb4_general_ci'
ENGINE=InnoDB
;
