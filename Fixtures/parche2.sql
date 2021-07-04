CREATE TABLE `account_guard` (
	`account_id` MEDIUMINT(8) UNSIGNED NOT NULL,
	`code` CHAR(5) NULL DEFAULT NULL COLLATE 'utf8mb4_general_ci',
	`created_at` DATETIME NOT NULL DEFAULT current_timestamp(),
	`code_last_sent` DATETIME NOT NULL DEFAULT current_timestamp(),
	`code_resend_attempts` SMALLINT(5) UNSIGNED NOT NULL DEFAULT '0',
	UNIQUE INDEX `fk_account_guard` (`account_id`) USING BTREE,
	CONSTRAINT `fk_account_guard` FOREIGN KEY (`account_id`) REFERENCES `ao_server_test`.`account` (`id`) ON UPDATE CASCADE ON DELETE CASCADE
) COLLATE='utf8mb4_general_ci' ENGINE=InnoDB;