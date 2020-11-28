DROP TABLE IF EXISTS `house_key`;
CREATE TABLE `house_key` (
  `key_obj` SMALLINT UNSIGNED NOT NULL PRIMARY KEY,
  `account_id` MEDIUMINT UNSIGNED NOT NULL,
  `assigned_at` TIMESTAMP NOT NULL DEFAULT current_timestamp(),
  
  CONSTRAINT fk_account FOREIGN KEY (account_id) REFERENCES account(id)

) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;